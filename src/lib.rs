//! Read spreadsheets into Python lists, with the row and column limits checked while reading.
//!
//! calamine builds a sheet's grid in one allocation sized by its farthest cell (`Range::from_sparse`), and a failed
//! allocation aborts the process: one value in XFD1048576 of a valid xlsx asks for 512 GB
//! (tafia/calamine#693). For xlsx, xlsm and xlsb this crate reads the cells one by one instead
//! (`worksheet_cells_reader`), stops with a Python exception as soon as a cell lies past the limits, and builds the
//! rows only then. xls and ods are densified by calamine while the workbook is opened; for them the limits are
//! checked afterwards, so a caller reading untrusted .xls / .ods files should still do it in a separate process.
//!
//! Cell values are the Python values python-calamine's `to_python()` gives (the conversion below follows its
//! `CellValue`, MIT, https://github.com/dimastbk/python-calamine), except that an empty cell is `None`.

use std::io::Cursor;

use calamine::{
    Data, DataRef, DataType, Dimensions, Reader, SheetType, SheetVisible, Sheets, open_workbook_auto_from_rs,
};
use chrono::Datelike;
use pyo3::create_exception;
use pyo3::exceptions::PyValueError;
use pyo3::prelude::*;
use pyo3::types::PyList;

create_exception!(sheet_reader, ReadError, PyValueError, "The workbook cannot be read.");
create_exception!(
    sheet_reader,
    LimitExceeded,
    ReadError,
    "A sheet reaches past the row or column limit."
);

const MIN_YEAR: i32 = 1; // Python's datetime.MINYEAR
const MAX_YEAR: i32 = 9999; // datetime.MAXYEAR

#[derive(Debug, Clone)]
enum Value {
    Int(i64),
    Float(f64),
    String(String),
    Time(chrono::NaiveTime),
    Date(chrono::NaiveDate),
    DateTime(chrono::NaiveDateTime),
    Timedelta(chrono::Duration),
    Bool(bool),
    Empty,
}

fn in_year_range<D: Datelike>(value: D) -> Option<D> {
    (MIN_YEAR..=MAX_YEAR).contains(&value.year()).then_some(value)
}

impl<T: DataType> From<&T> for Value {
    fn from(value: &T) -> Self {
        if value.is_int() {
            value.get_int().map(Value::Int).unwrap_or(Value::Empty)
        } else if value.is_float() {
            value.get_float().map(Value::Float).unwrap_or(Value::Empty)
        } else if value.is_string() {
            value
                .get_string()
                .map(|s| Value::String(s.to_owned()))
                .unwrap_or(Value::Empty)
        } else if value.is_datetime() {
            let dt = value.get_datetime().unwrap();
            let v = dt.as_f64();
            if dt.is_duration() {
                value.as_duration().map(Value::Timedelta)
            } else if v < 1.0 {
                value.as_time().map(Value::Time)
            } else if v == (v as u64) as f64 {
                value.as_date().and_then(in_year_range).map(Value::Date)
            } else {
                value.as_datetime().and_then(in_year_range).map(Value::DateTime)
            }
            .unwrap_or(Value::Float(v))
        } else if value.is_datetime_iso() {
            let v = value.get_datetime_iso().unwrap();
            if v.contains('T') {
                value.as_datetime().and_then(in_year_range).map(Value::DateTime)
            } else if v.contains(':') {
                value.as_time().map(Value::Time)
            } else {
                value.as_date().and_then(in_year_range).map(Value::Date)
            }
            .unwrap_or(Value::String(v.to_owned()))
        } else if value.is_duration_iso() {
            value.as_time().map(Value::Time).unwrap_or(
                value
                    .get_duration_iso()
                    .map(|s| Value::String(s.to_owned()))
                    .unwrap_or(Value::Empty),
            )
        } else if value.is_bool() {
            value.get_bool().map(Value::Bool).unwrap_or(Value::Empty)
        } else {
            Value::Empty
        }
    }
}

impl<'py> IntoPyObject<'py> for Value {
    type Target = PyAny;
    type Output = Bound<'py, PyAny>;
    type Error = PyErr;

    fn into_pyobject(self, py: Python<'py>) -> Result<Self::Output, Self::Error> {
        Ok(match self {
            Value::Int(v) => v.into_pyobject(py)?.into_any(),
            Value::Float(v) => v.into_pyobject(py)?.into_any(),
            Value::String(v) => v.into_pyobject(py)?.into_any(),
            Value::Bool(v) => v.into_pyobject(py)?.to_owned().into_any(),
            Value::Time(v) => v.into_pyobject(py)?.into_any(),
            Value::Date(v) => v.into_pyobject(py)?.into_any(),
            Value::DateTime(v) => v.into_pyobject(py)?.into_any(),
            Value::Timedelta(v) => v.into_pyobject(py)?.into_any(),
            Value::Empty => py.None().into_bound(py),
        })
    }
}

/// A sheet as read: its cells (row, column, value), the farthest row and column, merged ranges.
struct SheetData {
    name: String,
    visible: bool,
    cells: Vec<(u32, u32, Value)>,
    height: u32, // rows from the first one to the last used (0 when empty)
    width: u32,
    merged: Vec<((u32, u32), (u32, u32))>,
}

enum Failure {
    Read(String),
    Limit(String),
}

/// Collects the non-empty cells of a sheet, refusing as soon as one lies past the limits.
struct Collector<'a> {
    name: &'a str,
    max_rows: u32,
    max_cols: u32,
    cells: Vec<(u32, u32, Value)>,
    height: u32,
    width: u32,
}

impl<'a> Collector<'a> {
    fn new(name: &'a str, max_rows: u32, max_cols: u32) -> Self {
        Collector {
            name,
            max_rows,
            max_cols,
            cells: Vec::new(),
            height: 0,
            width: 0,
        }
    }

    fn add<T: DataType>(&mut self, (row, col): (u32, u32), value: &T) -> Result<(), Failure> {
        if value.is_empty() {
            return Ok(());
        }
        if row >= self.max_rows || col >= self.max_cols {
            return Err(self.past(row, col));
        }
        self.height = self.height.max(row + 1);
        self.width = self.width.max(col + 1);
        self.cells.push((row, col, Value::from(value)));
        Ok(())
    }

    fn past(&self, row: u32, col: u32) -> Failure {
        Failure::Limit(format!(
            "sheet '{}' reaches past {} rows or {} columns (a cell at row {}, column {})",
            self.name,
            self.max_rows,
            self.max_cols,
            row as u64 + 1,
            col as u64 + 1
        ))
    }
}

fn merged(dims: Vec<Dimensions>) -> Vec<((u32, u32), (u32, u32))> {
    dims.into_iter().map(|d| (d.start, d.end)).collect()
}

fn read_sheet(
    sheets: &mut Sheets<Cursor<Vec<u8>>>,
    name: &str,
    visible: bool,
    max_rows: u32,
    max_cols: u32,
) -> Result<SheetData, Failure> {
    let mut collect = Collector::new(name, max_rows, max_cols);
    let read = |e: &dyn std::fmt::Display| Failure::Read(format!("sheet '{name}': {e}"));
    let merges = match sheets {
        Sheets::Xlsx(xlsx) => {
            {
                let mut reader = xlsx.worksheet_cells_reader(name).map_err(|e| read(&e))?;
                while let Some(cell) = reader.next_cell().map_err(|e| read(&e))? {
                    let value: &DataRef = cell.get_value();
                    collect.add(cell.get_position(), value)?;
                }
            }
            merged(xlsx.merge_cells_by_sheet_name(name).map_err(|e| read(&e))?)
        }
        Sheets::Xlsb(xlsb) => {
            let mut reader = xlsb.worksheet_cells_reader(name).map_err(|e| read(&e))?;
            while let Some(cell) = reader.next_cell().map_err(|e| read(&e))? {
                let value: &DataRef = cell.get_value();
                collect.add(cell.get_position(), value)?;
            }
            Vec::new()
        }
        Sheets::Xls(xls) => {
            let range = xls.worksheet_range(name).map_err(|e| read(&e))?;
            dense(&range, &mut collect)?;
            merged(xls.merge_cells_by_sheet_name(name).map_err(|e| read(&e))?)
        }
        Sheets::Ods(ods) => {
            let range = ods.worksheet_range(name).map_err(|e| read(&e))?;
            dense(&range, &mut collect)?;
            Vec::new()
        }
    };
    Ok(SheetData {
        name: name.to_owned(),
        visible,
        cells: collect.cells,
        height: collect.height,
        width: collect.width,
        merged: merges,
    })
}

/// A range calamine has already built: its extent is checked before its cells are taken.
fn dense(range: &calamine::Range<Data>, collect: &mut Collector) -> Result<(), Failure> {
    if let Some((row, col)) = range.end()
        && (row >= collect.max_rows || col >= collect.max_cols)
    {
        return Err(collect.past(row, col));
    }
    let (start_row, start_col) = range.start().unwrap_or((0, 0));
    for (row, col, value) in range.used_cells() {
        collect.add((start_row + row as u32, start_col + col as u32), value)?;
    }
    Ok(())
}

fn read_workbook(data: Vec<u8>, max_rows: u32, max_cols: u32, hidden: bool) -> Result<Vec<SheetData>, Failure> {
    let mut sheets = open_workbook_auto_from_rs(Cursor::new(data)).map_err(|e| Failure::Read(e.to_string()))?;
    let wanted: Vec<(String, bool)> = sheets
        .sheets_metadata()
        .iter()
        .filter(|m| m.typ == SheetType::WorkSheet)
        .map(|m| (m.name.clone(), m.visible == SheetVisible::Visible))
        .filter(|(_, visible)| hidden || *visible)
        .collect();
    wanted
        .into_iter()
        .map(|(name, visible)| read_sheet(&mut sheets, &name, visible, max_rows, max_cols))
        .collect()
}

/// Rows of a sheet as Python lists: from the first row and column to the farthest used ones, None where empty.
fn rows<'py>(py: Python<'py>, sheet: SheetData) -> PyResult<Bound<'py, PyList>> {
    let width = sheet.width as usize;
    let mut grid: Vec<Vec<Option<Value>>> = Vec::with_capacity(sheet.height as usize);
    grid.resize_with(sheet.height as usize, || vec![None; width]);
    for (row, col, value) in sheet.cells {
        grid[row as usize][col as usize] = Some(value);
    }
    let out = PyList::empty(py);
    for row in grid {
        out.append(PyList::new(py, row.into_iter().map(|v| v.unwrap_or(Value::Empty)))?)?;
    }
    Ok(out)
}

/// Read every worksheet of a workbook (xlsx, xlsm, xlsb, xls, ods) from its bytes.
///
/// Returns a list of (name, visible, rows, merged) with rows a list of lists of cell values and merged a list of
/// ((first row, first column), (last row, last column)), 0-based. Hidden sheets are left out unless ``hidden``.
/// Raises LimitExceeded when a sheet has a cell at or past row ``max_rows`` or column ``max_cols`` (1-based: more
/// than that many rows or columns), ReadError when the file cannot be read.
#[pyfunction]
#[pyo3(signature = (data, *, max_rows, max_cols, hidden = false))]
fn read<'py>(
    py: Python<'py>,
    data: Vec<u8>,
    max_rows: u32,
    max_cols: u32,
    hidden: bool,
) -> PyResult<Bound<'py, PyList>> {
    let sheets = py
        .detach(|| read_workbook(data, max_rows, max_cols, hidden))
        .map_err(|f| match f {
            Failure::Read(m) => ReadError::new_err(m),
            Failure::Limit(m) => LimitExceeded::new_err(m),
        })?;
    let out = PyList::empty(py);
    for sheet in sheets {
        let (name, visible, merged) = (sheet.name.clone(), sheet.visible, sheet.merged.clone());
        out.append((name, visible, rows(py, sheet)?, merged))?;
    }
    Ok(out)
}

#[pymodule]
fn sheet_reader(m: &Bound<'_, PyModule>) -> PyResult<()> {
    m.add_function(wrap_pyfunction!(read, m)?)?;
    m.add("ReadError", m.py().get_type::<ReadError>())?;
    m.add("LimitExceeded", m.py().get_type::<LimitExceeded>())?;
    Ok(())
}
