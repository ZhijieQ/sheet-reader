# sheet-reader

Read spreadsheets (xlsx, xlsm, xlsb, xls, ods) into Python lists, with row and column limits checked **while
reading**. Built on [calamine](https://github.com/tafia/calamine) (Rust); installs as a prebuilt wheel, so users
need Python only.

```python
import sheet_reader

for name, visible, rows, merged in sheet_reader.read(data, max_rows=200_000, max_cols=500):
    ...  # rows: list of lists of cell values (None where empty); merged: [((r1, c1), (r2, c2)), ...], 0-based
```

- `read(data, *, max_rows, max_cols, hidden=False)` reads every worksheet (hidden ones only with `hidden=True`).
- Cell values are what [python-calamine](https://github.com/dimastbk/python-calamine)'s `to_python()` gives:
  `int`, `float`, `str`, `bool`, `datetime.date` / `datetime` / `time` / `timedelta`; empty cells and error values
  are `None`. Rows start at the first row and column (A1) and reach the farthest used cell.
- `sheet_reader.LimitExceeded` when a sheet has a cell past `max_rows` rows or `max_cols` columns;
  `sheet_reader.ReadError` (a `ValueError`) when the file cannot be read.

## Why

calamine builds a sheet's whole grid in one allocation sized by its farthest cell, and a failed allocation aborts
the process — a Python program cannot catch it ([tafia/calamine#693](https://github.com/tafia/calamine/issues/693)).
One value in the last cell Excel allows (XFD1048576) of an otherwise tiny, valid xlsx asks for 512 GB.

For **xlsx, xlsm and xlsb** this package reads the cells one by one (calamine's cells reader), refuses with
`LimitExceeded` as soon as a cell lies past the limits, and builds the rows only then. **xls and ods** are
densified by calamine while the workbook is opened, before any limit can be checked: if you read untrusted .xls or
.ods files, read them in a separate process with a memory limit.

## Build

```bash
maturin build --release        # wheel in target/wheels/ (abi3: one wheel per platform for Python 3.10+)
cargo clippy --all-targets && cargo fmt --check
```

## License

MIT. The cell value conversion follows python-calamine (MIT); see LICENSE.
