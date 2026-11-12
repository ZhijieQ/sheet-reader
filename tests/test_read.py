"""sheet_reader.read against python-calamine (the reference for cell values), and the limits it checks."""

from __future__ import annotations

import datetime as dt
import io
import subprocess
import sys

import pytest
from openpyxl import Workbook
from python_calamine import CalamineWorkbook, SheetTypeEnum, SheetVisibleEnum

import sheet_reader

BIG = {"max_rows": 1_048_576, "max_cols": 16_384}


def xlsx(fill) -> bytes:
    wb = Workbook()
    fill(wb)
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def reference(data: bytes) -> list:
    """What python-calamine gives for the visible worksheets, with empty cells as None."""
    wb = CalamineWorkbook.from_filelike(io.BytesIO(data))
    out = []
    for m in wb.sheets_metadata:
        if m.typ == SheetTypeEnum.WorkSheet and m.visible == SheetVisibleEnum.Visible:
            ws = wb.get_sheet_by_name(m.name)
            rows = [[None if c == "" else c for c in r] for r in ws.to_python(skip_empty_area=False)]
            out.append((m.name, rows, sorted(tuple(map(tuple, x)) for x in (ws.merged_cell_ranges or []))))
    return out


def typed(sheets: list) -> list:
    return [(n, [[(type(c).__name__, c) for c in r] for r in rows], merged) for n, rows, merged in sheets]


def ours(data: bytes, **kw) -> list:
    return [(n, rows, sorted(tuple(map(tuple, m)) for m in merged)) for n, _, rows, merged in sheet_reader.read(data, **(BIG | kw))]


def every_kind(wb: Workbook) -> None:
    ws = wb.active
    ws.title = "数据"
    ws.append(["运单号", "件数", "单价", "带电", "发货日期", "下单时间", "时刻", "备注"])
    ws.append(["BG-1", 2, 12.5, True, dt.date(2026, 11, 4), dt.datetime(2026, 11, 3, 9, 15, 30), dt.time(8, 30), None])
    ws.append(["BG-2", 10**12, -0.001, False, dt.date(1999, 12, 31), dt.datetime(2026, 1, 1, 0, 0, 1), dt.time(0, 0, 1), "  有空格 "])
    ws["C5"] = "中间空了一行"
    ws.merge_cells("A6:B7")
    ws["A6"] = "合并"
    second = wb.create_sheet("第二页")
    second["D3"] = 3.25
    hidden = wb.create_sheet("隐藏")
    hidden["A1"] = "看不见"
    hidden.sheet_state = "hidden"
    wb.create_sheet("空表")


def test_cells_are_python_calaminess_values():
    data = xlsx(every_kind)
    got = ours(data)
    assert typed(got) == typed(reference(data))
    assert [n for n, _, _ in got] == ["数据", "第二页", "空表"]
    assert got[1][1][2] == [None, None, None, 3.25] and got[2][1] == []


def test_hidden_sheets_on_request():
    sheets = sheet_reader.read(xlsx(every_kind), hidden=True, **BIG)
    assert [(n, visible) for n, visible, _, _ in sheets] == [("数据", True), ("第二页", True), ("隐藏", False), ("空表", True)]


@pytest.mark.parametrize("ref", ["XFD1048576", "XFD10", "A1048576"])
def test_a_cell_past_the_limits_is_refused_without_building_the_grid(ref):
    """XFD1048576 made calamine ask for 512 GB and abort the process (tafia/calamine#693). Run in a child process:
    an abort would end the test run."""
    code = f"""
import io, sheet_reader
from openpyxl import Workbook
wb = Workbook(); wb.active["A1"] = "x"; wb.active["{ref}"] = 1
buf = io.BytesIO(); wb.save(buf)
try:
    sheet_reader.read(buf.getvalue(), max_rows=200_000, max_cols=500)
except sheet_reader.LimitExceeded as exc:
    print("refused", isinstance(exc, ValueError))
"""
    out = subprocess.run([sys.executable, "-c", code], capture_output=True, text=True)
    assert (out.returncode, out.stdout.strip()) == (0, "refused True"), out.stderr[-300:]


def test_the_limits_count_rows_and_columns_from_one():
    data = xlsx(lambda wb: wb.active.cell(row=5, column=3, value=1))
    assert len(sheet_reader.read(data, max_rows=5, max_cols=3)[0][2]) == 5
    with pytest.raises(sheet_reader.LimitExceeded, match="row 5, column 3"):
        sheet_reader.read(data, max_rows=4, max_cols=3)
    with pytest.raises(sheet_reader.LimitExceeded):
        sheet_reader.read(data, max_rows=5, max_cols=2)


@pytest.mark.parametrize("data", [b"", b"not a workbook", b"PK\x03\x04broken zip", bytes(range(256)) * 10])
def test_what_is_not_a_workbook_is_a_read_error(data):
    with pytest.raises(sheet_reader.ReadError):
        sheet_reader.read(data, **BIG)
