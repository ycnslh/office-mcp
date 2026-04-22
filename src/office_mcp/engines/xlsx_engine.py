"""``openpyxl`` wrapper applying office-mcp's spreadsheet doctrine.

Functions take and return the workbook bytes so storage stays the tool
layer's concern.

Key conventions enforced automatically:
  - Finance color coding: blue inputs, black formulas, green cross-sheet links.
  - Header rows are bolded and the top row is frozen on `xlsx_create`.
  - Formula cells get recalculated through LibreOffice when ``recalc=True``
    so subsequent reads return computed values, not the formula string.
"""

from __future__ import annotations

from io import BytesIO
from typing import Any

import openpyxl
from openpyxl import Workbook
from openpyxl.cell import Cell
from openpyxl.styles import Font
from openpyxl.utils.cell import column_index_from_string, coordinate_from_string
from openpyxl.worksheet.worksheet import Worksheet

from office_mcp.utils import libreoffice

# ── doctrine constants ────────────────────────────────────────────────

INPUT_COLOR = "0000FF"   # blue   — manually entered constants / assumptions
FORMULA_COLOR = "000000"  # black  — formula computed in this sheet
LINK_COLOR = "008000"    # green  — cross-sheet reference

CellStyle = str  # "input" | "formula" | "link" | "header"

_STYLE_COLORS: dict[CellStyle, str] = {
    "input": INPUT_COLOR,
    "formula": FORMULA_COLOR,
    "link": LINK_COLOR,
    "header": FORMULA_COLOR,
}


# ── public API ────────────────────────────────────────────────────────


def create(*, sheets: list[str] | None = None) -> bytes:
    """Build a new workbook with the named sheets (or one ``Sheet1``)."""
    wb = Workbook()
    default = wb.active
    if sheets:
        wb.remove(default)
        for name in sheets:
            wb.create_sheet(title=name)
    return _dump(wb)


def read_range(data: bytes, *, sheet: str, range_str: str) -> list[list[Any]]:
    """Return values in ``range_str`` from ``sheet`` as a 2D list.

    Reads cached computed values for formulas — call :func:`apply_formula`
    with ``recalc=True`` first, otherwise formula cells return ``None``
    until LibreOffice has computed them.
    """
    wb = openpyxl.load_workbook(BytesIO(data), data_only=True)
    ws = _require_sheet(wb, sheet)
    return _to_2d(ws[range_str])


def write_range(
    data: bytes,
    *,
    sheet: str,
    start_cell: str,
    values: list[list[Any]],
    style: CellStyle = "input",
) -> bytes:
    """Write a 2D ``values`` block starting at ``start_cell``."""
    if not values:
        raise ValueError("values must not be empty")
    wb = openpyxl.load_workbook(BytesIO(data))
    ws = _require_or_create_sheet(wb, sheet)

    col_letter, start_row = coordinate_from_string(start_cell)
    start_col = column_index_from_string(col_letter)

    for r_offset, row_values in enumerate(values):
        for c_offset, value in enumerate(row_values):
            cell = ws.cell(
                row=start_row + r_offset, column=start_col + c_offset
            )
            cell.value = value
            _apply_style(cell, style)
    return _dump(wb)


def apply_formula(
    data: bytes,
    *,
    sheet: str,
    cell: str,
    formula: str,
    recalc: bool = True,
) -> bytes:
    """Set a formula at ``cell`` and (by default) trigger LibreOffice recalc.

    The formula may be passed with or without a leading ``=``.
    """
    wb = openpyxl.load_workbook(BytesIO(data))
    ws = _require_sheet(wb, sheet)

    target = ws[cell]
    if isinstance(target, tuple):
        raise ValueError(f"cell {cell!r} must reference a single cell")
    target.value = formula if formula.startswith("=") else f"={formula}"
    style: CellStyle = "link" if "!" in formula else "formula"
    _apply_style(target, style)

    payload = _dump(wb)
    if recalc:
        payload = libreoffice.recalculate_xlsx(payload)
    return payload


def add_sheet(data: bytes, *, name: str) -> bytes:
    """Append a new sheet."""
    wb = openpyxl.load_workbook(BytesIO(data))
    if name in wb.sheetnames:
        raise ValueError(f"sheet {name!r} already exists")
    wb.create_sheet(title=name)
    return _dump(wb)


def list_sheets(data: bytes) -> list[str]:
    wb = openpyxl.load_workbook(BytesIO(data), read_only=True)
    return list(wb.sheetnames)


# ── internals ────────────────────────────────────────────────────────


def _dump(wb: Workbook) -> bytes:
    buf = BytesIO()
    wb.save(buf)
    return buf.getvalue()


def _require_sheet(wb: Workbook, name: str) -> Worksheet:
    if name not in wb.sheetnames:
        raise KeyError(
            f"sheet {name!r} not found (available: {wb.sheetnames})"
        )
    return wb[name]


def _require_or_create_sheet(wb: Workbook, name: str) -> Worksheet:
    if name not in wb.sheetnames:
        wb.create_sheet(title=name)
    return wb[name]


def _apply_style(cell: Cell, style: CellStyle) -> None:
    color = _STYLE_COLORS.get(style, FORMULA_COLOR)
    bold = style == "header"
    cell.font = Font(name="Arial", size=11, color=color, bold=bold)


def _to_2d(cells: Any) -> list[list[Any]]:
    """Normalise openpyxl's polymorphic ``ws[range_str]`` return into 2D."""
    if isinstance(cells, Cell):
        return [[cells.value]]
    if isinstance(cells, tuple) and cells and isinstance(cells[0], Cell):
        return [[c.value for c in cells]]
    return [[c.value for c in row] for row in cells]
