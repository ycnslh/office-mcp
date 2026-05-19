"""``openpyxl`` wrapper applying office-mcp's spreadsheet doctrine.

Functions take and return the workbook bytes so storage stays the tool
layer's concern.

The public API is **batched-by-intent**: a single ``compose`` call
produces a complete workbook (sheets + value writes + formulas), and
``apply_updates`` lets the tool layer push N writes / N formulas onto
an existing workbook in a single read/modify/write round-trip.

Sheet schema (used by ``compose`` and ``apply_updates``)::

    {
        "name": "Inputs",
        "writes":   [{"start_cell": "A1",
                      "values": [["Region", "Units", "Price"]],
                      "style": "header"}],
        "formulas": [{"cell": "D2", "formula": "=B2*C2"}],
    }

``writes`` and ``formulas`` are both optional per-sheet — at least one
should be non-empty for the sheet to carry any data.

Key conventions enforced automatically:
  - Finance color coding: blue inputs, black formulas, green cross-sheet links.
  - Header rows are bolded.
  - Formula cells get recalculated through LibreOffice when ``recalc=True``
    (a single recalc at the end batches the whole workbook in one pass).
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

from . import _libreoffice as libreoffice

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
_VALID_STYLES = frozenset(_STYLE_COLORS)


# ── public API ────────────────────────────────────────────────────────


def compose(
    *,
    sheets: list[dict[str, Any]] | None = None,
    recalc: bool = True,
) -> bytes:
    """Build a complete workbook in one call.

    Args:
        sheets: Ordered list of sheet specifications. See module docstring
            for the schema. When omitted, produces an empty single-sheet
            workbook.
        recalc: When ``True`` (default) and any formulas are present,
            invoke LibreOffice once at the end to bake computed values.
    """
    wb = Workbook()
    default = wb.active
    sheets_spec = sheets or []
    if sheets_spec:
        wb.remove(default)
        for i, spec in enumerate(sheets_spec):
            name = spec.get("name")
            if not isinstance(name, str) or not name:
                raise ValueError(f"sheet {i}: 'name' must be a non-empty string")
            if name in wb.sheetnames:
                raise ValueError(f"sheet {i}: duplicate name {name!r}")
            ws = wb.create_sheet(title=name)
            _apply_writes_to_worksheet(
                ws, spec.get("writes") or [], where=f"sheet {i}"
            )
            _apply_formulas_to_worksheet(
                ws, spec.get("formulas") or [], where=f"sheet {i}"
            )
    payload = _dump(wb)
    if recalc and _has_formulas(sheets_spec):
        payload = libreoffice.recalculate_xlsx(payload)
    return payload


def apply_updates(
    data: bytes,
    *,
    sheets: list[dict[str, Any]],
    recalc: bool = True,
) -> bytes:
    """Apply batched writes and formulas across one or more sheets.

    Each sheet entry uses the same schema as :func:`compose`. Sheets
    that don't yet exist are created. Returns the updated bytes; runs
    a single LibreOffice recalc at the end when ``recalc=True`` and
    at least one formula was set.
    """
    if not sheets:
        return data
    wb = openpyxl.load_workbook(BytesIO(data))
    any_formula = False
    for i, spec in enumerate(sheets):
        name = spec.get("name")
        if not isinstance(name, str) or not name:
            raise ValueError(f"sheet {i}: 'name' must be a non-empty string")
        ws = _require_or_create_sheet(wb, name)
        _apply_writes_to_worksheet(
            ws, spec.get("writes") or [], where=f"sheet {i}"
        )
        if _apply_formulas_to_worksheet(
            ws, spec.get("formulas") or [], where=f"sheet {i}"
        ):
            any_formula = True
    payload = _dump(wb)
    if recalc and any_formula:
        payload = libreoffice.recalculate_xlsx(payload)
    return payload


def read_range(data: bytes, *, sheet: str, range_str: str) -> list[list[Any]]:
    """Return values in ``range_str`` from ``sheet`` as a 2D list.

    Reads cached computed values for formulas. If formulas have not been
    recalculated they will return ``None`` until LibreOffice has computed
    them — call ``compose`` / ``apply_updates`` with ``recalc=True``.
    """
    wb = openpyxl.load_workbook(BytesIO(data), data_only=True)
    ws = _require_sheet(wb, sheet)
    return _to_2d(ws[range_str])


def list_sheets(data: bytes) -> list[str]:
    wb = openpyxl.load_workbook(BytesIO(data), read_only=True)
    return list(wb.sheetnames)


# ── internals ────────────────────────────────────────────────────────


def _has_formulas(sheets: list[dict[str, Any]]) -> bool:
    return any(spec.get("formulas") for spec in sheets)


def _apply_writes_to_worksheet(
    ws: Worksheet, writes: list[dict[str, Any]], *, where: str
) -> None:
    for w_idx, write in enumerate(writes):
        start_cell = write.get("start_cell")
        if not isinstance(start_cell, str) or not start_cell:
            raise ValueError(
                f"{where} write {w_idx}: 'start_cell' must be a non-empty string"
            )
        values = write.get("values")
        if not isinstance(values, list) or not values:
            raise ValueError(
                f"{where} write {w_idx}: 'values' must be a non-empty 2D list"
            )
        style = write.get("style", "input")
        if style not in _VALID_STYLES:
            raise ValueError(
                f"{where} write {w_idx}: unknown style {style!r}; "
                f"expected one of {sorted(_VALID_STYLES)}"
            )
        col_letter, start_row = coordinate_from_string(start_cell)
        start_col = column_index_from_string(col_letter)
        for r_offset, row_values in enumerate(values):
            if not isinstance(row_values, list):
                raise ValueError(
                    f"{where} write {w_idx}: each row of 'values' must be a list"
                )
            for c_offset, value in enumerate(row_values):
                cell = ws.cell(
                    row=start_row + r_offset, column=start_col + c_offset
                )
                cell.value = value
                _apply_style(cell, style)


def _apply_formulas_to_worksheet(
    ws: Worksheet, formulas: list[dict[str, Any]], *, where: str
) -> bool:
    """Set formulas; return ``True`` if any formula was written."""
    wrote = False
    for f_idx, spec in enumerate(formulas):
        cell_ref = spec.get("cell")
        if not isinstance(cell_ref, str) or not cell_ref:
            raise ValueError(
                f"{where} formula {f_idx}: 'cell' must be a non-empty string"
            )
        formula = spec.get("formula")
        if not isinstance(formula, str) or not formula:
            raise ValueError(
                f"{where} formula {f_idx}: 'formula' must be a non-empty string"
            )
        target = ws[cell_ref]
        if isinstance(target, tuple):
            raise ValueError(
                f"{where} formula {f_idx}: cell {cell_ref!r} must be a single cell"
            )
        target.value = formula if formula.startswith("=") else f"={formula}"
        style: CellStyle = "link" if "!" in formula else "formula"
        _apply_style(target, style)
        wrote = True
    return wrote


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
