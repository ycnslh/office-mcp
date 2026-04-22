"""MCP tools for Excel spreadsheet creation and editing.

Engine work is offloaded to a worker thread via :func:`asyncio.to_thread`
so the LibreOffice recalc invoked by ``xlsx_apply_formula`` (3-5 s) does
not block the FastMCP event loop.
"""

from __future__ import annotations

import asyncio
from typing import Any

from office_mcp.context import require_current_user
from office_mcp.engines import xlsx_engine
from office_mcp.runtime import get_registry, mcp


async def xlsx_create(
    file_uri: str,
    sheets: list[str] | None = None,
) -> dict[str, Any]:
    """Create a new Excel workbook at ``file_uri``.

    Args:
        file_uri: Destination URI, e.g. ``local://budget/2026.xlsx``.
        sheets: Optional list of sheet names. Defaults to a single
            ``Sheet`` if omitted.
    """
    user = require_current_user()
    backend = get_registry().get(file_uri)
    payload = await asyncio.to_thread(xlsx_engine.create, sheets=sheets)
    meta = await backend.write(file_uri, payload, user_ctx=user)
    return {
        "file_uri": meta.uri,
        "size_bytes": meta.size_bytes,
        "etag": meta.etag,
        "sheets": sheets or ["Sheet"],
        "summary": (
            f"Created workbook at {meta.uri} with sheets "
            f"{sheets or ['Sheet']}"
        ),
    }


async def xlsx_read_range(
    file_uri: str,
    sheet: str,
    range_str: str,
) -> dict[str, Any]:
    """Read a rectangular range from a sheet.

    Args:
        file_uri: URI of an existing workbook.
        sheet: Sheet name.
        range_str: A1-style range, e.g. ``A1:C10`` or a single cell ``B3``.

    Returns:
        ``values`` as a 2D list of cached values. Formula cells return
        ``None`` until they have been recalculated.
    """
    user = require_current_user()
    backend = get_registry().get(file_uri)
    data = await backend.read(file_uri, user_ctx=user)
    values = await asyncio.to_thread(
        xlsx_engine.read_range, data, sheet=sheet, range_str=range_str
    )
    return {
        "file_uri": file_uri,
        "sheet": sheet,
        "range": range_str,
        "values": values,
    }


async def xlsx_write_range(
    file_uri: str,
    sheet: str,
    start_cell: str,
    values: list[list[Any]],
    style: str = "input",
) -> dict[str, Any]:
    """Write a 2D ``values`` block starting at ``start_cell``.

    Args:
        file_uri: URI of the workbook to mutate.
        sheet: Sheet name (created if it does not exist).
        start_cell: Top-left cell, e.g. ``A1``.
        values: 2D list of values to write.
        style: ``"input"`` (default, blue), ``"formula"`` (black),
            ``"link"`` (green) or ``"header"`` (bold).
    """
    user = require_current_user()
    backend = get_registry().get(file_uri)
    original = await backend.read(file_uri, user_ctx=user)
    updated = await asyncio.to_thread(
        xlsx_engine.write_range,
        original,
        sheet=sheet,
        start_cell=start_cell,
        values=values,
        style=style,
    )
    meta = await backend.write(file_uri, updated, user_ctx=user)
    rows = len(values)
    cols = len(values[0]) if values else 0
    return {
        "file_uri": meta.uri,
        "size_bytes": meta.size_bytes,
        "etag": meta.etag,
        "summary": (
            f"Wrote {rows}×{cols} {style} block at {sheet}!{start_cell} in {meta.uri}"
        ),
    }


async def xlsx_apply_formula(
    file_uri: str,
    sheet: str,
    cell: str,
    formula: str,
    recalc: bool = True,
) -> dict[str, Any]:
    """Set a formula in a single cell, optionally recalculating the workbook.

    With ``recalc=True`` (the default) LibreOffice opens the workbook
    headlessly and bakes computed values into every formula cell. This
    adds 3-5 s; pass ``recalc=False`` when chaining many formula writes
    and recalc once at the end with a final ``recalc=True`` call.
    """
    user = require_current_user()
    backend = get_registry().get(file_uri)
    original = await backend.read(file_uri, user_ctx=user)
    updated = await asyncio.to_thread(
        xlsx_engine.apply_formula,
        original,
        sheet=sheet,
        cell=cell,
        formula=formula,
        recalc=recalc,
    )
    meta = await backend.write(file_uri, updated, user_ctx=user)
    return {
        "file_uri": meta.uri,
        "size_bytes": meta.size_bytes,
        "etag": meta.etag,
        "recalculated": recalc,
        "summary": (
            f"Set formula {formula!r} at {sheet}!{cell} in {meta.uri}"
            + (" (recalculated)" if recalc else " (recalc skipped)")
        ),
    }


# Register with FastMCP imperatively — see tools/docx_tools.py for the why.
for _fn in (xlsx_create, xlsx_read_range, xlsx_write_range, xlsx_apply_formula):
    mcp.tool(_fn)
