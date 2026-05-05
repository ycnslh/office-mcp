"""MCP tools for Excel spreadsheet creation and editing.

Engine work is offloaded to a worker thread via :func:`asyncio.to_thread`
so the LibreOffice recalc invoked when ``recalc=True`` (3-5 s) does not
block the FastMCP event loop.

The editing surface is **batched by intent**: a single call to
``xlsx_create`` (with a ``sheets`` list carrying writes + formulas)
produces a complete workbook, and ``xlsx_update`` lets you apply many
writes / formulas across many sheets in one round-trip.
"""

from __future__ import annotations

import asyncio
from typing import Any

from office_mcp.context import require_current_user
from office_mcp.engines import xlsx_engine
from office_mcp.instrumentation import register_tool
from office_mcp.runtime import download_url_for, get_registry


async def xlsx_create(
    file_uri: str,
    sheets: list[dict[str, Any]],
    recalc: bool = True,
) -> dict[str, Any]:
    """Create a complete Excel workbook in a single call.

    Use this directly. Do not call ``fs_list``, ``fs_stat`` or any other
    tool to "check OneDrive" or "see if the file exists" first — this
    tool creates or overwrites at ``file_uri`` in one step. Pass the
    full workbook body via ``sheets`` (REQUIRED) and let one LibreOffice
    recalc at the end bake every formula in one pass. Do NOT follow up
    with ``xlsx_update``; the entire workbook fits here.

    Sheet schema (each entry of ``sheets`` is a dict)::

        {
            "name": "Inputs",
            "writes":   [{"start_cell": "A1",
                          "values": [["Region", "Units", "Price"]],
                          "style": "header"},
                         {"start_cell": "A2",
                          "values": [["EU", 1200, 49.9],
                                     ["US",  900, 52.0]],
                          "style": "input"}],
            "formulas": [{"cell": "D2", "formula": "=B2*C2"},
                         {"cell": "D3", "formula": "=B3*C3"},
                         {"cell": "D4", "formula": "=SUM(D2:D3)"}],
        }

    ``style`` is one of ``"input"`` (blue), ``"formula"`` (black),
    ``"link"`` (green for cross-sheet refs) or ``"header"`` (bold).

    Args:
        file_uri: Destination URI, e.g. ``od://AI Generated/budget.xlsx``.
        sheets: REQUIRED. Ordered list of sheet specs (name + writes
            + formulas). Pass ``[]`` only when you genuinely want an
            empty workbook with the default ``Sheet``.
        recalc: When ``True`` (default), LibreOffice computes formula
            cells once at the end so subsequent reads return values,
            not formula strings. Set ``False`` only if you intend to
            recalc later via a separate ``xlsx_update``.
    """
    user = require_current_user()
    backend = get_registry().get(file_uri)
    payload = await asyncio.to_thread(
        xlsx_engine.compose, sheets=sheets, recalc=recalc
    )
    meta = await backend.write(file_uri, payload, user_ctx=user)
    sheet_names = [s.get("name") for s in sheets if s.get("name")]
    return {
        "file_uri": meta.uri,
        "size_bytes": meta.size_bytes,
        "etag": meta.etag,
        "download_url": download_url_for(meta, user.user_id),
        "sheets": sheet_names,
        "recalculated": recalc,
        "summary": (
            f"Created workbook at {meta.uri}"
            + (f" with sheets {sheet_names}" if sheet_names else "")
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


async def xlsx_update(
    file_uri: str,
    sheets: list[dict[str, Any]],
    recalc: bool = True,
) -> dict[str, Any]:
    """Apply many writes and formulas to an existing workbook in one call.

    Use this — not multiple sequential calls — to push several writes
    or formulas across one or more sheets. ``sheets`` uses the same
    schema as ``xlsx_create``; sheets that don't exist yet are created
    automatically.

    Args:
        file_uri: URI of the workbook to mutate.
        sheets: Ordered list of sheet update specs (name + writes +
            formulas). At least one sheet entry is required.
        recalc: Run a single LibreOffice recalc at the end when
            formulas were set. Default ``True``.
    """
    user = require_current_user()
    backend = get_registry().get(file_uri)
    original = await backend.read(file_uri, user_ctx=user)
    updated = await asyncio.to_thread(
        xlsx_engine.apply_updates, original, sheets=sheets, recalc=recalc
    )
    meta = await backend.write(file_uri, updated, user_ctx=user)
    write_count = sum(len(s.get("writes") or []) for s in sheets)
    formula_count = sum(len(s.get("formulas") or []) for s in sheets)
    return {
        "file_uri": meta.uri,
        "size_bytes": meta.size_bytes,
        "etag": meta.etag,
        "download_url": download_url_for(meta, user.user_id),
        "recalculated": recalc and formula_count > 0,
        "summary": (
            f"Applied {write_count} write(s) and {formula_count} formula(s) "
            f"across {len(sheets)} sheet(s) in {meta.uri}"
        ),
    }


# Register with FastMCP imperatively — see tools/docx_tools.py for the why.
for _fn in (xlsx_create, xlsx_read_range, xlsx_update):
    register_tool(_fn)
