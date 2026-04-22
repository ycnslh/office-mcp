"""MCP tools for Word document creation and editing.

Each tool returns a small dict keyed on ``file_uri`` so the calling LLM can
surface the result to the user and chain further edits without juggling
opaque handles.
"""

from __future__ import annotations

from typing import Any

from office_mcp.context import require_current_user
from office_mcp.engines import docx_engine
from office_mcp.runtime import get_registry, mcp


@mcp.tool
async def docx_create(file_uri: str, title: str | None = None) -> dict[str, Any]:
    """Create a new, empty Word document at ``file_uri``.

    Args:
        file_uri: Destination URI, e.g. ``local://reports/q3.docx``.
        title: Optional document title rendered as the Word ``Title`` style.

    Returns:
        ``file_uri`` / ``size_bytes`` / ``etag`` and a one-line ``summary``.
    """
    user = require_current_user()
    backend = get_registry().get(file_uri)
    payload = docx_engine.create(title=title)
    meta = await backend.write(file_uri, payload, user_ctx=user)
    return {
        "file_uri": meta.uri,
        "size_bytes": meta.size_bytes,
        "etag": meta.etag,
        "summary": (
            f"Created document at {meta.uri}"
            + (f" with title {title!r}" if title else "")
        ),
    }


@mcp.tool
async def docx_read(file_uri: str) -> dict[str, Any]:
    """Read a Word document and return its structured content.

    Args:
        file_uri: URI of an existing document.

    Returns:
        Paragraphs (with ``heading``/``paragraph`` types), tables (as row
        arrays), and counters. Useful for the LLM to plan an edit.
    """
    user = require_current_user()
    backend = get_registry().get(file_uri)
    data = await backend.read(file_uri, user_ctx=user)
    structure = docx_engine.read(data)
    return {
        "file_uri": file_uri,
        "size_bytes": len(data),
        **structure,
    }


@mcp.tool
async def docx_append_section(
    file_uri: str,
    heading: str,
    paragraphs: list[str],
    level: int = 1,
) -> dict[str, Any]:
    """Append a heading + body paragraphs to the end of a document.

    Args:
        file_uri: URI of the document to mutate.
        heading: Section heading text.
        paragraphs: Body paragraphs (each rendered as a Word paragraph).
        level: Heading level 1-3 (defaults to 1).
    """
    user = require_current_user()
    backend = get_registry().get(file_uri)
    original = await backend.read(file_uri, user_ctx=user)
    updated = docx_engine.append_section(
        original, heading=heading, paragraphs=paragraphs, level=level
    )
    meta = await backend.write(file_uri, updated, user_ctx=user)
    return {
        "file_uri": meta.uri,
        "size_bytes": meta.size_bytes,
        "etag": meta.etag,
        "summary": (
            f"Appended H{level} section {heading!r} "
            f"({len(paragraphs)} paragraphs) to {meta.uri}"
        ),
    }


@mcp.tool
async def docx_insert_table(
    file_uri: str,
    headers: list[str],
    rows: list[list[str]],
) -> dict[str, Any]:
    """Append a table to the document with explicit DXA column widths.

    Args:
        file_uri: URI of the document to mutate.
        headers: Column headers (rendered bold).
        rows: Body rows; each must have exactly ``len(headers)`` cells.
    """
    user = require_current_user()
    backend = get_registry().get(file_uri)
    original = await backend.read(file_uri, user_ctx=user)
    updated = docx_engine.insert_table(original, headers=headers, rows=rows)
    meta = await backend.write(file_uri, updated, user_ctx=user)
    return {
        "file_uri": meta.uri,
        "size_bytes": meta.size_bytes,
        "etag": meta.etag,
        "summary": (
            f"Inserted {len(rows)}×{len(headers)} table into {meta.uri}"
        ),
    }
