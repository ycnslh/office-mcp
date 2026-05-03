"""MCP tools for Word document creation and editing.

Each tool returns a small dict keyed on ``file_uri`` so the calling LLM can
surface the result to the user and chain further edits without juggling
opaque handles.

Registration is done at the bottom of the module via ``mcp.tool(fn)`` so the
module-level names remain plain async functions — callable directly from
tests without going through the FastMCP transport.

The editing surface is **batched by intent**: a single call to
``docx_create`` (with a ``blocks`` list) can produce a complete document,
and editing tools (``docx_append``, ``docx_replace``, ``docx_remove``) each
accept a list of operations applied in one round-trip.
"""

from __future__ import annotations

import asyncio
from typing import Any

from office_mcp.context import require_current_user
from office_mcp.engines import docx_engine
from office_mcp.instrumentation import register_tool
from office_mcp.runtime import download_url_for, get_registry
from office_mcp.utils import libreoffice


async def docx_create(
    file_uri: str,
    title: str | None = None,
    blocks: list[dict[str, Any]] | None = None,
) -> dict[str, Any]:
    """Create a complete Word document in a single call.

    Use this directly. Do not call ``fs_list``, ``fs_stat`` or any other
    tool to "check OneDrive" or "see if the file exists" first — this
    tool creates or overwrites at ``file_uri`` in one step, surfaces
    ``auth_required`` on the first OneDrive touch the same way a list
    would, and normalizes bare-name URIs server-side.

    Pass the entire document body as ``blocks`` — an ordered list of typed
    content blocks — and this one call writes the finished file. Do not
    follow up with separate append/insert calls; everything fits here.

    Block schema:
        ``{"type": "heading", "level": 1, "text": "..."}`` (level: 1, 2 or 3)
        ``{"type": "paragraph", "text": "..."}``
        ``{"type": "table", "headers": ["..."], "rows": [["..."]]}``

    Args:
        file_uri: Destination URI, e.g. ``local://reports/q3.docx`` or
            ``od://me/Documents/report.docx``.
        title: Optional document title rendered as the Word ``Title`` style
            at the top of the document.
        blocks: Ordered list of headings, paragraphs and tables that make
            up the document body. Empty (or omitted) produces a doc with
            only the title.

    Returns:
        ``file_uri`` / ``size_bytes`` / ``etag`` / ``download_url`` and a
        one-line ``summary``.
    """
    user = require_current_user()
    backend = get_registry().get(file_uri)
    payload = docx_engine.compose(title=title, blocks=blocks or [])
    meta = await backend.write(file_uri, payload, user_ctx=user)
    block_count = len(blocks) if blocks else 0
    return {
        "file_uri": meta.uri,
        "size_bytes": meta.size_bytes,
        "etag": meta.etag,
        "download_url": download_url_for(meta, user.user_id),
        "summary": (
            f"Created document at {meta.uri}"
            + (f" with title {title!r}" if title else "")
            + (f" and {block_count} block(s)" if block_count else "")
        ),
    }


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


async def docx_append(
    file_uri: str,
    blocks: list[dict[str, Any]],
) -> dict[str, Any]:
    """Append multiple blocks (headings, paragraphs, tables) to a document in one call.

    Use this — not multiple sequential calls — to add several sections at
    once. ``blocks`` is an ordered list using the same schema as
    ``docx_create``:

        ``{"type": "heading", "level": 1, "text": "..."}``
        ``{"type": "paragraph", "text": "..."}``
        ``{"type": "table", "headers": ["..."], "rows": [["..."]]}``

    Args:
        file_uri: URI of the document to mutate.
        blocks: Ordered list of blocks to append at the end of the document.
    """
    user = require_current_user()
    backend = get_registry().get(file_uri)
    original = await backend.read(file_uri, user_ctx=user)
    updated = docx_engine.append_blocks(original, blocks=blocks)
    meta = await backend.write(file_uri, updated, user_ctx=user)
    return {
        "file_uri": meta.uri,
        "size_bytes": meta.size_bytes,
        "etag": meta.etag,
        "download_url": download_url_for(meta, user.user_id),
        "summary": f"Appended {len(blocks)} block(s) to {meta.uri}",
    }


async def docx_replace(
    file_uri: str,
    replacements: list[dict[str, Any]],
) -> dict[str, Any]:
    """Apply many find/replace substitutions to a document in one call.

    Each replacement is ``{"find": str, "replace": str, "regex": bool}``
    where ``regex`` defaults to false. All replacements are applied in a
    single pass over the document (paragraphs, headings and table cells).

    Args:
        file_uri: URI of the document to mutate.
        replacements: List of substitutions to apply.

    Returns the standard response plus ``replacements_made`` (total count).
    """
    user = require_current_user()
    backend = get_registry().get(file_uri)
    original = await backend.read(file_uri, user_ctx=user)
    updated, count = docx_engine.replace_text(original, replacements=replacements)
    meta = await backend.write(file_uri, updated, user_ctx=user)
    return {
        "file_uri": meta.uri,
        "size_bytes": meta.size_bytes,
        "etag": meta.etag,
        "download_url": download_url_for(meta, user.user_id),
        "replacements_made": count,
        "summary": (
            f"Applied {len(replacements)} replacement rule(s), "
            f"{count} substitution(s) made in {meta.uri}"
        ),
    }


async def docx_remove(
    file_uri: str,
    headings: list[str],
) -> dict[str, Any]:
    """Remove one or more sections from a document in one call.

    Each entry in ``headings`` is the exact text of a heading. The
    matching heading paragraph and every block that follows it — until
    the next heading of the same or higher level — is removed.

    Args:
        file_uri: URI of the document to mutate.
        headings: List of heading texts identifying sections to drop.

    Returns the standard response plus ``sections_removed`` (count of
    headings actually matched and removed).
    """
    user = require_current_user()
    backend = get_registry().get(file_uri)
    original = await backend.read(file_uri, user_ctx=user)
    updated, removed = docx_engine.remove_sections(original, headings=headings)
    meta = await backend.write(file_uri, updated, user_ctx=user)
    return {
        "file_uri": meta.uri,
        "size_bytes": meta.size_bytes,
        "etag": meta.etag,
        "download_url": download_url_for(meta, user.user_id),
        "sections_removed": removed,
        "summary": (
            f"Removed {removed}/{len(headings)} section(s) from {meta.uri}"
        ),
    }


async def docx_to_pdf(
    file_uri: str,
    output_uri: str | None = None,
) -> dict[str, Any]:
    """Render a Word document as PDF through headless LibreOffice.

    Args:
        file_uri: URI of the source ``.docx``.
        output_uri: Where to write the PDF. Defaults to the source URI
            with the extension swapped to ``.pdf``.

    Returns the destination ``file_uri``, ``size_bytes`` and ``etag``.
    """
    user = require_current_user()
    source_backend = get_registry().get(file_uri)
    data = await source_backend.read(file_uri, user_ctx=user)

    pdf_bytes = await asyncio.to_thread(
        libreoffice.convert, data, src_ext="docx", dst_ext="pdf"
    )

    target_uri = output_uri or _swap_extension(file_uri, ".pdf")
    destination_backend = get_registry().get(target_uri)
    meta = await destination_backend.write(
        target_uri, pdf_bytes, user_ctx=user
    )
    return {
        "source_uri": file_uri,
        "file_uri": meta.uri,
        "size_bytes": meta.size_bytes,
        "etag": meta.etag,
        "download_url": download_url_for(meta, user.user_id),
        "summary": (
            f"Rendered {file_uri} → {meta.uri} ({meta.size_bytes} bytes)"
        ),
    }


def _swap_extension(uri: str, new_ext: str) -> str:
    """Return ``uri`` with its file extension replaced by ``new_ext``.

    If ``uri`` has no extension, ``new_ext`` is appended.
    """
    if not new_ext.startswith("."):
        new_ext = "." + new_ext
    dot = uri.rfind(".")
    slash = uri.rfind("/")
    if dot > slash:  # there is an extension after the last path separator
        return uri[:dot] + new_ext
    return uri + new_ext


# Register with FastMCP through the instrumentation wrapper. Done imperatively
# so the module-level names above remain plain async functions (the @mcp.tool
# decorator wraps them in a FunctionTool that is not directly callable).
for _fn in (
    docx_create,
    docx_read,
    docx_append,
    docx_replace,
    docx_remove,
    docx_to_pdf,
):
    register_tool(_fn)
