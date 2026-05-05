"""MCP tools for PowerPoint deck creation and editing.

Batched-by-intent surface
-------------------------
A single ``pptx_create`` call (with a ``slides`` list) produces a
complete deck. ``pptx_append`` pushes N slides in one round-trip and
``pptx_replace`` patches N slides by index in one call. The granular
``pptx_add_*`` / ``pptx_set_*`` family (4 + 2 tools) is gone — the
batched arguments cover every flavour through a discriminated ``type``
field, so the LLM always provides exactly the fields the slide needs.

Per-file write serialisation
----------------------------
Every mutating tool wraps its read → modify → write block in
``async with _locks.for_uri(file_uri):``. Two coroutines hitting the
same file no longer race — the second waits for the first to finish.
Without it, concurrent writes from a single chat turn trip Microsoft
Graph's eTag concurrency check (HTTP 409 ``resourceModified``).
"""

from __future__ import annotations

import asyncio
from typing import Any

from office_mcp.context import require_current_user
from office_mcp.engines import pptx_engine
from office_mcp.instrumentation import register_tool
from office_mcp.runtime import download_url_for, get_registry
from office_mcp.tools import _locks


async def pptx_create(
    file_uri: str,
    slides: list[dict[str, Any]],
    theme: str = "corporate",
) -> dict[str, Any]:
    """Create a complete PowerPoint deck in a single call.

    Use this directly. Do not call ``fs_list``, ``fs_stat`` or any other
    tool to "check OneDrive" or "see if the file exists" first — this
    tool creates or overwrites at ``file_uri`` in one step. Pass the
    full deck via ``slides`` (REQUIRED) and do NOT chain a follow-up
    ``pptx_append``; the entire deck fits here.

    Slide schema — every slide carries a ``type``::

        {"type": "title",       "title": "...", "subtitle": "..."}
        {"type": "content",     "title": "...", "bullets": [...]}
        {"type": "two-content", "title": "...", "left": [...], "right": [...]}
        {"type": "section",     "title": "..."}

    Args:
        file_uri: Destination URI, e.g. ``od://AI Generated/pitch.pptx``.
        slides: REQUIRED. Ordered list of slide blocks. Pass ``[]`` only
            when you genuinely want an empty deck (rare).
        theme: ``corporate`` (default), ``minimal``, or ``cyberpunk``.
            See ``office://rules/pptx``.
    """
    user = require_current_user()
    backend = get_registry().get(file_uri)
    async with _locks.for_uri(file_uri):
        payload = await asyncio.to_thread(
            pptx_engine.compose, theme=theme, slides=slides
        )
        meta = await backend.write(file_uri, payload, user_ctx=user)
    slide_count = len(slides)
    return {
        "file_uri": meta.uri,
        "size_bytes": meta.size_bytes,
        "etag": meta.etag,
        "download_url": download_url_for(meta, user.user_id),
        "theme": theme,
        "slide_count": slide_count,
        "summary": (
            f"Created {theme} deck at {meta.uri} with {slide_count} slide(s)"
        ),
    }


async def pptx_read(file_uri: str) -> dict[str, Any]:
    """Return slide titles and body text.

    Useful for the LLM to plan the next slide without having to remember
    everything it has already added.
    """
    user = require_current_user()
    backend = get_registry().get(file_uri)
    data = await backend.read(file_uri, user_ctx=user)
    structure = await asyncio.to_thread(pptx_engine.read, data)
    return {
        "file_uri": file_uri,
        "size_bytes": len(data),
        **structure,
    }


async def pptx_append(
    file_uri: str,
    slides: list[dict[str, Any]],
    theme: str = "corporate",
) -> dict[str, Any]:
    """Append multiple slides to a deck in one call.

    Use this — not multiple sequential calls — to add several slides at
    once. ``slides`` uses the same schema as ``pptx_create``.

    Args:
        file_uri: Deck URI.
        slides: Ordered list of slide blocks to append at the end.
        theme: Keep consistent with the rest of the deck.
    """
    user = require_current_user()
    backend = get_registry().get(file_uri)
    async with _locks.for_uri(file_uri):
        original = await backend.read(file_uri, user_ctx=user)
        updated = await asyncio.to_thread(
            pptx_engine.append_slides, original, slides=slides, theme=theme
        )
        meta = await backend.write(file_uri, updated, user_ctx=user)
    return {
        "file_uri": meta.uri,
        "size_bytes": meta.size_bytes,
        "etag": meta.etag,
        "download_url": download_url_for(meta, user.user_id),
        "summary": f"Appended {len(slides)} slide(s) to {meta.uri}",
    }


async def pptx_replace(
    file_uri: str,
    updates: list[dict[str, Any]],
    theme: str = "corporate",
) -> dict[str, Any]:
    """Patch one or more existing slides in a single call.

    Each entry of ``updates`` is ``{"slide_index": int, "title"?,
    "bullets"?, "left"?, "right"?}``. Fields that are present overwrite
    that part of the slide; absent fields are left alone. Bullet
    replacements clear the existing list before writing the new one
    (no append).

    Args:
        file_uri: Deck URI.
        updates: Ordered list of slide patches.
        theme: Keep consistent with the rest of the deck.
    """
    user = require_current_user()
    backend = get_registry().get(file_uri)
    async with _locks.for_uri(file_uri):
        original = await backend.read(file_uri, user_ctx=user)
        updated = await asyncio.to_thread(
            pptx_engine.replace_slides, original, updates=updates, theme=theme
        )
        meta = await backend.write(file_uri, updated, user_ctx=user)
    return {
        "file_uri": meta.uri,
        "size_bytes": meta.size_bytes,
        "etag": meta.etag,
        "download_url": download_url_for(meta, user.user_id),
        "summary": f"Patched {len(updates)} slide(s) in {meta.uri}",
    }


# Register with FastMCP imperatively — see tools/docx_tools.py for the why.
for _fn in (pptx_create, pptx_read, pptx_append, pptx_replace):
    register_tool(_fn)
