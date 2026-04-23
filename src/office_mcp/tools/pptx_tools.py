"""MCP tools for PowerPoint deck creation and editing."""

from __future__ import annotations

import asyncio
from typing import Any

from office_mcp.context import require_current_user
from office_mcp.engines import pptx_engine
from office_mcp.runtime import get_registry, mcp


async def pptx_create(
    file_uri: str,
    theme: str = "corporate",
    title: str | None = None,
    subtitle: str | None = None,
) -> dict[str, Any]:
    """Create a new PowerPoint deck at ``file_uri``.

    Args:
        file_uri: Destination URI, e.g. ``local://decks/pitch.pptx``.
        theme: Visual theme — ``corporate`` (default), ``minimal``, or
            ``cyberpunk``. See ``office://rules/pptx``.
        title: Optional deck title; when set, a title slide is added.
        subtitle: Optional subtitle for the title slide.
    """
    user = require_current_user()
    backend = get_registry().get(file_uri)
    payload = await asyncio.to_thread(
        pptx_engine.create, theme=theme, title=title, subtitle=subtitle
    )
    meta = await backend.write(file_uri, payload, user_ctx=user)
    return {
        "file_uri": meta.uri,
        "size_bytes": meta.size_bytes,
        "etag": meta.etag,
        "theme": theme,
        "summary": (
            f"Created {theme} deck at {meta.uri}"
            + (f" with title {title!r}" if title else "")
        ),
    }


async def pptx_add_slide(
    file_uri: str,
    title: str,
    layout: str = "content",
    content: list[str] | None = None,
    content_right: list[str] | None = None,
    theme: str = "corporate",
) -> dict[str, Any]:
    """Append a slide to an existing deck.

    Args:
        file_uri: URI of the deck to mutate.
        title: Slide headline.
        layout: ``title`` / ``content`` / ``section`` / ``two-content`` / ``blank``.
        content: Bullet lines for the main placeholder (``content`` and
            ``two-content`` layouts).
        content_right: Bullet lines for the second column (``two-content``).
        theme: Colors / fonts; keep consistent with the rest of the deck.
    """
    user = require_current_user()
    backend = get_registry().get(file_uri)
    original = await backend.read(file_uri, user_ctx=user)
    updated = await asyncio.to_thread(
        pptx_engine.add_slide,
        original,
        layout=layout,
        title=title,
        content=content,
        content_right=content_right,
        theme=theme,
    )
    meta = await backend.write(file_uri, updated, user_ctx=user)
    bullet_count = len(content or []) + len(content_right or [])
    return {
        "file_uri": meta.uri,
        "size_bytes": meta.size_bytes,
        "etag": meta.etag,
        "summary": (
            f"Added {layout} slide {title!r} "
            f"({bullet_count} bullets) to {meta.uri}"
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


# Register with FastMCP imperatively — see tools/docx_tools.py for the why.
for _fn in (pptx_create, pptx_add_slide, pptx_read):
    mcp.tool(_fn)
