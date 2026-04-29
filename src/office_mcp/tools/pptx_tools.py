"""MCP tools for PowerPoint deck creation and editing."""

from __future__ import annotations

import asyncio
from typing import Any

from office_mcp.context import require_current_user
from office_mcp.engines import pptx_engine
from office_mcp.instrumentation import register_tool
from office_mcp.runtime import build_download_url, get_registry


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
        "download_url": build_download_url(meta.uri, user.user_id),
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
    """Append a NEW slide to the end of an existing deck.

    To change a slide that already exists, use ``pptx_update_slide`` —
    this tool only appends.

    Layouts (default is ``content``, which is what you want 90% of the time):

    * ``content`` — title + bulleted body. Pass ``title="My headline"``
      and ``content=["First point", "Second point", ...]``. **This is the
      default body-slide layout; pick this whenever the user asks for a
      regular slide with text.**
    * ``two-content`` — title + two bullet columns. Pass ``content``
      (left column) and ``content_right`` (right column).
    * ``section`` — section divider. Title only, no bullets. Use to
      break the deck into parts.
    * ``title`` — title-slide layout (deck cover). Title + subtitle, no
      bullets. **Do not use this for body slides** — bullets passed via
      ``content`` are silently dropped on this layout.
    * ``blank`` — empty slide; title only, no body.

    Args:
        file_uri: URI of the deck to mutate (e.g. ``od://decks/q3.pptx``).
        title: Slide headline.
        layout: One of ``content`` (default), ``two-content``, ``section``,
            ``title``, ``blank``. See list above for when to pick which.
        content: Bullet lines for the main placeholder. Honoured for
            ``content`` and ``two-content`` layouts; ignored for the rest.
        content_right: Bullet lines for the second column (``two-content``).
        theme: ``corporate`` / ``minimal`` / ``cyberpunk``. Keep
            consistent with the rest of the deck.
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
        "download_url": build_download_url(meta.uri, user.user_id),
        "summary": (
            f"Added {layout} slide {title!r} "
            f"({bullet_count} bullets) to {meta.uri}"
        ),
    }


async def pptx_update_slide(
    file_uri: str,
    slide_index: int,
    title: str | None = None,
    content: list[str] | None = None,
    content_right: list[str] | None = None,
    theme: str = "corporate",
) -> dict[str, Any]:
    """Mutate an existing slide in place — use this to fill in body
    text on a slide that was created with just a title.

    All fields are optional and only acted on when supplied. Pass
    ``content`` to (re)populate the main bullet column; pass
    ``content_right`` for the second column on a ``two-content`` slide;
    pass ``title`` to rewrite the headline. Bullet replacements clear
    the existing paragraphs first.

    Args:
        file_uri: URI of the deck to mutate.
        slide_index: Zero-based index. Slide 0 is the first slide
            (typically the cover added by ``pptx_create``); call
            ``pptx_read`` first if you need the count.
        title: New slide title (leave ``None`` to keep the current one).
        content: Bullet lines for the main placeholder. Replaces — does
            not append to — the existing bullets.
        content_right: Bullet lines for the second column.
        theme: Colors / fonts to apply to the rewritten text. Keep
            consistent with the rest of the deck.
    """
    user = require_current_user()
    backend = get_registry().get(file_uri)
    original = await backend.read(file_uri, user_ctx=user)
    updated = await asyncio.to_thread(
        pptx_engine.update_slide,
        original,
        index=slide_index,
        title=title,
        content=content,
        content_right=content_right,
        theme=theme,
    )
    meta = await backend.write(file_uri, updated, user_ctx=user)
    fields_changed = [
        name
        for name, value in (
            ("title", title),
            ("content", content),
            ("content_right", content_right),
        )
        if value is not None
    ]
    return {
        "file_uri": meta.uri,
        "size_bytes": meta.size_bytes,
        "etag": meta.etag,
        "download_url": build_download_url(meta.uri, user.user_id),
        "summary": (
            f"Updated slide {slide_index} of {meta.uri} "
            f"(fields: {', '.join(fields_changed) or 'none'})"
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
for _fn in (pptx_create, pptx_add_slide, pptx_update_slide, pptx_read):
    register_tool(_fn)
