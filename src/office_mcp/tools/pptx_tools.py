"""MCP tools for PowerPoint deck creation and editing.

Layout-specific tools
---------------------
Slides come in distinct flavours (cover / bullet body / two-column /
section divider). Earlier we exposed a single ``pptx_add_slide`` that
took a ``layout`` enum plus an optional ``content: list[str] | None``;
weaker LLMs (e.g. MiniMax) read the optional ``None`` default as
"skip" and produced decks with titles but no body. Splitting the API
so each tool's name maps to one slide flavour, with **required**
``bullets`` arguments, removes that ambiguity. The surface is wider
(4 add tools instead of 1) but each call is shorter and harder to
mis-fill.

Per-file write serialisation
----------------------------
Every mutating tool wraps its read → modify → write block in
``async with _locks.for_uri(file_uri):``. Two coroutines hitting the
same file no longer race — the second waits for the first to finish.
Without it, concurrent ``pptx_set_*`` calls from a single chat turn
trip Microsoft Graph's eTag concurrency check (HTTP 409
``resourceModified``).
"""

from __future__ import annotations

import asyncio
from typing import Any

from office_mcp.context import require_current_user
from office_mcp.engines import pptx_engine
from office_mcp.instrumentation import register_tool
from office_mcp.runtime import build_download_url, get_registry
from office_mcp.tools import _locks


async def _read(file_uri: str) -> tuple[bytes, Any]:
    user = require_current_user()
    backend = get_registry().get(file_uri)
    data = await backend.read(file_uri, user_ctx=user)
    return data, backend


async def _write(
    file_uri: str, payload: bytes, *, summary: str
) -> dict[str, Any]:
    user = require_current_user()
    backend = get_registry().get(file_uri)
    meta = await backend.write(file_uri, payload, user_ctx=user)
    return {
        "file_uri": meta.uri,
        "size_bytes": meta.size_bytes,
        "etag": meta.etag,
        "download_url": build_download_url(meta.uri, user.user_id),
        "summary": summary,
    }


# ── deck creation ─────────────────────────────────────────────────────


async def pptx_create(
    file_uri: str,
    theme: str = "corporate",
    title: str | None = None,
    subtitle: str | None = None,
) -> dict[str, Any]:
    """Create a new PowerPoint deck at ``file_uri``.

    Pass ``title`` to add a cover slide as the first slide; otherwise
    the deck is empty and you populate it via the ``pptx_add_*_slide``
    family.

    Args:
        file_uri: Destination URI, e.g. ``od://decks/pitch.pptx``.
        theme: ``corporate`` (default), ``minimal``, or ``cyberpunk``.
            See ``office://rules/pptx``.
        title: Optional cover-slide title.
        subtitle: Optional cover-slide subtitle (only meaningful with
            ``title``).
    """
    user = require_current_user()
    backend = get_registry().get(file_uri)
    async with _locks.for_uri(file_uri):
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


# ── slide-flavour-specific add tools ──────────────────────────────────


async def pptx_add_title_slide(
    file_uri: str,
    title: str,
    subtitle: str = "",
    theme: str = "corporate",
) -> dict[str, Any]:
    """Append a TITLE slide (deck cover style) to the deck.

    The title slide is the cover/section-cover layout: a big title and
    an optional subtitle, no bullets. Use this for the first slide of
    a fresh deck or to break it into named parts.

    Args:
        file_uri: Deck URI.
        title: Big headline shown centred.
        subtitle: Optional smaller text under the title.
        theme: Keep consistent with the rest of the deck.
    """
    async with _locks.for_uri(file_uri):
        original, _ = await _read(file_uri)
        updated = await asyncio.to_thread(
            pptx_engine.add_slide,
            original,
            layout="title",
            title=title,
            subtitle=subtitle or None,
            theme=theme,
        )
        return await _write(
            file_uri, updated, summary=f"Added title slide {title!r} to {file_uri}"
        )


async def pptx_add_content_slide(
    file_uri: str,
    title: str,
    bullets: list[str],
    theme: str = "corporate",
) -> dict[str, Any]:
    """Append a CONTENT slide — title + bulleted body.

    This is the standard body-slide layout. Use it for every regular
    slide that has a headline and a list of points underneath. The
    ``bullets`` argument is required; pass at least one string. Each
    list element becomes one bullet line.

    Args:
        file_uri: Deck URI.
        title: Slide headline.
        bullets: One string per bullet line, in display order. Required.
        theme: Keep consistent with the rest of the deck.
    """
    async with _locks.for_uri(file_uri):
        original, _ = await _read(file_uri)
        updated = await asyncio.to_thread(
            pptx_engine.add_slide,
            original,
            layout="content",
            title=title,
            content=bullets,
            theme=theme,
        )
        return await _write(
            file_uri,
            updated,
            summary=(
                f"Added content slide {title!r} ({len(bullets)} bullets) "
                f"to {file_uri}"
            ),
        )


async def pptx_add_two_content_slide(
    file_uri: str,
    title: str,
    left_bullets: list[str],
    right_bullets: list[str],
    theme: str = "corporate",
) -> dict[str, Any]:
    """Append a TWO-CONTENT slide — title + two parallel bullet columns.

    Use it for compare/contrast slides (pros/cons, before/after, etc).
    Both ``left_bullets`` and ``right_bullets`` are required and should
    each contain at least one entry; pass a single-item list if you
    want a one-line column.

    Args:
        file_uri: Deck URI.
        title: Slide headline.
        left_bullets: Left-column bullets, one string per line. Required.
        right_bullets: Right-column bullets, one string per line. Required.
        theme: Keep consistent with the rest of the deck.
    """
    async with _locks.for_uri(file_uri):
        original, _ = await _read(file_uri)
        updated = await asyncio.to_thread(
            pptx_engine.add_slide,
            original,
            layout="two-content",
            title=title,
            content=left_bullets,
            content_right=right_bullets,
            theme=theme,
        )
        return await _write(
            file_uri,
            updated,
            summary=(
                f"Added two-content slide {title!r} "
                f"({len(left_bullets)}+{len(right_bullets)} bullets) to {file_uri}"
            ),
        )


async def pptx_add_section_slide(
    file_uri: str,
    title: str,
    theme: str = "corporate",
) -> dict[str, Any]:
    """Append a SECTION-HEADER slide — large title only, no bullets.

    Use it as a divider between sections of a longer deck (e.g.
    "Part II — Risks"). For a content-bearing slide use
    ``pptx_add_content_slide``.

    Args:
        file_uri: Deck URI.
        title: Section headline.
        theme: Keep consistent with the rest of the deck.
    """
    async with _locks.for_uri(file_uri):
        original, _ = await _read(file_uri)
        updated = await asyncio.to_thread(
            pptx_engine.add_slide,
            original,
            layout="section",
            title=title,
            theme=theme,
        )
        return await _write(
            file_uri, updated, summary=f"Added section slide {title!r} to {file_uri}"
        )


# ── slide-mutation tools ──────────────────────────────────────────────


async def pptx_set_slide_title(
    file_uri: str,
    slide_index: int,
    title: str,
    theme: str = "corporate",
) -> dict[str, Any]:
    """Replace the title of slide ``slide_index`` (0-based).

    Args:
        file_uri: Deck URI.
        slide_index: Zero-based index of the slide to mutate. Slide 0
            is typically the cover. Call ``pptx_read`` first to confirm
            the count if unsure.
        title: New title text.
        theme: Apply theme styling to the rewritten title.
    """
    async with _locks.for_uri(file_uri):
        original, _ = await _read(file_uri)
        updated = await asyncio.to_thread(
            pptx_engine.update_slide,
            original,
            index=slide_index,
            title=title,
            theme=theme,
        )
        return await _write(
            file_uri,
            updated,
            summary=f"Set title of slide {slide_index} to {title!r} in {file_uri}",
        )


async def pptx_set_slide_bullets(
    file_uri: str,
    slide_index: int,
    bullets: list[str],
    theme: str = "corporate",
) -> dict[str, Any]:
    """Replace the bullet list on slide ``slide_index`` (0-based).

    Drops the old bullets first then writes ``bullets`` in their place,
    so you don't end up with old + new concatenated. Use this to fill
    body content on a slide that was created title-only, or to rewrite
    an existing list.

    Args:
        file_uri: Deck URI.
        slide_index: Zero-based index. Slide 0 is typically the cover.
        bullets: One string per bullet line, in display order. Pass an
            empty list to clear the bullets entirely.
        theme: Apply theme styling to the rewritten bullets.
    """
    async with _locks.for_uri(file_uri):
        original, _ = await _read(file_uri)
        updated = await asyncio.to_thread(
            pptx_engine.update_slide,
            original,
            index=slide_index,
            content=bullets,
            theme=theme,
        )
        return await _write(
            file_uri,
            updated,
            summary=(
                f"Set {len(bullets)} bullets on slide {slide_index} of {file_uri}"
            ),
        )


# ── deck inspection ───────────────────────────────────────────────────


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
for _fn in (
    pptx_create,
    pptx_add_title_slide,
    pptx_add_content_slide,
    pptx_add_two_content_slide,
    pptx_add_section_slide,
    pptx_set_slide_title,
    pptx_set_slide_bullets,
    pptx_read,
):
    register_tool(_fn)
