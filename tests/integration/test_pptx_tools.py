"""Integration tests for the layout-specific pptx tools."""

from __future__ import annotations

import asyncio
from collections.abc import Iterator
from pathlib import Path

import pytest

from office_mcp.context import UserContext, bind_user
from office_mcp.runtime import set_registry
from office_mcp.storage import LocalBackend, StorageRegistry
from office_mcp.tools.pptx_tools import (
    pptx_add_content_slide,
    pptx_add_section_slide,
    pptx_add_title_slide,
    pptx_add_two_content_slide,
    pptx_create,
    pptx_read,
    pptx_set_slide_bullets,
    pptx_set_slide_title,
)


@pytest.fixture
def registry(tmp_path: Path) -> Iterator[StorageRegistry]:
    reg = StorageRegistry()
    reg.register(LocalBackend(root=tmp_path))
    set_registry(reg)
    yield reg


@pytest.fixture
def yacine() -> Iterator[UserContext]:
    ctx = UserContext(user_id="yacine", email="yacine@netwatch.fr")
    with bind_user(ctx):
        yield ctx


@pytest.mark.asyncio
async def test_pptx_create_writes_to_storage_backend(
    registry: StorageRegistry, tmp_path: Path, yacine: UserContext
) -> None:
    result = await pptx_create(
        "local://decks/pitch.pptx",
        theme="cyberpunk",
        title="Demo pitch",
        subtitle="Phase 5",
    )

    assert result["file_uri"] == "local://decks/pitch.pptx"
    assert result["theme"] == "cyberpunk"
    assert (tmp_path / yacine.user_id / "decks" / "pitch.pptx").exists()


@pytest.mark.asyncio
async def test_full_deck_roundtrip_uses_each_layout_tool(
    registry: StorageRegistry, yacine: UserContext
) -> None:
    """One deck, one slide per layout-specific tool, then read it back."""
    await pptx_create("local://d.pptx", title="Deck")
    await pptx_add_title_slide(
        "local://d.pptx", title="Cover", subtitle="Q3 review"
    )
    await pptx_add_content_slide(
        "local://d.pptx",
        title="Highlights",
        bullets=["Revenue +18% QoQ", "EU expansion launched"],
    )
    await pptx_add_two_content_slide(
        "local://d.pptx",
        title="Pros vs Cons",
        left_bullets=["Fast"],
        right_bullets=["Fragile"],
    )
    await pptx_add_section_slide("local://d.pptx", title="Closing")

    payload = await pptx_read("local://d.pptx")
    assert payload["slide_count"] == 5
    titles = [s["title"] for s in payload["slides"]]
    assert titles == ["Deck", "Cover", "Highlights", "Pros vs Cons", "Closing"]
    assert "Revenue +18% QoQ" in payload["slides"][2]["body"]
    assert "Fast" in payload["slides"][3]["body"]
    assert "Fragile" in payload["slides"][3]["body"]


@pytest.mark.asyncio
async def test_set_slide_bullets_fills_a_title_only_slide(
    registry: StorageRegistry, yacine: UserContext
) -> None:
    """Reproduces the demo bug: a deck created with bullet-less slides
    can be filled in afterwards via pptx_set_slide_bullets."""
    await pptx_create("local://d.pptx", title="Deck")
    await pptx_add_content_slide(
        "local://d.pptx", title="Empty for now", bullets=["placeholder"]
    )

    await pptx_set_slide_bullets(
        "local://d.pptx",
        slide_index=1,
        bullets=["Real point 1", "Real point 2", "Real point 3"],
    )

    payload = await pptx_read("local://d.pptx")
    assert payload["slides"][1]["body"] == [
        "Real point 1",
        "Real point 2",
        "Real point 3",
    ]


@pytest.mark.asyncio
async def test_set_slide_title_rewrites_only_the_title(
    registry: StorageRegistry, yacine: UserContext
) -> None:
    await pptx_create("local://d.pptx", title="Deck")
    await pptx_add_content_slide(
        "local://d.pptx", title="Old", bullets=["keep me"]
    )

    await pptx_set_slide_title("local://d.pptx", slide_index=1, title="New")

    payload = await pptx_read("local://d.pptx")
    assert payload["slides"][1]["title"] == "New"
    assert payload["slides"][1]["body"] == ["keep me"]


@pytest.mark.asyncio
async def test_concurrent_set_calls_are_serialised_per_file(
    registry: StorageRegistry, tmp_path: Path, yacine: UserContext
) -> None:
    """Pre-fix, firing several pptx_set_* calls in parallel against the
    same file produced lost updates because each read started from the
    same baseline before the others' writes landed. The per-URI lock
    forces them to apply sequentially, so the final deck has *all*
    edits — not just the last writer's."""
    await pptx_create("local://d.pptx", title="Deck")
    # Three placeholder slides we'll overwrite concurrently.
    for i in range(3):
        await pptx_add_content_slide(
            "local://d.pptx", title=f"slot-{i}", bullets=["x"]
        )

    await asyncio.gather(
        pptx_set_slide_bullets(
            "local://d.pptx", slide_index=1, bullets=["one"]
        ),
        pptx_set_slide_bullets(
            "local://d.pptx", slide_index=2, bullets=["two"]
        ),
        pptx_set_slide_bullets(
            "local://d.pptx", slide_index=3, bullets=["three"]
        ),
    )

    payload = await pptx_read("local://d.pptx")
    # All three slot updates must have survived.
    bodies = [payload["slides"][i]["body"] for i in (1, 2, 3)]
    assert ["one"] in bodies
    assert ["two"] in bodies
    assert ["three"] in bodies


@pytest.mark.asyncio
async def test_pptx_tools_require_user_context(
    registry: StorageRegistry,
) -> None:
    with pytest.raises(RuntimeError, match="No UserContext bound"):
        await pptx_create("local://x.pptx")
