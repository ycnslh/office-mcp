"""Integration tests for the pptx tools."""

from __future__ import annotations

import asyncio
from collections.abc import Iterator
from pathlib import Path

import pytest

from office_mcp.context import UserContext, bind_user
from office_mcp.runtime import set_registry
from office_mcp.storage import LocalBackend, StorageRegistry
from office_mcp.tools.pptx_tools import (
    pptx_append,
    pptx_create,
    pptx_read,
    pptx_replace,
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
        slides=[
            {"type": "title", "title": "Demo pitch", "subtitle": "Phase 5"}
        ],
    )

    assert result["file_uri"] == "local://decks/pitch.pptx"
    assert result["theme"] == "cyberpunk"
    assert result["slide_count"] == 1
    assert (tmp_path / yacine.user_id / "decks" / "pitch.pptx").exists()


@pytest.mark.asyncio
async def test_pptx_create_full_deck_one_call(
    registry: StorageRegistry, yacine: UserContext
) -> None:
    """One deck assembled in a single call covering all slide flavours."""
    await pptx_create(
        "local://d.pptx",
        slides=[
            {"type": "title", "title": "Deck"},
            {"type": "title", "title": "Cover", "subtitle": "Q3 review"},
            {
                "type": "content",
                "title": "Highlights",
                "bullets": ["Revenue +18% QoQ", "EU expansion launched"],
            },
            {
                "type": "two-content",
                "title": "Pros vs Cons",
                "left": ["Fast"],
                "right": ["Fragile"],
            },
            {"type": "section", "title": "Closing"},
        ],
    )

    payload = await pptx_read("local://d.pptx")
    assert payload["slide_count"] == 5
    titles = [s["title"] for s in payload["slides"]]
    assert titles == ["Deck", "Cover", "Highlights", "Pros vs Cons", "Closing"]
    assert "Revenue +18% QoQ" in payload["slides"][2]["body"]
    assert "Fast" in payload["slides"][3]["body"]
    assert "Fragile" in payload["slides"][3]["body"]


@pytest.mark.asyncio
async def test_pptx_append_extends_existing_deck(
    registry: StorageRegistry, yacine: UserContext
) -> None:
    await pptx_create(
        "local://d.pptx", slides=[{"type": "title", "title": "Deck"}]
    )

    result = await pptx_append(
        "local://d.pptx",
        slides=[
            {"type": "section", "title": "Part II"},
            {
                "type": "content",
                "title": "Bullets",
                "bullets": ["A", "B", "C"],
            },
        ],
    )
    assert "2 slide(s)" in result["summary"]

    payload = await pptx_read("local://d.pptx")
    titles = [s["title"] for s in payload["slides"]]
    assert titles == ["Deck", "Part II", "Bullets"]


@pytest.mark.asyncio
async def test_pptx_replace_patches_existing_slides(
    registry: StorageRegistry, yacine: UserContext
) -> None:
    await pptx_create(
        "local://d.pptx",
        slides=[
            {"type": "title", "title": "Deck"},
            {"type": "content", "title": "Old", "bullets": ["keep"]},
        ],
    )

    await pptx_replace(
        "local://d.pptx",
        updates=[
            {
                "slide_index": 1,
                "title": "New",
                "bullets": ["new A", "new B"],
            }
        ],
    )

    payload = await pptx_read("local://d.pptx")
    assert payload["slides"][1]["title"] == "New"
    assert payload["slides"][1]["body"] == ["new A", "new B"]


@pytest.mark.asyncio
async def test_pptx_concurrent_replaces_are_serialised_per_file(
    registry: StorageRegistry, tmp_path: Path, yacine: UserContext
) -> None:
    """Several pptx_replace calls fired in parallel against the same
    file used to lose updates because each read started from the same
    baseline before the others' writes landed. The per-URI lock forces
    them to apply sequentially."""
    await pptx_create(
        "local://d.pptx",
        slides=[
            {"type": "title", "title": "Deck"},
            {"type": "content", "title": "slot-1", "bullets": ["x"]},
            {"type": "content", "title": "slot-2", "bullets": ["x"]},
            {"type": "content", "title": "slot-3", "bullets": ["x"]},
        ],
    )

    await asyncio.gather(
        pptx_replace(
            "local://d.pptx",
            updates=[{"slide_index": 1, "bullets": ["one"]}],
        ),
        pptx_replace(
            "local://d.pptx",
            updates=[{"slide_index": 2, "bullets": ["two"]}],
        ),
        pptx_replace(
            "local://d.pptx",
            updates=[{"slide_index": 3, "bullets": ["three"]}],
        ),
    )

    payload = await pptx_read("local://d.pptx")
    bodies = [payload["slides"][i]["body"] for i in (1, 2, 3)]
    assert ["one"] in bodies
    assert ["two"] in bodies
    assert ["three"] in bodies


@pytest.mark.asyncio
async def test_pptx_tools_require_user_context(
    registry: StorageRegistry,
) -> None:
    with pytest.raises(RuntimeError, match="No UserContext bound"):
        await pptx_create("local://x.pptx", slides=[])
