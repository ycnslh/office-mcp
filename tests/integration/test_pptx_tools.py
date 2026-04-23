"""Integration tests for the 3 core pptx tools."""

from __future__ import annotations

from collections.abc import Iterator
from pathlib import Path

import pytest

from office_mcp.context import UserContext, bind_user
from office_mcp.runtime import set_registry
from office_mcp.storage import LocalBackend, StorageRegistry
from office_mcp.tools.pptx_tools import pptx_add_slide, pptx_create, pptx_read


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
async def test_pptx_add_slide_then_read_roundtrip(
    registry: StorageRegistry, yacine: UserContext
) -> None:
    await pptx_create("local://d.pptx", title="Deck")
    await pptx_add_slide(
        "local://d.pptx",
        title="Highlights",
        layout="content",
        content=["Revenue +18% QoQ", "EU expansion launched"],
    )
    await pptx_add_slide(
        "local://d.pptx",
        title="Pros vs Cons",
        layout="two-content",
        content=["Fast"],
        content_right=["Fragile"],
    )

    payload = await pptx_read("local://d.pptx")
    assert payload["slide_count"] == 3
    titles = [s["title"] for s in payload["slides"]]
    assert titles == ["Deck", "Highlights", "Pros vs Cons"]
    assert "Revenue +18% QoQ" in payload["slides"][1]["body"]
    assert "Fast" in payload["slides"][2]["body"]
    assert "Fragile" in payload["slides"][2]["body"]


@pytest.mark.asyncio
async def test_pptx_tools_require_user_context(
    registry: StorageRegistry,
) -> None:
    with pytest.raises(RuntimeError, match="No UserContext bound"):
        await pptx_create("local://x.pptx")
