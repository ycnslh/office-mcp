"""Integration tests for the 4 core docx tools.

Tools are invoked directly (they remain plain callables after the
``@mcp.tool`` decorator) with a test-scoped registry and user binding.
"""

from __future__ import annotations

from collections.abc import Iterator
from pathlib import Path

import pytest

from office_mcp.context import UserContext, bind_user
from office_mcp.runtime import set_registry
from office_mcp.storage import LocalBackend, StorageRegistry
from office_mcp.tools.docx_tools import (
    docx_append_section,
    docx_create,
    docx_insert_table,
    docx_read,
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
async def test_docx_create_writes_to_storage_backend(
    registry: StorageRegistry, tmp_path: Path, yacine: UserContext
) -> None:
    result = await docx_create("local://reports/q3.docx", title="Q3 report")

    assert result["file_uri"] == "local://reports/q3.docx"
    assert result["size_bytes"] > 0
    assert result["etag"] is not None
    assert "Q3 report" in result["summary"]
    assert (tmp_path / yacine.user_id / "reports" / "q3.docx").exists()


@pytest.mark.asyncio
async def test_docx_create_then_read_roundtrip(
    registry: StorageRegistry, yacine: UserContext
) -> None:
    await docx_create("local://simple.docx", title="Hello world")
    payload = await docx_read("local://simple.docx")

    assert payload["file_uri"] == "local://simple.docx"
    assert payload["paragraph_count"] >= 1
    assert any("Hello world" in p["text"] for p in payload["paragraphs"])


@pytest.mark.asyncio
async def test_docx_append_section_persists(
    registry: StorageRegistry, yacine: UserContext
) -> None:
    await docx_create("local://doc.docx", title="Doc")

    first = await docx_append_section(
        "local://doc.docx",
        heading="Intro",
        paragraphs=["lead paragraph", "supporting detail"],
        level=1,
    )
    assert "Intro" in first["summary"]
    assert "2 paragraphs" in first["summary"]

    after = await docx_read("local://doc.docx")
    texts = [p["text"] for p in after["paragraphs"]]
    assert "Intro" in texts
    assert "lead paragraph" in texts
    assert "supporting detail" in texts


@pytest.mark.asyncio
async def test_docx_insert_table_persists(
    registry: StorageRegistry, yacine: UserContext
) -> None:
    await docx_create("local://table.docx")

    result = await docx_insert_table(
        "local://table.docx",
        headers=["Product", "Units", "Revenue"],
        rows=[["Alpha", "12", "1200"], ["Beta", "7", "910"]],
    )
    assert "2×3 table" in result["summary"]

    after = await docx_read("local://table.docx")
    assert after["table_count"] == 1
    assert after["tables"][0][0] == ["Product", "Units", "Revenue"]
    assert after["tables"][0][1] == ["Alpha", "12", "1200"]


@pytest.mark.asyncio
async def test_docx_tools_require_user_context(registry: StorageRegistry) -> None:
    # No bind_user here — require_current_user must refuse.
    with pytest.raises(RuntimeError, match="No UserContext bound"):
        await docx_create("local://x.docx")
