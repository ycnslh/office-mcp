"""Integration tests for the docx tools.

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
    _swap_extension,
    docx_append,
    docx_create,
    docx_read,
    docx_remove,
    docx_replace,
    docx_to_pdf,
)
from office_mcp.utils import libreoffice

_NEEDS_SOFFICE = pytest.mark.skipif(
    not libreoffice.is_available(),
    reason="LibreOffice headless is not installed in this environment",
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
async def test_docx_create_with_blocks_produces_full_document(
    registry: StorageRegistry, yacine: UserContext
) -> None:
    result = await docx_create(
        "local://full.docx",
        title="Histoire du Dragon",
        blocks=[
            {"type": "heading", "level": 1, "text": "Le Dernier Dragon"},
            {"type": "paragraph", "text": "Pyralis vivait depuis…"},
            {"type": "heading", "level": 1, "text": "La Rencontre"},
            {"type": "paragraph", "text": "Élina arriva un soir d'orage."},
            {
                "type": "table",
                "headers": ["Personnage", "Rôle"],
                "rows": [["Pyralis", "Dragon"], ["Élina", "Héroïne"]],
            },
        ],
    )

    assert "5 block(s)" in result["summary"]
    payload = await docx_read("local://full.docx")
    texts = [p["text"] for p in payload["paragraphs"]]
    assert "Le Dernier Dragon" in texts
    assert "La Rencontre" in texts
    assert payload["table_count"] == 1
    assert payload["tables"][0][1] == ["Pyralis", "Dragon"]


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
async def test_docx_append_persists_multiple_blocks(
    registry: StorageRegistry, yacine: UserContext
) -> None:
    await docx_create("local://doc.docx", title="Doc")

    result = await docx_append(
        "local://doc.docx",
        blocks=[
            {"type": "heading", "level": 1, "text": "Intro"},
            {"type": "paragraph", "text": "lead paragraph"},
            {"type": "paragraph", "text": "supporting detail"},
            {
                "type": "table",
                "headers": ["Product", "Units"],
                "rows": [["Alpha", "12"]],
            },
        ],
    )
    assert "4 block(s)" in result["summary"]

    after = await docx_read("local://doc.docx")
    texts = [p["text"] for p in after["paragraphs"]]
    assert "Intro" in texts
    assert "lead paragraph" in texts
    assert after["table_count"] == 1
    assert after["tables"][0][0] == ["Product", "Units"]


@pytest.mark.asyncio
async def test_docx_replace_substitutes_and_returns_count(
    registry: StorageRegistry, yacine: UserContext
) -> None:
    await docx_create(
        "local://template.docx",
        title="Hello {name}",
        blocks=[
            {"type": "paragraph", "text": "Dear {name}, your code is {code}."},
        ],
    )

    result = await docx_replace(
        "local://template.docx",
        replacements=[
            {"find": "{name}", "replace": "Yacine"},
            {"find": "{code}", "replace": "42"},
        ],
    )

    assert result["replacements_made"] == 3  # title + paragraph (×2 placeholders)
    after = await docx_read("local://template.docx")
    flat = " | ".join(p["text"] for p in after["paragraphs"])
    assert "Yacine" in flat
    assert "{name}" not in flat
    assert "42" in flat


@pytest.mark.asyncio
async def test_docx_remove_drops_targeted_section(
    registry: StorageRegistry, yacine: UserContext
) -> None:
    await docx_create(
        "local://multi.docx",
        title="Doc",
        blocks=[
            {"type": "heading", "level": 1, "text": "Keep"},
            {"type": "paragraph", "text": "kept body"},
            {"type": "heading", "level": 1, "text": "Drop"},
            {"type": "paragraph", "text": "dropped body"},
            {"type": "heading", "level": 1, "text": "Also kept"},
        ],
    )

    result = await docx_remove("local://multi.docx", headings=["Drop"])

    assert result["sections_removed"] == 1
    after = await docx_read("local://multi.docx")
    texts = [p["text"] for p in after["paragraphs"]]
    assert "Keep" in texts
    assert "Drop" not in texts
    assert "dropped body" not in texts
    assert "Also kept" in texts


@pytest.mark.asyncio
async def test_docx_tools_require_user_context(registry: StorageRegistry) -> None:
    # No bind_user here — require_current_user must refuse.
    with pytest.raises(RuntimeError, match="No UserContext bound"):
        await docx_create("local://x.docx")


@pytest.mark.parametrize(
    ("uri", "expected"),
    [
        ("local://reports/q3.docx", "local://reports/q3.pdf"),
        ("local://no-ext", "local://no-ext.pdf"),
        ("local://path.with.dots/file.docx", "local://path.with.dots/file.pdf"),
    ],
)
def test_swap_extension(uri: str, expected: str) -> None:
    assert _swap_extension(uri, ".pdf") == expected


@pytest.mark.asyncio
@_NEEDS_SOFFICE
async def test_docx_to_pdf_writes_valid_pdf(
    registry: StorageRegistry, tmp_path: Path, yacine: UserContext
) -> None:
    await docx_create(
        "local://reports/q3.docx",
        title="Q3 report",
        blocks=[
            {"type": "heading", "level": 1, "text": "Intro"},
            {"type": "paragraph", "text": "Quarterly highlights."},
        ],
    )

    result = await docx_to_pdf("local://reports/q3.docx")

    assert result["file_uri"] == "local://reports/q3.pdf"
    assert result["source_uri"] == "local://reports/q3.docx"
    pdf_path = tmp_path / yacine.user_id / "reports" / "q3.pdf"
    assert pdf_path.exists()
    assert pdf_path.read_bytes().startswith(b"%PDF-")


@pytest.mark.asyncio
@_NEEDS_SOFFICE
async def test_docx_to_pdf_custom_output_uri(
    registry: StorageRegistry, tmp_path: Path, yacine: UserContext
) -> None:
    await docx_create("local://in.docx", title="Hello")

    result = await docx_to_pdf(
        "local://in.docx", output_uri="local://out/renamed.pdf"
    )

    assert result["file_uri"] == "local://out/renamed.pdf"
    assert (tmp_path / yacine.user_id / "out" / "renamed.pdf").exists()
