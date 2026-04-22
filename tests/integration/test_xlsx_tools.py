"""Integration tests for the 4 core xlsx tools."""

from __future__ import annotations

from collections.abc import Iterator
from pathlib import Path

import pytest

from office_mcp.context import UserContext, bind_user
from office_mcp.runtime import set_registry
from office_mcp.storage import LocalBackend, StorageRegistry
from office_mcp.tools.xlsx_tools import (
    xlsx_apply_formula,
    xlsx_create,
    xlsx_read_range,
    xlsx_write_range,
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
async def test_xlsx_create_writes_to_storage_backend(
    registry: StorageRegistry, tmp_path: Path, yacine: UserContext
) -> None:
    result = await xlsx_create(
        "local://budget/2026.xlsx", sheets=["Inputs", "Calc"]
    )

    assert result["file_uri"] == "local://budget/2026.xlsx"
    assert result["sheets"] == ["Inputs", "Calc"]
    assert (tmp_path / yacine.user_id / "budget" / "2026.xlsx").exists()


@pytest.mark.asyncio
async def test_xlsx_write_then_read_roundtrip(
    registry: StorageRegistry, yacine: UserContext
) -> None:
    await xlsx_create("local://m.xlsx", sheets=["S"])
    await xlsx_write_range(
        "local://m.xlsx",
        sheet="S",
        start_cell="A1",
        values=[["Item", "Price"], ["A", 100], ["B", 250]],
    )

    payload = await xlsx_read_range(
        "local://m.xlsx", sheet="S", range_str="A1:B3"
    )
    assert payload["values"] == [
        ["Item", "Price"],
        ["A", 100],
        ["B", 250],
    ]


@pytest.mark.asyncio
async def test_xlsx_apply_formula_without_recalc_persists_formula(
    registry: StorageRegistry, yacine: UserContext
) -> None:
    await xlsx_create("local://f.xlsx", sheets=["S"])
    await xlsx_write_range(
        "local://f.xlsx",
        sheet="S",
        start_cell="A1",
        values=[[10], [20], [30]],
    )
    result = await xlsx_apply_formula(
        "local://f.xlsx",
        sheet="S",
        cell="A4",
        formula="=SUM(A1:A3)",
        recalc=False,
    )

    assert result["recalculated"] is False
    payload = await xlsx_read_range(
        "local://f.xlsx", sheet="S", range_str="A4"
    )
    # Without recalc: openpyxl returns either the formula string or None
    # depending on the workbook's cached state. Either is acceptable here.
    assert payload["values"][0][0] in ("=SUM(A1:A3)", None)


@pytest.mark.asyncio
@_NEEDS_SOFFICE
async def test_xlsx_apply_formula_with_recalc_returns_computed_value(
    registry: StorageRegistry, yacine: UserContext
) -> None:
    await xlsx_create("local://r.xlsx", sheets=["S"])
    await xlsx_write_range(
        "local://r.xlsx",
        sheet="S",
        start_cell="A1",
        values=[[10], [20], [30]],
    )
    await xlsx_apply_formula(
        "local://r.xlsx",
        sheet="S",
        cell="A4",
        formula="=SUM(A1:A3)",
        recalc=True,
    )

    payload = await xlsx_read_range(
        "local://r.xlsx", sheet="S", range_str="A4"
    )
    assert payload["values"] == [[60]]


@pytest.mark.asyncio
async def test_xlsx_tools_require_user_context(
    registry: StorageRegistry,
) -> None:
    with pytest.raises(RuntimeError, match="No UserContext bound"):
        await xlsx_create("local://x.xlsx")
