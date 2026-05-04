"""Integration tests for the xlsx tools."""

from __future__ import annotations

from collections.abc import Iterator
from pathlib import Path

import pytest

from office_mcp.context import UserContext, bind_user
from office_mcp.runtime import set_registry
from office_mcp.storage import LocalBackend, StorageRegistry
from office_mcp.tools.xlsx_tools import (
    xlsx_create,
    xlsx_read_range,
    xlsx_update,
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
        "local://budget/2026.xlsx",
        sheets=[{"name": "Inputs"}, {"name": "Calc"}],
        recalc=False,
    )

    assert result["file_uri"] == "local://budget/2026.xlsx"
    assert result["sheets"] == ["Inputs", "Calc"]
    assert (tmp_path / yacine.user_id / "budget" / "2026.xlsx").exists()


@pytest.mark.asyncio
async def test_xlsx_create_full_workbook_with_writes_and_formulas(
    registry: StorageRegistry, yacine: UserContext
) -> None:
    result = await xlsx_create(
        "local://full.xlsx",
        sheets=[
            {
                "name": "Inputs",
                "writes": [
                    {
                        "start_cell": "A1",
                        "values": [
                            ["Region", "Units", "Price"],
                            ["EU", 1200, 49.9],
                            ["US", 900, 52.0],
                        ],
                        "style": "input",
                    }
                ],
                "formulas": [
                    {"cell": "D2", "formula": "=B2*C2"},
                    {"cell": "D3", "formula": "=B3*C3"},
                    {"cell": "D4", "formula": "=SUM(D2:D3)"},
                ],
            }
        ],
        recalc=False,
    )

    assert result["sheets"] == ["Inputs"]
    assert "Created workbook" in result["summary"]

    payload = await xlsx_read_range(
        "local://full.xlsx", sheet="Inputs", range_str="A1:C3"
    )
    assert payload["values"][0] == ["Region", "Units", "Price"]
    assert payload["values"][1] == ["EU", 1200, 49.9]


@pytest.mark.asyncio
async def test_xlsx_update_appends_to_existing_workbook(
    registry: StorageRegistry, yacine: UserContext
) -> None:
    await xlsx_create("local://m.xlsx", sheets=[{"name": "S"}], recalc=False)
    result = await xlsx_update(
        "local://m.xlsx",
        sheets=[
            {
                "name": "S",
                "writes": [
                    {
                        "start_cell": "A1",
                        "values": [["Item", "Price"], ["A", 100], ["B", 250]],
                    }
                ],
            }
        ],
        recalc=False,
    )

    assert "1 write(s) and 0 formula(s)" in result["summary"]
    payload = await xlsx_read_range(
        "local://m.xlsx", sheet="S", range_str="A1:B3"
    )
    assert payload["values"] == [
        ["Item", "Price"],
        ["A", 100],
        ["B", 250],
    ]


@pytest.mark.asyncio
async def test_xlsx_update_creates_missing_sheet(
    registry: StorageRegistry, yacine: UserContext
) -> None:
    await xlsx_create(
        "local://m2.xlsx", sheets=[{"name": "Original"}], recalc=False
    )
    await xlsx_update(
        "local://m2.xlsx",
        sheets=[
            {
                "name": "Brand New",
                "writes": [{"start_cell": "A1", "values": [["x"]]}],
            }
        ],
        recalc=False,
    )

    payload = await xlsx_read_range(
        "local://m2.xlsx", sheet="Brand New", range_str="A1"
    )
    assert payload["values"] == [["x"]]


@pytest.mark.asyncio
@_NEEDS_SOFFICE
async def test_xlsx_create_with_recalc_returns_computed_values(
    registry: StorageRegistry, yacine: UserContext
) -> None:
    await xlsx_create(
        "local://r.xlsx",
        sheets=[
            {
                "name": "S",
                "writes": [
                    {"start_cell": "A1", "values": [[10], [20], [30]]}
                ],
                "formulas": [{"cell": "A4", "formula": "=SUM(A1:A3)"}],
            }
        ],
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
