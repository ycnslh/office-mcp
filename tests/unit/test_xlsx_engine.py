"""Unit tests for the ``openpyxl`` engine."""

from __future__ import annotations

from io import BytesIO

import openpyxl
import pytest

from office_mcp.engines import xlsx_engine
from office_mcp.utils import libreoffice

_NEEDS_SOFFICE = pytest.mark.skipif(
    not libreoffice.is_available(),
    reason="LibreOffice headless is not installed in this environment",
)


def test_create_returns_valid_xlsx_with_default_sheet() -> None:
    data = xlsx_engine.create()

    assert data.startswith(b"PK")
    assert xlsx_engine.list_sheets(data) == ["Sheet"]


def test_create_with_named_sheets() -> None:
    data = xlsx_engine.create(sheets=["Inputs", "Calc", "Outputs"])

    assert xlsx_engine.list_sheets(data) == ["Inputs", "Calc", "Outputs"]


def test_write_range_then_read_range_roundtrip() -> None:
    data = xlsx_engine.create(sheets=["Sheet1"])
    data = xlsx_engine.write_range(
        data,
        sheet="Sheet1",
        start_cell="B2",
        values=[["Region", "Revenue"], ["EU", 1200], ["US", 900]],
    )

    values = xlsx_engine.read_range(data, sheet="Sheet1", range_str="B2:C4")
    assert values == [["Region", "Revenue"], ["EU", 1200], ["US", 900]]


def test_read_range_single_cell_returns_2d() -> None:
    data = xlsx_engine.create(sheets=["S"])
    data = xlsx_engine.write_range(
        data, sheet="S", start_cell="A1", values=[["hello"]]
    )

    assert xlsx_engine.read_range(data, sheet="S", range_str="A1") == [["hello"]]


def test_write_range_applies_input_color() -> None:
    data = xlsx_engine.create(sheets=["S"])
    data = xlsx_engine.write_range(
        data, sheet="S", start_cell="A1", values=[[42]], style="input"
    )

    wb = openpyxl.load_workbook(BytesIO(data))
    assert wb["S"]["A1"].font.color.rgb.endswith(xlsx_engine.INPUT_COLOR)


def test_write_range_creates_missing_sheet() -> None:
    data = xlsx_engine.create(sheets=["Original"])
    data = xlsx_engine.write_range(
        data, sheet="NewSheet", start_cell="A1", values=[["x"]]
    )

    assert "NewSheet" in xlsx_engine.list_sheets(data)


def test_apply_formula_without_recalc_writes_formula_string() -> None:
    data = xlsx_engine.create(sheets=["S"])
    data = xlsx_engine.write_range(
        data, sheet="S", start_cell="A1", values=[[10], [20], [30]]
    )
    data = xlsx_engine.apply_formula(
        data, sheet="S", cell="A4", formula="=SUM(A1:A3)", recalc=False
    )

    wb = openpyxl.load_workbook(BytesIO(data))
    assert wb["S"]["A4"].value == "=SUM(A1:A3)"
    assert wb["S"]["A4"].font.color.rgb.endswith(xlsx_engine.FORMULA_COLOR)


def test_apply_formula_normalises_leading_equals() -> None:
    data = xlsx_engine.create(sheets=["S"])
    data = xlsx_engine.apply_formula(
        data, sheet="S", cell="A1", formula="SUM(B1:B3)", recalc=False
    )
    wb = openpyxl.load_workbook(BytesIO(data))
    assert wb["S"]["A1"].value == "=SUM(B1:B3)"


def test_apply_formula_uses_link_color_for_cross_sheet_reference() -> None:
    data = xlsx_engine.create(sheets=["Inputs", "Calc"])
    data = xlsx_engine.write_range(
        data, sheet="Inputs", start_cell="B3", values=[[42]]
    )
    data = xlsx_engine.apply_formula(
        data, sheet="Calc", cell="A1", formula="=Inputs!B3", recalc=False
    )

    wb = openpyxl.load_workbook(BytesIO(data))
    assert wb["Calc"]["A1"].font.color.rgb.endswith(xlsx_engine.LINK_COLOR)


def test_add_sheet_appends_a_new_one() -> None:
    data = xlsx_engine.create(sheets=["Inputs"])
    data = xlsx_engine.add_sheet(data, name="Outputs")

    assert xlsx_engine.list_sheets(data) == ["Inputs", "Outputs"]


def test_add_sheet_rejects_duplicate_name() -> None:
    data = xlsx_engine.create(sheets=["Inputs"])
    with pytest.raises(ValueError, match="already exists"):
        xlsx_engine.add_sheet(data, name="Inputs")


def test_read_range_unknown_sheet_raises() -> None:
    data = xlsx_engine.create(sheets=["Sheet1"])
    with pytest.raises(KeyError, match="ghost"):
        xlsx_engine.read_range(data, sheet="ghost", range_str="A1")


def test_apply_formula_unknown_sheet_raises() -> None:
    data = xlsx_engine.create(sheets=["Sheet1"])
    with pytest.raises(KeyError, match="ghost"):
        xlsx_engine.apply_formula(
            data, sheet="ghost", cell="A1", formula="=1+1", recalc=False
        )


def test_write_range_rejects_empty_values() -> None:
    data = xlsx_engine.create(sheets=["S"])
    with pytest.raises(ValueError, match="values"):
        xlsx_engine.write_range(
            data, sheet="S", start_cell="A1", values=[]
        )


# ── Recalc-dependent tests (require LibreOffice headless) ────────────


@_NEEDS_SOFFICE
def test_apply_formula_with_recalc_bakes_computed_values() -> None:
    data = xlsx_engine.create(sheets=["S"])
    data = xlsx_engine.write_range(
        data, sheet="S", start_cell="A1", values=[[10], [20], [30]]
    )
    data = xlsx_engine.apply_formula(
        data, sheet="S", cell="A4", formula="=SUM(A1:A3)", recalc=True
    )

    values = xlsx_engine.read_range(data, sheet="S", range_str="A4")
    assert values == [[60]]
