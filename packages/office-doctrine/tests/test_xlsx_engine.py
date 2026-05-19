"""Unit tests for the ``openpyxl`` engine."""

from __future__ import annotations

from io import BytesIO

import openpyxl
import pytest

from office_doctrine import _libreoffice as libreoffice
from office_doctrine import xlsx as xlsx_engine

_NEEDS_SOFFICE = pytest.mark.skipif(
    not libreoffice.is_available(),
    reason="LibreOffice headless is not installed in this environment",
)


def test_compose_returns_valid_xlsx_with_default_sheet() -> None:
    data = xlsx_engine.compose()

    assert data.startswith(b"PK")
    assert xlsx_engine.list_sheets(data) == ["Sheet"]


def test_compose_with_named_sheets() -> None:
    data = xlsx_engine.compose(
        sheets=[{"name": "Inputs"}, {"name": "Calc"}, {"name": "Outputs"}],
        recalc=False,
    )

    assert xlsx_engine.list_sheets(data) == ["Inputs", "Calc", "Outputs"]


def test_compose_writes_values_then_round_trip_via_read_range() -> None:
    data = xlsx_engine.compose(
        sheets=[
            {
                "name": "Sheet1",
                "writes": [
                    {
                        "start_cell": "B2",
                        "values": [["Region", "Revenue"], ["EU", 1200], ["US", 900]],
                        "style": "input",
                    }
                ],
            }
        ],
        recalc=False,
    )

    values = xlsx_engine.read_range(data, sheet="Sheet1", range_str="B2:C4")
    assert values == [["Region", "Revenue"], ["EU", 1200], ["US", 900]]


def test_compose_applies_input_color() -> None:
    data = xlsx_engine.compose(
        sheets=[
            {
                "name": "S",
                "writes": [
                    {"start_cell": "A1", "values": [[42]], "style": "input"}
                ],
            }
        ],
        recalc=False,
    )

    wb = openpyxl.load_workbook(BytesIO(data))
    assert wb["S"]["A1"].font.color.rgb.endswith(xlsx_engine.INPUT_COLOR)


def test_compose_with_formula_no_recalc_writes_formula_string() -> None:
    data = xlsx_engine.compose(
        sheets=[
            {
                "name": "S",
                "writes": [
                    {
                        "start_cell": "A1",
                        "values": [[10], [20], [30]],
                        "style": "input",
                    }
                ],
                "formulas": [{"cell": "A4", "formula": "=SUM(A1:A3)"}],
            }
        ],
        recalc=False,
    )

    wb = openpyxl.load_workbook(BytesIO(data))
    assert wb["S"]["A4"].value == "=SUM(A1:A3)"
    assert wb["S"]["A4"].font.color.rgb.endswith(xlsx_engine.FORMULA_COLOR)


def test_compose_normalises_leading_equals_on_formulas() -> None:
    data = xlsx_engine.compose(
        sheets=[
            {
                "name": "S",
                "formulas": [{"cell": "A1", "formula": "SUM(B1:B3)"}],
            }
        ],
        recalc=False,
    )
    wb = openpyxl.load_workbook(BytesIO(data))
    assert wb["S"]["A1"].value == "=SUM(B1:B3)"


def test_compose_uses_link_color_for_cross_sheet_reference() -> None:
    data = xlsx_engine.compose(
        sheets=[
            {
                "name": "Inputs",
                "writes": [
                    {"start_cell": "B3", "values": [[42]], "style": "input"}
                ],
            },
            {
                "name": "Calc",
                "formulas": [{"cell": "A1", "formula": "=Inputs!B3"}],
            },
        ],
        recalc=False,
    )

    wb = openpyxl.load_workbook(BytesIO(data))
    assert wb["Calc"]["A1"].font.color.rgb.endswith(xlsx_engine.LINK_COLOR)


def test_compose_rejects_duplicate_sheet_names() -> None:
    with pytest.raises(ValueError, match="duplicate"):
        xlsx_engine.compose(
            sheets=[{"name": "S"}, {"name": "S"}], recalc=False
        )


def test_compose_rejects_empty_sheet_name() -> None:
    with pytest.raises(ValueError, match="'name'"):
        xlsx_engine.compose(sheets=[{"name": ""}], recalc=False)


def test_compose_rejects_unknown_style() -> None:
    with pytest.raises(ValueError, match="unknown style"):
        xlsx_engine.compose(
            sheets=[
                {
                    "name": "S",
                    "writes": [
                        {"start_cell": "A1", "values": [[1]], "style": "neon"}
                    ],
                }
            ],
            recalc=False,
        )


def test_compose_rejects_empty_values() -> None:
    with pytest.raises(ValueError, match="'values'"):
        xlsx_engine.compose(
            sheets=[
                {"name": "S", "writes": [{"start_cell": "A1", "values": []}]}
            ],
            recalc=False,
        )


def test_apply_updates_creates_missing_sheet() -> None:
    data = xlsx_engine.compose(sheets=[{"name": "Original"}], recalc=False)
    data = xlsx_engine.apply_updates(
        data,
        sheets=[
            {
                "name": "NewSheet",
                "writes": [{"start_cell": "A1", "values": [["x"]]}],
            }
        ],
        recalc=False,
    )

    assert "NewSheet" in xlsx_engine.list_sheets(data)


def test_apply_updates_appends_writes_and_formulas() -> None:
    data = xlsx_engine.compose(
        sheets=[{"name": "S"}], recalc=False
    )
    data = xlsx_engine.apply_updates(
        data,
        sheets=[
            {
                "name": "S",
                "writes": [
                    {"start_cell": "A1", "values": [[10], [20], [30]]}
                ],
                "formulas": [{"cell": "A4", "formula": "=SUM(A1:A3)"}],
            }
        ],
        recalc=False,
    )

    wb = openpyxl.load_workbook(BytesIO(data))
    assert wb["S"]["A1"].value == 10
    assert wb["S"]["A4"].value == "=SUM(A1:A3)"


def test_read_range_unknown_sheet_raises() -> None:
    data = xlsx_engine.compose(sheets=[{"name": "Sheet1"}], recalc=False)
    with pytest.raises(KeyError, match="ghost"):
        xlsx_engine.read_range(data, sheet="ghost", range_str="A1")


def test_read_range_single_cell_returns_2d() -> None:
    data = xlsx_engine.compose(
        sheets=[
            {
                "name": "S",
                "writes": [{"start_cell": "A1", "values": [["hello"]]}],
            }
        ],
        recalc=False,
    )
    assert xlsx_engine.read_range(data, sheet="S", range_str="A1") == [["hello"]]


# ── Recalc-dependent tests (require LibreOffice headless) ────────────


@_NEEDS_SOFFICE
def test_compose_with_recalc_bakes_computed_values() -> None:
    data = xlsx_engine.compose(
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

    values = xlsx_engine.read_range(data, sheet="S", range_str="A4")
    assert values == [[60]]
