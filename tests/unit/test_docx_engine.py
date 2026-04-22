"""Unit tests for the ``python-docx`` engine."""

from __future__ import annotations

from io import BytesIO

import pytest
from docx import Document
from docx.oxml.ns import qn

from office_mcp.engines import docx_engine


def test_create_returns_valid_docx_bytes() -> None:
    data = docx_engine.create()

    # A .docx is a zip — starts with the PK signature.
    assert data.startswith(b"PK")
    # Must be re-openable by python-docx.
    doc = Document(BytesIO(data))
    assert doc.paragraphs is not None


def test_create_with_title_renders_as_heading() -> None:
    data = docx_engine.create(title="Quarterly report")
    doc = Document(BytesIO(data))

    heading_texts = [p.text for p in doc.paragraphs if p.style and "Title" in (p.style.name or "")]
    assert "Quarterly report" in heading_texts


def test_create_applies_default_font() -> None:
    data = docx_engine.create()
    doc = Document(BytesIO(data))

    normal = doc.styles["Normal"]
    assert normal.font.name == "Arial"
    assert normal.font.size.pt == 11


def test_read_returns_structured_paragraphs_and_tables() -> None:
    data = docx_engine.create(title="Report")
    data = docx_engine.append_section(
        data, heading="Intro", paragraphs=["hello world", "second line"], level=1
    )
    data = docx_engine.insert_table(
        data, headers=["Col A", "Col B"], rows=[["1", "2"], ["3", "4"]]
    )

    structure = docx_engine.read(data)

    headings = [p for p in structure["paragraphs"] if p["type"] == "heading"]
    paragraphs = [p for p in structure["paragraphs"] if p["type"] == "paragraph"]
    assert any(h["text"] == "Intro" and h["level"] == 1 for h in headings)
    assert any(p["text"] == "hello world" for p in paragraphs)
    assert structure["table_count"] == 1
    assert structure["tables"][0][0] == ["Col A", "Col B"]
    assert structure["tables"][0][1] == ["1", "2"]


def test_append_section_adds_heading_and_paragraphs() -> None:
    data = docx_engine.create()
    data = docx_engine.append_section(
        data, heading="Methods", paragraphs=["para one", "para two"], level=2
    )

    doc = Document(BytesIO(data))
    texts = [p.text for p in doc.paragraphs]
    assert "Methods" in texts
    assert "para one" in texts
    assert "para two" in texts


def test_append_section_rejects_invalid_level() -> None:
    data = docx_engine.create()
    with pytest.raises(ValueError, match="heading level"):
        docx_engine.append_section(
            data, heading="Bad", paragraphs=["x"], level=9
        )


def test_insert_table_bolds_header_row() -> None:
    data = docx_engine.create()
    data = docx_engine.insert_table(
        data, headers=["Name", "Amount"], rows=[["A", "10"]]
    )
    doc = Document(BytesIO(data))
    table = doc.tables[0]

    for cell in table.rows[0].cells:
        # At least one run in the header cell must be bold.
        bold_runs = [
            run for para in cell.paragraphs for run in para.runs if run.bold
        ]
        assert bold_runs, f"header cell {cell.text!r} has no bold run"


def test_insert_table_sets_dxa_widths_on_table_and_cells() -> None:
    data = docx_engine.create()
    data = docx_engine.insert_table(
        data, headers=["A", "B", "C"], rows=[["1", "2", "3"]]
    )
    doc = Document(BytesIO(data))
    table = doc.tables[0]

    tbl_w = table._tbl.tblPr.find(qn("w:tblW"))
    assert tbl_w is not None
    assert tbl_w.get(qn("w:type")) == "dxa"
    assert int(tbl_w.get(qn("w:w"))) == 9360

    expected_col = 9360 // 3
    for row in table.rows:
        for cell in row.cells:
            tc_w = cell._tc.tcPr.find(qn("w:tcW"))
            assert tc_w is not None, f"cell {cell.text!r} missing tcW"
            assert tc_w.get(qn("w:type")) == "dxa"
            assert int(tc_w.get(qn("w:w"))) == expected_col


def test_insert_table_rejects_row_with_wrong_column_count() -> None:
    data = docx_engine.create()
    with pytest.raises(ValueError, match="row 0"):
        docx_engine.insert_table(
            data, headers=["A", "B"], rows=[["only one"]]
        )


def test_insert_table_rejects_empty_headers() -> None:
    data = docx_engine.create()
    with pytest.raises(ValueError, match="headers"):
        docx_engine.insert_table(data, headers=[], rows=[])
