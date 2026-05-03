"""Unit tests for the ``python-docx`` engine."""

from __future__ import annotations

from io import BytesIO

import pytest
from docx import Document
from docx.oxml.ns import qn

from office_mcp.engines import docx_engine


def test_compose_returns_valid_docx_bytes() -> None:
    data = docx_engine.compose()

    # A .docx is a zip — starts with the PK signature.
    assert data.startswith(b"PK")
    # Must be re-openable by python-docx.
    doc = Document(BytesIO(data))
    assert doc.paragraphs is not None


def test_compose_with_title_renders_as_word_title_style() -> None:
    data = docx_engine.compose(title="Quarterly report")
    doc = Document(BytesIO(data))

    title_paragraphs = [
        p for p in doc.paragraphs if p.style and "Title" in (p.style.name or "")
    ]
    assert any("Quarterly report" == p.text for p in title_paragraphs)


def test_compose_applies_default_font() -> None:
    data = docx_engine.compose()
    doc = Document(BytesIO(data))

    normal = doc.styles["Normal"]
    assert normal.font.name == "Arial"
    assert normal.font.size.pt == 11


def test_compose_renders_blocks_in_order() -> None:
    data = docx_engine.compose(
        title="Report",
        blocks=[
            {"type": "heading", "level": 1, "text": "Intro"},
            {"type": "paragraph", "text": "hello world"},
            {"type": "paragraph", "text": "second line"},
            {"type": "heading", "level": 2, "text": "Details"},
            {
                "type": "table",
                "headers": ["Col A", "Col B"],
                "rows": [["1", "2"], ["3", "4"]],
            },
        ],
    )
    structure = docx_engine.read(data)

    headings = [p for p in structure["paragraphs"] if p["type"] == "heading"]
    paragraphs = [p for p in structure["paragraphs"] if p["type"] == "paragraph"]
    assert any(h["text"] == "Intro" and h["level"] == 1 for h in headings)
    assert any(h["text"] == "Details" and h["level"] == 2 for h in headings)
    assert any(p["text"] == "hello world" for p in paragraphs)
    assert structure["table_count"] == 1
    assert structure["tables"][0][0] == ["Col A", "Col B"]
    assert structure["tables"][0][1] == ["1", "2"]


def test_compose_rejects_invalid_heading_level() -> None:
    with pytest.raises(ValueError, match="heading level"):
        docx_engine.compose(
            blocks=[{"type": "heading", "level": 9, "text": "bad"}]
        )


def test_compose_rejects_unknown_block_type() -> None:
    with pytest.raises(ValueError, match="unknown type"):
        docx_engine.compose(blocks=[{"type": "image", "src": "x.png"}])


def test_compose_rejects_table_with_empty_headers() -> None:
    with pytest.raises(ValueError, match="table headers"):
        docx_engine.compose(
            blocks=[{"type": "table", "headers": [], "rows": []}]
        )


def test_compose_rejects_table_row_with_wrong_column_count() -> None:
    with pytest.raises(ValueError, match="row 0"):
        docx_engine.compose(
            blocks=[
                {
                    "type": "table",
                    "headers": ["A", "B"],
                    "rows": [["only one"]],
                }
            ]
        )


def test_table_block_bolds_header_row() -> None:
    data = docx_engine.compose(
        blocks=[
            {
                "type": "table",
                "headers": ["Name", "Amount"],
                "rows": [["A", "10"]],
            }
        ]
    )
    doc = Document(BytesIO(data))
    table = doc.tables[0]
    for cell in table.rows[0].cells:
        bold_runs = [
            run for para in cell.paragraphs for run in para.runs if run.bold
        ]
        assert bold_runs, f"header cell {cell.text!r} has no bold run"


def test_table_block_sets_dxa_widths_on_table_and_cells() -> None:
    data = docx_engine.compose(
        blocks=[
            {
                "type": "table",
                "headers": ["A", "B", "C"],
                "rows": [["1", "2", "3"]],
            }
        ]
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


def test_append_blocks_extends_existing_document() -> None:
    data = docx_engine.compose(title="Doc")
    data = docx_engine.append_blocks(
        data,
        blocks=[
            {"type": "heading", "level": 2, "text": "Methods"},
            {"type": "paragraph", "text": "para one"},
            {"type": "paragraph", "text": "para two"},
        ],
    )

    doc = Document(BytesIO(data))
    texts = [p.text for p in doc.paragraphs]
    assert "Methods" in texts
    assert "para one" in texts
    assert "para two" in texts


def test_replace_text_substitutes_in_paragraphs_and_tables() -> None:
    data = docx_engine.compose(
        title="Hello {name}",
        blocks=[
            {"type": "heading", "level": 1, "text": "Welcome {name}"},
            {"type": "paragraph", "text": "Dear {name}, your code is {code}."},
            {
                "type": "table",
                "headers": ["Key", "Value"],
                "rows": [["name", "{name}"]],
            },
        ],
    )

    updated, count = docx_engine.replace_text(
        data,
        replacements=[
            {"find": "{name}", "replace": "Yacine"},
            {"find": "{code}", "replace": "42"},
        ],
    )

    # 4 occurrences of {name} (title, heading, paragraph, table cell) + 1 of {code}.
    assert count == 5
    structure = docx_engine.read(updated)
    flat_text = " | ".join(p["text"] for p in structure["paragraphs"])
    assert "Yacine" in flat_text
    assert "{name}" not in flat_text
    assert "42" in flat_text
    assert structure["tables"][0][1] == ["name", "Yacine"]


def test_replace_text_supports_regex() -> None:
    data = docx_engine.compose(
        blocks=[{"type": "paragraph", "text": "Order #1234 and #5678 are pending."}]
    )

    updated, count = docx_engine.replace_text(
        data,
        replacements=[{"find": r"#\d+", "replace": "<id>", "regex": True}],
    )

    assert count == 2
    structure = docx_engine.read(updated)
    text = structure["paragraphs"][0]["text"]
    assert "<id>" in text
    assert "#1234" not in text


def test_replace_text_no_op_returns_zero() -> None:
    data = docx_engine.compose(blocks=[{"type": "paragraph", "text": "stable"}])

    _, count = docx_engine.replace_text(
        data, replacements=[{"find": "absent", "replace": "x"}]
    )
    assert count == 0


def test_replace_text_rejects_empty_find() -> None:
    data = docx_engine.compose()
    with pytest.raises(ValueError, match="'find'"):
        docx_engine.replace_text(
            data, replacements=[{"find": "", "replace": "x"}]
        )


def test_remove_sections_drops_heading_and_following_blocks() -> None:
    data = docx_engine.compose(
        title="Report",
        blocks=[
            {"type": "heading", "level": 1, "text": "Keep me"},
            {"type": "paragraph", "text": "kept paragraph"},
            {"type": "heading", "level": 1, "text": "Drop me"},
            {"type": "paragraph", "text": "dropped paragraph"},
            {"type": "heading", "level": 2, "text": "Drop subsection"},
            {"type": "paragraph", "text": "also dropped"},
            {"type": "heading", "level": 1, "text": "Keep too"},
            {"type": "paragraph", "text": "final paragraph"},
        ],
    )

    updated, removed = docx_engine.remove_sections(data, headings=["Drop me"])

    assert removed == 1
    structure = docx_engine.read(updated)
    texts = [p["text"] for p in structure["paragraphs"]]
    assert "Keep me" in texts
    assert "kept paragraph" in texts
    assert "Drop me" not in texts
    assert "dropped paragraph" not in texts
    assert "Drop subsection" not in texts
    assert "also dropped" not in texts
    assert "Keep too" in texts
    assert "final paragraph" in texts


def test_remove_sections_stops_at_higher_level_heading() -> None:
    data = docx_engine.compose(
        blocks=[
            {"type": "heading", "level": 2, "text": "Drop"},
            {"type": "paragraph", "text": "p1"},
            {"type": "heading", "level": 1, "text": "Keep"},
            {"type": "paragraph", "text": "p2"},
        ],
    )

    updated, removed = docx_engine.remove_sections(data, headings=["Drop"])

    assert removed == 1
    structure = docx_engine.read(updated)
    texts = [p["text"] for p in structure["paragraphs"]]
    assert "Drop" not in texts
    assert "p1" not in texts
    assert "Keep" in texts
    assert "p2" in texts


def test_remove_sections_unmatched_heading_returns_zero() -> None:
    data = docx_engine.compose(
        blocks=[
            {"type": "heading", "level": 1, "text": "Only one"},
            {"type": "paragraph", "text": "body"},
        ],
    )

    updated, removed = docx_engine.remove_sections(data, headings=["Missing"])
    assert removed == 0
    structure = docx_engine.read(updated)
    assert any(p["text"] == "Only one" for p in structure["paragraphs"])


def test_remove_sections_handles_table_inside_section() -> None:
    data = docx_engine.compose(
        blocks=[
            {"type": "heading", "level": 1, "text": "Drop"},
            {
                "type": "table",
                "headers": ["A", "B"],
                "rows": [["1", "2"]],
            },
            {"type": "heading", "level": 1, "text": "Keep"},
        ],
    )

    updated, removed = docx_engine.remove_sections(data, headings=["Drop"])

    assert removed == 1
    structure = docx_engine.read(updated)
    assert structure["table_count"] == 0
    assert any(p["text"] == "Keep" for p in structure["paragraphs"])
