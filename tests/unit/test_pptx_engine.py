"""Unit tests for the ``python-pptx`` engine."""

from __future__ import annotations

from io import BytesIO

import pytest
from pptx import Presentation

from office_mcp.engines import pptx_engine


def test_compose_returns_valid_pptx_bytes() -> None:
    data = pptx_engine.compose()

    assert data.startswith(b"PK")
    prs = Presentation(BytesIO(data))
    # No slides requested — deck has zero slides.
    assert len(prs.slides) == 0


def test_compose_renders_title_slide() -> None:
    data = pptx_engine.compose(
        slides=[
            {"type": "title", "title": "Q3 Review", "subtitle": "Yacine — Netwatch"}
        ]
    )

    prs = Presentation(BytesIO(data))
    assert len(prs.slides) == 1
    assert prs.slides[0].shapes.title.text == "Q3 Review"


def test_list_themes_returns_known_three() -> None:
    assert set(pptx_engine.list_themes()) == {"corporate", "minimal", "cyberpunk"}


def test_compose_rejects_unknown_theme() -> None:
    with pytest.raises(ValueError, match="unknown theme"):
        pptx_engine.compose(theme="dracula")


def test_compose_renders_content_slide_with_bullets() -> None:
    data = pptx_engine.compose(
        slides=[
            {"type": "title", "title": "Deck"},
            {
                "type": "content",
                "title": "Highlights",
                "bullets": [
                    "Revenue +18% QoQ",
                    "EU expansion launched",
                    "Hired 3 engineers",
                ],
            },
        ]
    )

    prs = Presentation(BytesIO(data))
    assert len(prs.slides) == 2
    slide = prs.slides[1]
    assert slide.shapes.title.text == "Highlights"
    body = slide.placeholders[1].text_frame
    assert [p.text for p in body.paragraphs] == [
        "Revenue +18% QoQ",
        "EU expansion launched",
        "Hired 3 engineers",
    ]


def test_compose_renders_two_content_slide() -> None:
    data = pptx_engine.compose(
        slides=[
            {"type": "title", "title": "Deck"},
            {
                "type": "two-content",
                "title": "Pros vs Cons",
                "left": ["Fast", "Cheap"],
                "right": ["Fragile", "Locks us in"],
            },
        ]
    )

    prs = Presentation(BytesIO(data))
    slide = prs.slides[1]
    left = [p.text for p in slide.placeholders[1].text_frame.paragraphs]
    right = [p.text for p in slide.placeholders[2].text_frame.paragraphs]
    assert left == ["Fast", "Cheap"]
    assert right == ["Fragile", "Locks us in"]


def test_compose_renders_section_slide() -> None:
    data = pptx_engine.compose(
        slides=[
            {"type": "title", "title": "Deck"},
            {"type": "section", "title": "Part II — Details"},
        ]
    )

    prs = Presentation(BytesIO(data))
    assert prs.slides[1].shapes.title.text == "Part II — Details"


def test_compose_rejects_unknown_slide_type() -> None:
    with pytest.raises(ValueError, match="unknown type"):
        pptx_engine.compose(
            slides=[{"type": "grid", "title": "x"}]
        )


def test_compose_rejects_content_slide_without_bullets() -> None:
    with pytest.raises(ValueError, match="'bullets'"):
        pptx_engine.compose(
            slides=[{"type": "content", "title": "x"}]
        )


def test_compose_rejects_two_content_slide_missing_columns() -> None:
    with pytest.raises(ValueError, match="'left'"):
        pptx_engine.compose(
            slides=[
                {"type": "two-content", "title": "x", "right": ["only right"]}
            ]
        )


def test_read_extracts_slides_structure() -> None:
    data = pptx_engine.compose(
        slides=[
            {"type": "title", "title": "Deck", "subtitle": "sub"},
            {
                "type": "content",
                "title": "Highlights",
                "bullets": ["point a", "point b"],
            },
            {"type": "section", "title": "Closing"},
        ]
    )

    parsed = pptx_engine.read(data)
    assert parsed["slide_count"] == 3
    assert parsed["slides"][0]["title"] == "Deck"
    assert parsed["slides"][1]["title"] == "Highlights"
    assert "point a" in parsed["slides"][1]["body"]
    assert "point b" in parsed["slides"][1]["body"]
    assert parsed["slides"][2]["title"] == "Closing"


def test_title_runs_get_theme_color() -> None:
    data = pptx_engine.compose(
        theme="cyberpunk",
        slides=[{"type": "title", "title": "Demo"}],
    )

    prs = Presentation(BytesIO(data))
    runs = prs.slides[0].shapes.title.text_frame.paragraphs[0].runs
    assert runs
    rgb = runs[0].font.color.rgb
    cyan = pptx_engine.THEMES["cyberpunk"].title
    assert rgb == cyan


def test_append_slides_extends_existing_deck() -> None:
    data = pptx_engine.compose(
        slides=[{"type": "title", "title": "Deck"}]
    )
    data = pptx_engine.append_slides(
        data,
        slides=[
            {"type": "section", "title": "Part II"},
            {
                "type": "content",
                "title": "Bullets",
                "bullets": ["A", "B", "C"],
            },
        ],
    )

    parsed = pptx_engine.read(data)
    assert parsed["slide_count"] == 3
    assert parsed["slides"][1]["title"] == "Part II"
    assert parsed["slides"][2]["body"] == ["A", "B", "C"]


def test_replace_slides_can_rewrite_title_only() -> None:
    data = pptx_engine.compose(
        slides=[
            {"type": "title", "title": "Deck"},
            {"type": "content", "title": "Old", "bullets": ["keep me"]},
        ]
    )

    data = pptx_engine.replace_slides(
        data, updates=[{"slide_index": 1, "title": "New title"}]
    )

    parsed = pptx_engine.read(data)
    assert parsed["slides"][1]["title"] == "New title"
    assert parsed["slides"][1]["body"] == ["keep me"]


def test_replace_slides_replaces_existing_bullets() -> None:
    data = pptx_engine.compose(
        slides=[
            {"type": "title", "title": "Deck"},
            {
                "type": "content",
                "title": "Highlights",
                "bullets": ["old A", "old B"],
            },
        ]
    )

    data = pptx_engine.replace_slides(
        data,
        updates=[{"slide_index": 1, "bullets": ["new only"]}],
    )

    parsed = pptx_engine.read(data)
    assert parsed["slides"][1]["body"] == ["new only"]


def test_replace_slides_out_of_range_raises() -> None:
    data = pptx_engine.compose(
        slides=[{"type": "title", "title": "Deck"}]
    )
    with pytest.raises(IndexError, match="out of range"):
        pptx_engine.replace_slides(
            data, updates=[{"slide_index": 99, "title": "no"}]
        )


def test_replace_slides_can_rewrite_two_content_columns() -> None:
    data = pptx_engine.compose(
        slides=[
            {"type": "title", "title": "Deck"},
            {
                "type": "two-content",
                "title": "Compare",
                "left": ["L1"],
                "right": ["R1"],
            },
        ]
    )

    data = pptx_engine.replace_slides(
        data,
        updates=[
            {
                "slide_index": 1,
                "left": ["L2", "L3"],
                "right": ["R2"],
            }
        ],
    )

    parsed = pptx_engine.read(data)
    body = parsed["slides"][1]["body"]
    assert "L2" in body and "L3" in body
    assert "R2" in body
