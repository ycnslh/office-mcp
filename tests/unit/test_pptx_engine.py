"""Unit tests for the ``python-pptx`` engine."""

from __future__ import annotations

from io import BytesIO

import pytest
from pptx import Presentation

from office_mcp.engines import pptx_engine


def test_create_returns_valid_pptx_bytes() -> None:
    data = pptx_engine.create()

    assert data.startswith(b"PK")
    prs = Presentation(BytesIO(data))
    # No title slide was requested — deck has zero slides.
    assert len(prs.slides) == 0


def test_create_with_title_adds_title_slide() -> None:
    data = pptx_engine.create(title="Q3 Review", subtitle="Yacine — Netwatch")

    prs = Presentation(BytesIO(data))
    assert len(prs.slides) == 1
    assert prs.slides[0].shapes.title.text == "Q3 Review"


def test_list_themes_returns_known_three() -> None:
    assert set(pptx_engine.list_themes()) == {"corporate", "minimal", "cyberpunk"}


def test_create_rejects_unknown_theme() -> None:
    with pytest.raises(ValueError, match="unknown theme"):
        pptx_engine.create(theme="dracula")


def test_add_slide_content_layout_writes_bullets() -> None:
    data = pptx_engine.create(title="Deck")
    data = pptx_engine.add_slide(
        data,
        layout="content",
        title="Highlights",
        content=["Revenue +18% QoQ", "EU expansion launched", "Hired 3 engineers"],
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


def test_add_slide_two_content_layout_fills_both_columns() -> None:
    data = pptx_engine.create(title="Deck")
    data = pptx_engine.add_slide(
        data,
        layout="two-content",
        title="Pros vs Cons",
        content=["Fast", "Cheap"],
        content_right=["Fragile", "Locks us in"],
    )

    prs = Presentation(BytesIO(data))
    slide = prs.slides[1]
    left = [p.text for p in slide.placeholders[1].text_frame.paragraphs]
    right = [p.text for p in slide.placeholders[2].text_frame.paragraphs]
    assert left == ["Fast", "Cheap"]
    assert right == ["Fragile", "Locks us in"]


def test_add_slide_section_layout_sets_title_only() -> None:
    data = pptx_engine.create(title="Deck")
    data = pptx_engine.add_slide(data, layout="section", title="Part II — Details")

    prs = Presentation(BytesIO(data))
    assert prs.slides[1].shapes.title.text == "Part II — Details"


def test_add_slide_rejects_unknown_layout() -> None:
    data = pptx_engine.create()
    with pytest.raises(ValueError, match="unknown layout"):
        pptx_engine.add_slide(data, layout="grid", title="x")


def test_read_extracts_slides_structure() -> None:
    data = pptx_engine.create(title="Deck", subtitle="sub")
    data = pptx_engine.add_slide(
        data,
        layout="content",
        title="Highlights",
        content=["point a", "point b"],
    )
    data = pptx_engine.add_slide(data, layout="section", title="Closing")

    parsed = pptx_engine.read(data)
    assert parsed["slide_count"] == 3
    assert parsed["slides"][0]["title"] == "Deck"
    assert parsed["slides"][1]["title"] == "Highlights"
    assert "point a" in parsed["slides"][1]["body"]
    assert "point b" in parsed["slides"][1]["body"]
    assert parsed["slides"][2]["title"] == "Closing"


def test_title_runs_get_theme_color() -> None:
    data = pptx_engine.create(title="Demo", theme="cyberpunk")

    prs = Presentation(BytesIO(data))
    runs = prs.slides[0].shapes.title.text_frame.paragraphs[0].runs
    assert runs
    rgb = runs[0].font.color.rgb
    cyan = pptx_engine.THEMES["cyberpunk"].title
    assert rgb == cyan


def test_update_slide_fills_bullets_on_existing_slide() -> None:
    """User-reported case: deck has a title-only slide, the next call
    must populate it with bullets — not append a new slide."""
    data = pptx_engine.create(title="Deck")
    data = pptx_engine.add_slide(data, layout="content", title="Highlights")

    data = pptx_engine.update_slide(
        data, index=1, content=["First", "Second", "Third"]
    )

    parsed = pptx_engine.read(data)
    assert parsed["slide_count"] == 2
    assert parsed["slides"][1]["title"] == "Highlights"
    assert parsed["slides"][1]["body"] == ["First", "Second", "Third"]


def test_update_slide_replaces_existing_bullets() -> None:
    data = pptx_engine.create(title="Deck")
    data = pptx_engine.add_slide(
        data, layout="content", title="Highlights", content=["old A", "old B"]
    )

    data = pptx_engine.update_slide(data, index=1, content=["new only"])

    parsed = pptx_engine.read(data)
    assert parsed["slides"][1]["body"] == ["new only"]


def test_update_slide_can_rewrite_title_only() -> None:
    data = pptx_engine.create(title="Deck")
    data = pptx_engine.add_slide(
        data, layout="content", title="Old", content=["keep me"]
    )

    data = pptx_engine.update_slide(data, index=1, title="New title")

    parsed = pptx_engine.read(data)
    assert parsed["slides"][1]["title"] == "New title"
    assert parsed["slides"][1]["body"] == ["keep me"]


def test_update_slide_two_content_replaces_both_columns() -> None:
    data = pptx_engine.create(title="Deck")
    data = pptx_engine.add_slide(
        data,
        layout="two-content",
        title="Pros / Cons",
        content=["fast"],
        content_right=["fragile"],
    )

    data = pptx_engine.update_slide(
        data,
        index=1,
        content=["fast", "cheap"],
        content_right=["fragile", "lock-in"],
    )

    prs = Presentation(BytesIO(data))
    slide = prs.slides[1]
    left = [p.text for p in slide.placeholders[1].text_frame.paragraphs]
    right = [p.text for p in slide.placeholders[2].text_frame.paragraphs]
    assert left == ["fast", "cheap"]
    assert right == ["fragile", "lock-in"]


def test_update_slide_clears_bullets_when_passed_empty_list() -> None:
    data = pptx_engine.create(title="Deck")
    data = pptx_engine.add_slide(
        data, layout="content", title="Trim me", content=["a", "b"]
    )

    data = pptx_engine.update_slide(data, index=1, content=[])

    parsed = pptx_engine.read(data)
    assert parsed["slides"][1]["body"] == []


def test_update_slide_raises_on_out_of_range_index() -> None:
    data = pptx_engine.create(title="Deck")  # one slide
    with pytest.raises(IndexError, match="out of range"):
        pptx_engine.update_slide(data, index=5, title="oops")


def test_update_slide_silently_skips_missing_placeholder() -> None:
    """Section layout has no body placeholder — passing content is a
    no-op rather than an error so a confused LLM doesn't crash on
    layout mismatch."""
    data = pptx_engine.create(title="Deck")
    data = pptx_engine.add_slide(data, layout="section", title="Divider")

    # Should not raise; content silently dropped because placeholder
    # idx 1 may not exist on a section header.
    pptx_engine.update_slide(data, index=1, content=["ignored"])
