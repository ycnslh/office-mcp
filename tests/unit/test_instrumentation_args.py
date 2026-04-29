"""Unit tests for the structured-log arg sanitiser.

The wrapper logs a preview of every tool call's kwargs so the operator
can see what the LLM passed without having to enable verbose MCP
tracing. The preview must be:

* small enough not to blow up a JSON log line on a multi-paragraph
  ``content`` argument;
* informative enough to tell apart "the LLM passed ``content=None``"
  from "the LLM passed ``content=[]``" from "the LLM passed
  ``content=['a', 'b', 'c']``";
* type-aware (a stray bytes payload becomes a length, not a base64
  blob).
"""

from __future__ import annotations

from office_mcp.instrumentation import _ARG_PREVIEW_CHARS, _sanitise_args


def test_passes_none_through() -> None:
    assert _sanitise_args({"content": None}) == {"content": None}


def test_passes_short_strings_through() -> None:
    assert _sanitise_args({"file_uri": "od://Documents/x.pptx"}) == {
        "file_uri": "od://Documents/x.pptx"
    }


def test_truncates_long_strings_with_overflow_hint() -> None:
    long = "x" * (_ARG_PREVIEW_CHARS + 50)

    out = _sanitise_args({"prompt": long})

    preview = out["prompt"]
    assert isinstance(preview, str)
    assert preview.startswith("x" * _ARG_PREVIEW_CHARS)
    assert "+50 chars" in preview


def test_summarises_lists_by_length_and_element_types() -> None:
    out = _sanitise_args({"content": ["a", "b", "c"]})

    assert out["content"] == "<list len=3 types=['str']>"


def test_distinguishes_empty_list_from_none() -> None:
    """Operator must be able to tell ``content=[]`` apart from ``content=None``."""
    assert _sanitise_args({"content": []})["content"] == "<list len=0 types=[]>"
    assert _sanitise_args({"content": None})["content"] is None


def test_summarises_bytes_by_length_not_value() -> None:
    out = _sanitise_args({"payload": b"\x00\x01\x02\x03"})

    assert out["payload"] == "<bytes len=4>"


def test_summarises_dict_by_keys() -> None:
    out = _sanitise_args({"options": {"theme": "corporate", "ratio": 16}})

    assert out["options"] == "<dict keys=['ratio', 'theme']>"


def test_passes_numeric_and_bool_through() -> None:
    out = _sanitise_args({"slide_index": 3, "recalc": True, "ratio": 1.5})

    assert out == {"slide_index": 3, "recalc": True, "ratio": 1.5}


def test_unknown_object_falls_back_to_type_name() -> None:
    class Thing:
        pass

    out = _sanitise_args({"weird": Thing()})

    assert out["weird"] == "<Thing>"
