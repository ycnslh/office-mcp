"""MCP resources exposing the office-mcp formatting doctrine.

The content is shipped as plain Markdown alongside this module so humans can
read the files directly (in code review, from the browser) without running
the server.
"""

from __future__ import annotations

from pathlib import Path

from office_mcp.runtime import mcp

_CONTENT_DIR = Path(__file__).parent / "content"


def _load(name: str) -> str:
    return (_CONTENT_DIR / name).read_text(encoding="utf-8")


@mcp.resource("office://rules/docx", mime_type="text/markdown")
def docx_rules() -> str:
    """Formatting rules and conventions for Word documents."""
    return _load("docx_rules.md")


@mcp.resource("office://rules/xlsx", mime_type="text/markdown")
def xlsx_rules() -> str:
    """Formatting rules and conventions for Excel spreadsheets."""
    return _load("xlsx_rules.md")


@mcp.resource("office://rules/pptx", mime_type="text/markdown")
def pptx_rules() -> str:
    """Formatting rules and conventions for PowerPoint decks."""
    return _load("pptx_rules.md")
