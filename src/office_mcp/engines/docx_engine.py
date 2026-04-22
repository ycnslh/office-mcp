"""``python-docx`` wrapper applying office-mcp's formatting doctrine.

Every function takes (and, when mutating, returns) the full ``.docx`` byte
payload. Keeping the filesystem concern out of the engine makes it trivial
to swap in a non-local storage backend.
"""

from __future__ import annotations

from io import BytesIO
from typing import Any

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt

# ── doctrine constants ────────────────────────────────────────────────

_BODY_FONT = "Arial"
_BODY_FONT_SIZE_PT = 11
_HEADING_SIZE_PT = {1: 16, 2: 14, 3: 12}
# Total usable width: A4 minus 1" margins = 6.5" = 9360 DXA (1/20 pt).
_TABLE_TOTAL_DXA = 9360
_ALLOWED_HEADING_LEVELS = range(1, 4)


# ── public API ────────────────────────────────────────────────────────


def create(*, title: str | None = None) -> bytes:
    """Build a new document and return its bytes."""
    doc = Document()
    _apply_default_style(doc)
    if title:
        doc.add_heading(title, level=0)
    return _dump(doc)


def read(data: bytes) -> dict[str, Any]:
    """Parse a .docx payload into a structured, LLM-friendly dict."""
    doc = Document(BytesIO(data))
    paragraphs: list[dict[str, Any]] = []
    for para in doc.paragraphs:
        style = (para.style.name or "") if para.style else ""
        text = para.text
        if not text.strip() and not style.startswith("Heading"):
            continue
        if style.startswith("Heading"):
            paragraphs.append(
                {"type": "heading", "level": _heading_level(style), "text": text}
            )
        else:
            paragraphs.append({"type": "paragraph", "text": text})
    tables = [
        [[cell.text for cell in row.cells] for row in table.rows]
        for table in doc.tables
    ]
    return {
        "paragraphs": paragraphs,
        "tables": tables,
        "paragraph_count": len(paragraphs),
        "table_count": len(tables),
    }


def append_section(
    data: bytes,
    *,
    heading: str,
    paragraphs: list[str],
    level: int = 1,
) -> bytes:
    """Append a heading + paragraphs at the end of the document."""
    if level not in _ALLOWED_HEADING_LEVELS:
        raise ValueError(
            f"heading level must be one of {list(_ALLOWED_HEADING_LEVELS)}, got {level}"
        )
    doc = Document(BytesIO(data))
    doc.add_heading(heading, level=level)
    for text in paragraphs:
        doc.add_paragraph(text)
    return _dump(doc)


def insert_table(
    data: bytes,
    *,
    headers: list[str],
    rows: list[list[str]],
) -> bytes:
    """Append a table with explicit DXA column widths and a bold header row."""
    if not headers:
        raise ValueError("headers must not be empty")
    for i, row in enumerate(rows):
        if len(row) != len(headers):
            raise ValueError(
                f"row {i} has {len(row)} cells but table has {len(headers)} columns"
            )

    doc = Document(BytesIO(data))
    table = doc.add_table(rows=1 + len(rows), cols=len(headers))
    table.style = "Light Grid Accent 1"

    header_cells = table.rows[0].cells
    for col, value in enumerate(headers):
        header_cells[col].text = value
        for para in header_cells[col].paragraphs:
            for run in para.runs:
                run.bold = True

    for row_idx, row_values in enumerate(rows, start=1):
        cells = table.rows[row_idx].cells
        for col, value in enumerate(row_values):
            cells[col].text = str(value)

    _apply_double_widths(table)
    return _dump(doc)


# ── internals ────────────────────────────────────────────────────────


def _dump(doc: Document) -> bytes:
    buf = BytesIO()
    doc.save(buf)
    return buf.getvalue()


def _apply_default_style(doc: Document) -> None:
    """Set the document-wide body font + heading sizes."""
    normal = doc.styles["Normal"]
    normal.font.name = _BODY_FONT
    normal.font.size = Pt(_BODY_FONT_SIZE_PT)
    for level, size in _HEADING_SIZE_PT.items():
        try:
            heading = doc.styles[f"Heading {level}"]
        except KeyError:  # pragma: no cover — default template always has these
            continue
        heading.font.name = _BODY_FONT
        heading.font.size = Pt(size)


def _heading_level(style_name: str) -> int:
    """Return ``1``/``2``/``3`` for ``Heading 1``/``Heading 2``/``Heading 3``."""
    tail = style_name.removeprefix("Heading").strip()
    if tail.isdigit():
        return int(tail)
    return 0  # Title or unknown level


def _apply_double_widths(table: Any) -> None:
    """Pin table and per-cell widths in DXA so Word and LibreOffice agree.

    Without this, many viewers auto-fit to content and the same .docx can
    look radically different across clients. The fix is to set both the
    table-level ``tblW`` and each cell's ``tcW`` to equal DXA values.
    """
    cols = len(table.columns)
    if cols == 0:
        return
    col_width = _TABLE_TOTAL_DXA // cols

    tbl_pr = table._tbl.tblPr
    tbl_w = tbl_pr.find(qn("w:tblW"))
    if tbl_w is None:
        tbl_w = OxmlElement("w:tblW")
        tbl_pr.append(tbl_w)
    tbl_w.set(qn("w:w"), str(_TABLE_TOTAL_DXA))
    tbl_w.set(qn("w:type"), "dxa")

    for row in table.rows:
        for cell in row.cells:
            tc_pr = cell._tc.get_or_add_tcPr()
            tc_w = tc_pr.find(qn("w:tcW"))
            if tc_w is None:
                tc_w = OxmlElement("w:tcW")
                tc_pr.append(tc_w)
            tc_w.set(qn("w:w"), str(col_width))
            tc_w.set(qn("w:type"), "dxa")
