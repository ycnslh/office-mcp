"""``python-docx`` wrapper applying office-mcp's formatting doctrine.

Every function takes (and, when mutating, returns) the full ``.docx`` byte
payload. Keeping the filesystem concern out of the engine makes it trivial
to swap in a non-local storage backend.

The public API is built around **batched** operations: a single ``compose``
or ``append_blocks`` call applies many headings, paragraphs and tables in
one shot, so the calling LLM never has to chain sequential edits to
produce a complete document.

Block schema (used by ``compose`` and ``append_blocks``)::

    {"type": "heading", "level": 1, "text": "..."}    # level: 1, 2 or 3
    {"type": "paragraph", "text": "..."}
    {"type": "table", "headers": ["..."], "rows": [["..."]]}
"""

from __future__ import annotations

import re
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
_BLOCK_TYPES = ("heading", "paragraph", "table")


# ── public API ────────────────────────────────────────────────────────


def compose(
    *,
    title: str | None = None,
    blocks: list[dict[str, Any]] | None = None,
) -> bytes:
    """Build a complete document in one shot.

    Args:
        title: Optional document title (Word ``Title`` style).
        blocks: Ordered list of typed content blocks. See module docstring
            for the schema.

    Returns the serialized ``.docx`` bytes.
    """
    doc = Document()
    _apply_default_style(doc)
    if title:
        doc.add_heading(title, level=0)
    _render_blocks(doc, blocks or [])
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


def append_blocks(data: bytes, *, blocks: list[dict[str, Any]]) -> bytes:
    """Append a batch of blocks (headings, paragraphs, tables) to the end."""
    doc = Document(BytesIO(data))
    _render_blocks(doc, blocks)
    return _dump(doc)


def replace_text(
    data: bytes, *, replacements: list[dict[str, Any]]
) -> tuple[bytes, int]:
    """Apply a batch of find/replace operations.

    Each replacement is ``{"find": str, "replace": str, "regex": bool}``.
    Returns the updated bytes and the total number of substitutions made.

    Replacements operate at the paragraph level: when a match is found the
    paragraph's runs are collapsed into a single run holding the new text,
    so per-run formatting on a changed paragraph is lost. Headings and
    table cells are searched too.
    """
    if not replacements:
        return data, 0

    compiled: list[tuple[re.Pattern[str], str]] = []
    for i, rep in enumerate(replacements):
        find = rep.get("find")
        if not isinstance(find, str) or not find:
            raise ValueError(f"replacement {i}: 'find' must be a non-empty string")
        replace = rep.get("replace", "")
        if not isinstance(replace, str):
            raise ValueError(f"replacement {i}: 'replace' must be a string")
        is_regex = bool(rep.get("regex", False))
        pattern = re.compile(find) if is_regex else re.compile(re.escape(find))
        compiled.append((pattern, replace))

    doc = Document(BytesIO(data))
    total = 0
    for para in _iter_all_paragraphs(doc):
        total += _apply_replacements_to_paragraph(para, compiled)
    return _dump(doc), total


def remove_sections(
    data: bytes, *, headings: list[str]
) -> tuple[bytes, int]:
    """Remove sections matching the given heading texts.

    For each heading text in ``headings``, the matching heading paragraph
    and every subsequent block (paragraphs and tables) up to but not
    including the next heading of the same or higher level is removed.

    Matching is exact and case-sensitive on the heading text. A heading
    text is considered matched at most once: duplicates in the document
    require duplicated entries in ``headings``.

    Returns the updated bytes and the number of sections removed.
    """
    if not headings:
        return data, 0

    doc = Document(BytesIO(data))
    body = doc.element.body
    children = list(body.iterchildren())

    targets = list(headings)
    to_remove: list[Any] = []
    i = 0
    while i < len(children):
        el = children[i]
        level = _element_heading_level(el)
        if level is not None:
            text = "".join(t.text or "" for t in el.iter(qn("w:t")))
            if text in targets:
                targets.remove(text)
                j = i + 1
                while j < len(children):
                    next_level = _element_heading_level(children[j])
                    if next_level is not None and next_level <= level:
                        break
                    j += 1
                to_remove.extend(children[i:j])
                i = j
                continue
        i += 1

    for el in to_remove:
        body.remove(el)

    removed = len(headings) - len(targets)
    return _dump(doc), removed


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


def _render_blocks(doc: Document, blocks: list[dict[str, Any]]) -> None:
    """Render an ordered list of typed blocks into ``doc``."""
    for i, block in enumerate(blocks):
        btype = block.get("type")
        if btype == "heading":
            level = block.get("level", 1)
            if level not in _ALLOWED_HEADING_LEVELS:
                raise ValueError(
                    f"block {i}: heading level must be one of "
                    f"{list(_ALLOWED_HEADING_LEVELS)}, got {level}"
                )
            text = block.get("text", "")
            doc.add_heading(text, level=level)
        elif btype == "paragraph":
            doc.add_paragraph(block.get("text", ""))
        elif btype == "table":
            headers = block.get("headers") or []
            rows = block.get("rows") or []
            _add_table(doc, headers=headers, rows=rows, block_index=i)
        else:
            raise ValueError(
                f"block {i}: unknown type {btype!r}, "
                f"expected one of {_BLOCK_TYPES}"
            )


def _add_table(
    doc: Document,
    *,
    headers: list[str],
    rows: list[list[str]],
    block_index: int,
) -> None:
    """Append a table with a bold header row and DXA-pinned column widths."""
    if not headers:
        raise ValueError(f"block {block_index}: table headers must not be empty")
    for r, row in enumerate(rows):
        if len(row) != len(headers):
            raise ValueError(
                f"block {block_index}: row {r} has {len(row)} cells "
                f"but table has {len(headers)} columns"
            )

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


def _heading_level(style_name: str) -> int:
    """Return ``1``/``2``/``3`` for ``Heading 1``/``Heading 2``/``Heading 3``."""
    tail = style_name.removeprefix("Heading").strip()
    if tail.isdigit():
        return int(tail)
    return 0  # Title or unknown level


def _element_heading_level(element: Any) -> int | None:
    """Return the heading level of a body element, or ``None`` if not a heading.

    Inspects a ``<w:p>`` for ``<w:pPr><w:pStyle w:val="Heading N"/></w:pPr>``.
    """
    if element.tag != qn("w:p"):
        return None
    pStyle = element.find(qn("w:pPr") + "/" + qn("w:pStyle"))
    if pStyle is None:
        return None
    val = pStyle.get(qn("w:val")) or ""
    if not val.startswith("Heading"):
        return None
    tail = val.removeprefix("Heading").strip()
    return int(tail) if tail.isdigit() else None


def _iter_all_paragraphs(doc: Document) -> Any:
    """Yield every paragraph in the document, including those inside table cells."""
    yield from doc.paragraphs
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                yield from cell.paragraphs


def _apply_replacements_to_paragraph(
    paragraph: Any, compiled: list[tuple[re.Pattern[str], str]]
) -> int:
    """Apply every compiled replacement to ``paragraph``. Returns the count."""
    original = paragraph.text
    new_text = original
    count = 0
    for pattern, replace in compiled:
        new_text, n = pattern.subn(replace, new_text)
        count += n
    if count == 0 or new_text == original:
        return 0

    runs = paragraph.runs
    if not runs:
        paragraph.add_run(new_text)
        return count
    runs[0].text = new_text
    for run in runs[1:]:
        run.text = ""
    return count


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
