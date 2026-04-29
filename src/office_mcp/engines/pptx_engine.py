"""``python-pptx`` wrapper with named layouts and a small theme palette.

python-pptx is good at creating slides but fiddly at editing slide masters.
Phase 5 keeps theming light: palette (title/body/background colors + font)
is applied per slide at creation time. Full template-based themes (ship an
actual ``.pptx`` file and open it with ``Presentation(path)``) can land in
V2 if the client wants richer decks.
"""

from __future__ import annotations

from dataclasses import dataclass
from io import BytesIO
from typing import Any

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.presentation import Presentation as _Presentation
from pptx.util import Pt

# ── themes ────────────────────────────────────────────────────────────


@dataclass(frozen=True)
class Theme:
    name: str
    bg: RGBColor
    title: RGBColor
    body: RGBColor
    accent: RGBColor
    font: str
    title_pt: int = 36
    body_pt: int = 18


THEMES: dict[str, Theme] = {
    "corporate": Theme(
        name="corporate",
        bg=RGBColor(0xFF, 0xFF, 0xFF),
        title=RGBColor(0x1F, 0x4E, 0x79),   # deep blue
        body=RGBColor(0x2B, 0x2B, 0x2B),    # near-black
        accent=RGBColor(0x2E, 0x75, 0xB6),  # mid blue
        font="Calibri",
    ),
    "minimal": Theme(
        name="minimal",
        bg=RGBColor(0xFF, 0xFF, 0xFF),
        title=RGBColor(0x00, 0x00, 0x00),
        body=RGBColor(0x33, 0x33, 0x33),
        accent=RGBColor(0x88, 0x88, 0x88),
        font="Helvetica",
        title_pt=40,
        body_pt=16,
    ),
    "cyberpunk": Theme(
        name="cyberpunk",
        bg=RGBColor(0x0A, 0x0A, 0x1A),       # near-black navy
        title=RGBColor(0x00, 0xFF, 0xFF),    # cyan
        body=RGBColor(0xF5, 0xF5, 0xF5),     # off-white
        accent=RGBColor(0xFF, 0x1F, 0xAE),   # magenta
        font="Consolas",
    ),
}

DEFAULT_THEME = "corporate"


# ── layouts ───────────────────────────────────────────────────────────

# Maps human layout names to python-pptx slide_layout indices in the
# default blank template. Validated at runtime so a surprising template
# with fewer layouts falls back to the closest match.
_LAYOUTS: dict[str, int] = {
    "title": 0,        # Title slide (title + subtitle)
    "content": 1,      # Title and content (bullets)
    "section": 2,      # Section header
    "two-content": 3,  # Title + two bullet columns
    "blank": 6,
}
_ALLOWED_LAYOUTS = tuple(_LAYOUTS.keys())


# ── public API ────────────────────────────────────────────────────────


def create(
    *,
    theme: str = DEFAULT_THEME,
    title: str | None = None,
    subtitle: str | None = None,
) -> bytes:
    """Build a new deck; if ``title`` is supplied, add a title slide."""
    _require_theme(theme)
    prs = Presentation()
    if title:
        slide = _add_layout_slide(prs, "title")
        _set_title(slide, title, theme)
        if subtitle:
            _set_subtitle(slide, subtitle, theme)
        _apply_background(slide, theme)
    return _dump(prs)


def add_slide(
    data: bytes,
    *,
    layout: str = "content",
    title: str,
    content: list[str] | None = None,
    content_right: list[str] | None = None,
    theme: str = DEFAULT_THEME,
) -> bytes:
    """Append a slide using the named ``layout`` and populate its shapes."""
    if layout not in _LAYOUTS:
        raise ValueError(
            f"unknown layout {layout!r}; expected one of {list(_ALLOWED_LAYOUTS)}"
        )
    _require_theme(theme)
    prs = Presentation(BytesIO(data))
    slide = _add_layout_slide(prs, layout)
    _set_title(slide, title, theme)

    if layout == "content" and content:
        _fill_bullets(slide, placeholder_idx=1, lines=content, theme=theme)
    elif layout == "two-content":
        if content:
            _fill_bullets(slide, placeholder_idx=1, lines=content, theme=theme)
        if content_right:
            _fill_bullets(slide, placeholder_idx=2, lines=content_right, theme=theme)
    # section / blank / title: title-only, nothing else to fill.

    _apply_background(slide, theme)
    return _dump(prs)


def update_slide(
    data: bytes,
    *,
    index: int,
    title: str | None = None,
    content: list[str] | None = None,
    content_right: list[str] | None = None,
    theme: str = DEFAULT_THEME,
) -> bytes:
    """Mutate the slide at ``index`` (0-based) in-place.

    Each field is optional and only acted on when supplied:

    * ``title``: replaces the slide title (no-op if the slide has no
      title placeholder).
    * ``content``: replaces the bullets in placeholder idx 1.
    * ``content_right``: replaces the bullets in placeholder idx 2.

    Bullet replacements clear the existing paragraphs first so the
    caller doesn't end up with the old and new lists concatenated.
    Targets a placeholder that doesn't exist on the slide (e.g.
    passing ``content`` to a section/title layout) is silently
    ignored — we don't pretend to know the layout's intent.
    """
    _require_theme(theme)
    prs = Presentation(BytesIO(data))
    slides = list(prs.slides)
    if not 0 <= index < len(slides):
        raise IndexError(
            f"slide index {index} out of range (deck has {len(slides)} slides)"
        )
    slide = slides[index]
    if title is not None:
        _set_title(slide, title, theme)
    if content is not None:
        _replace_bullets(slide, placeholder_idx=1, lines=content, theme=theme)
    if content_right is not None:
        _replace_bullets(slide, placeholder_idx=2, lines=content_right, theme=theme)
    return _dump(prs)


def read(data: bytes) -> dict[str, Any]:
    """Extract slide titles and body text for LLM consumption."""
    prs = Presentation(BytesIO(data))
    slides: list[dict[str, Any]] = []
    for idx, slide in enumerate(prs.slides, start=1):
        title = _safe_title(slide)
        body: list[str] = []
        for shape in slide.shapes:
            if not shape.has_text_frame:
                continue
            if shape.shape_type == MSO_SHAPE_TYPE.PLACEHOLDER and shape == slide.shapes.title:
                continue
            for para in shape.text_frame.paragraphs:
                text = para.text.strip()
                if text:
                    body.append(text)
        slides.append(
            {
                "index": idx,
                "title": title,
                "body": body,
                "shape_count": len(slide.shapes),
            }
        )
    return {
        "slide_count": len(prs.slides),
        "slides": slides,
    }


def list_themes() -> list[str]:
    return sorted(THEMES.keys())


# ── internals ────────────────────────────────────────────────────────


def _dump(prs: _Presentation) -> bytes:
    buf = BytesIO()
    prs.save(buf)
    return buf.getvalue()


def _require_theme(name: str) -> Theme:
    if name not in THEMES:
        raise ValueError(
            f"unknown theme {name!r}; available: {list_themes()}"
        )
    return THEMES[name]


def _add_layout_slide(prs: _Presentation, layout: str) -> Any:
    target_idx = _LAYOUTS[layout]
    if target_idx >= len(prs.slide_layouts):
        # Fall back to the closest available layout index.
        target_idx = min(target_idx, len(prs.slide_layouts) - 1)
    return prs.slides.add_slide(prs.slide_layouts[target_idx])


def _safe_title(slide: Any) -> str:
    if not slide.shapes.title:
        return ""
    return slide.shapes.title.text or ""


def _set_title(slide: Any, text: str, theme_name: str) -> None:
    if not slide.shapes.title:
        return
    theme = THEMES[theme_name]
    slide.shapes.title.text = text
    _style_text_frame(slide.shapes.title.text_frame, color=theme.title,
                      font=theme.font, size_pt=theme.title_pt, bold=True)


def _set_subtitle(slide: Any, text: str, theme_name: str) -> None:
    placeholder = _placeholder(slide, 1)
    if placeholder is None:
        return
    theme = THEMES[theme_name]
    placeholder.text = text
    _style_text_frame(placeholder.text_frame, color=theme.body,
                      font=theme.font, size_pt=theme.body_pt)


def _fill_bullets(
    slide: Any, *, placeholder_idx: int, lines: list[str], theme: str
) -> None:
    placeholder = _placeholder(slide, placeholder_idx)
    if placeholder is None or not lines:
        return
    palette = THEMES[theme]
    tf = placeholder.text_frame
    tf.text = lines[0]
    for line in lines[1:]:
        para = tf.add_paragraph()
        para.text = line
    _style_text_frame(tf, color=palette.body, font=palette.font, size_pt=palette.body_pt)


def _replace_bullets(
    slide: Any, *, placeholder_idx: int, lines: list[str], theme: str
) -> None:
    """Clear and refill the bullet placeholder at ``placeholder_idx``.

    ``_fill_bullets`` assumes a fresh placeholder; on an existing slide
    it would *append* to the current paragraph list. ``_replace_bullets``
    drops the old paragraphs first, then delegates back to ``_fill_bullets``
    so styling stays consistent.
    """
    placeholder = _placeholder(slide, placeholder_idx)
    if placeholder is None:
        return
    tf = placeholder.text_frame
    tf.clear()
    if not lines:
        return
    _fill_bullets(slide, placeholder_idx=placeholder_idx, lines=lines, theme=theme)


def _placeholder(slide: Any, idx: int) -> Any | None:
    # ``slide.placeholders`` is indexable by idx property, not by position.
    try:
        return slide.placeholders[idx]
    except KeyError:
        return None


def _style_text_frame(
    text_frame: Any,
    *,
    color: RGBColor,
    font: str,
    size_pt: int,
    bold: bool = False,
) -> None:
    for para in text_frame.paragraphs:
        for run in para.runs:
            run.font.name = font
            run.font.size = Pt(size_pt)
            run.font.bold = bold
            run.font.color.rgb = color


def _apply_background(slide: Any, theme_name: str) -> None:
    theme = THEMES[theme_name]
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = theme.bg
