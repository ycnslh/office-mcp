# Word Document Formatting Rules

These rules define how the office-mcp engine produces professional `.docx`
files. Consult this resource **before** calling a `docx_*` tool — the engine
applies most of these automatically, but several rules constrain what you
should ask the tool to do.

## Page setup

- Page size: A4 (or US Letter for North-American audiences — override via template).
- Margins: 1" on every side (2.54 cm).
- Default body font: Arial 11 pt.
- Heading sizes: H1 16 pt, H2 14 pt, H3 12 pt, all Arial.

## Structure

- Use `Heading 1` for top-level sections (one per major chapter).
- Use `Heading 2` for subsections, `Heading 3` sparingly for sub-subsections.
- Never fake a heading with bold + larger font — use the real style so the
  reader can build a table of contents and screen readers can navigate.
- Prefer one document title (rendered as Word's `Title` style) over a bare H1.

## Paragraphs

- Single line spacing, 8 pt space-after between paragraphs.
- Justify long-form text; left-align tables, short captions, and UI labels.
- No consecutive blank paragraphs — use the paragraph-level "space after"
  property for vertical rhythm.

## Lists

- Use Word's built-in bullet or number list styles.
- Never substitute Unicode glyphs (`•`, `●`, `–`, `*`, `-`) for real list
  formatting: they break list reordering, screen readers, and exports.
- Cap list nesting at 3 levels; beyond that, prefer a table.

## Tables

- **Always** set explicit column widths in DXA (1/20 of a point). The
  `docx_insert_table` tool does this automatically: total 9360 DXA =
  6.5" usable width (fits A4 with 1" margins).
- Distribute columns evenly unless the data genuinely asks for asymmetry.
- Bold the header row. Use a subtle border style (the engine defaults to
  "Light Grid Accent 1").
- Keep cell text short; for long prose, use a paragraph and reference the
  table by caption.

## Forbidden

- Emojis and decorative Unicode in body text (engineering docs, contracts,
  reports — the default target). If you need emphasis, use a callout box
  or a highlighted run.
- Manual page breaks as a layout tool. Use a proper section break.
- Inline color text for emphasis. Use a named character style (`Strong`,
  `Emphasis`) or a highlight.
- Multiple consecutive blank lines to fake spacing.

## When in doubt

Use a template. Call the `office://templates/list` resource (V2) to see
what is available; when nothing fits, fall back to `docx_create` and build
the document section by section.
