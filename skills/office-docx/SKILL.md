---
name: office-docx
description: Use when creating or editing Word (.docx) documents. The skill ships executable Python scripts (scripts/create.py and scripts/edit.py) that the agent runs inside the LibreChat code interpreter — the generated .docx then appears as a rich file artifact in the chat. Encodes the formatting doctrine the office-doctrine engine enforces (Arial 11pt body, 1-inch margins, real heading styles, DXA-width tables with bold headers) and the batched-by-intent calling pattern.
invocationMode: auto
---

# Word Document Formatting Rules

These rules define how the `office_doctrine` engine produces professional
`.docx` files. Apply them **before** invoking the scripts shipped with
this skill — the engine applies most of them automatically, but several
constrain what you should ask it to do.

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
  table block (`{"type": "table", ...}`) does this automatically: total
  9360 DXA = 6.5" usable width (fits A4 with 1" margins).
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

## Usage

The skill produces real `.docx` files via the LibreChat code
interpreter. The scripts are available at `office-docx/scripts/` (the
path the skill is mounted under inside the sandbox, per
[LibreChat PR #12831](https://github.com/danny-avila/LibreChat/pull/12831)).
Always invoke with `python3` — the code interpreter image does **not**
expose a `python` alias.

### Create a new document — `scripts/create.py`

Payload schema (stdin JSON):

```json
{
  "title": "Optional document title or null",
  "blocks": [
    {"type": "heading",   "level": 1, "text": "..."},
    {"type": "paragraph", "text": "..."},
    {"type": "table",     "headers": ["..."], "rows": [["..."]]}
  ]
}
```

`level` for headings must be 1, 2, or 3. `blocks` is processed in order
— **build the entire document in one call**, do not chain a
`create` followed by sequential `append`s when you already know the
full content.

Invocation from the code interpreter:

```python
import json, subprocess

payload = {
    "title": "Q3 Report",
    "blocks": [
        {"type": "heading", "level": 1, "text": "Summary"},
        {"type": "paragraph", "text": "Revenue grew 14% QoQ."},
        {"type": "table",
         "headers": ["Region", "Revenue"],
         "rows": [["EMEA", "1.2M"], ["AMER", "0.9M"]]},
    ],
}
result = subprocess.run(
    ["python3", "office-docx/scripts/create.py"],
    input=json.dumps(payload).encode(),
    capture_output=True,
    check=True,
)
with open("q3_report.docx", "wb") as f:
    f.write(result.stdout)
```

The file written to disk is what LibreChat picks up and renders as a
rich artifact preview.

### Edit an existing document — `scripts/edit.py`

Same stdin-JSON pattern, three actions:

```json
{
  "action": "append",
  "input_path": "/path/to/existing.docx",
  "params": {
    "blocks":       [],
    "replacements": [{"find": "...", "replace": "...", "regex": false}],
    "headings":     ["Heading to drop"]
  }
}
```

`action` is one of `append`, `replace`, or `remove`. Only the matching
`params` field is read: `blocks` for `append`, `replacements` for
`replace`, `headings` for `remove`. The bytes of the updated doc are
written to stdout. `replace` and `remove` additionally emit a one-line
JSON `{"count": N}` on stderr so you can report how many substitutions
or sections landed.

```python
payload = {
    "action": "replace",
    "input_path": "q3_report.docx",
    "params": {"replacements": [{"find": "14%", "replace": "17%"}]},
}
result = subprocess.run(
    ["python3", "office-docx/scripts/edit.py"],
    input=json.dumps(payload).encode(),
    capture_output=True,
    check=True,
)
with open("q3_report.docx", "wb") as f:
    f.write(result.stdout)
# result.stderr ends with {"count": 1}
```

Always use `python3`, never `python` — there is no symlink in the
code interpreter image.
