---
name: office-xlsx
description: Use when creating or editing Excel (.xlsx) workbooks. The skill ships executable Python scripts (scripts/create.py and scripts/edit.py) that the agent runs inside the LibreChat code interpreter — the generated .xlsx then appears as a rich file artifact in the chat. Encodes the finance color-coding convention (blue inputs / black formulas / green links), number/date formats, one-topic-per-sheet rule, formula discipline (named ranges, no volatile functions), and a single end-of-call LibreOffice recalc.
invocationMode: auto
---

# Excel Spreadsheet Formatting Rules

These rules govern `.xlsx` files produced by `office_doctrine`. The
engine applies the formatting bits (color coding, fonts) automatically;
the structural rules constrain how you should organise data through
the create/edit scripts.

## Color coding (finance convention)

Universal financial-modeling convention. The engine applies it
automatically based on the `style` field of each `writes` block
(`"input"` blue, `"formula"` black, `"link"` green, `"header"` bold)
and from the formula content (`!` in the formula → cross-sheet link →
green). Do **not** override unless the workbook has its own house
style.

| Cell type | Color           | Used for                                              |
|-----------|-----------------|-------------------------------------------------------|
| **Input**   | blue (#0000FF) | manually entered constants, assumptions, hard-coded values |
| **Formula** | black (#000000) | derived values computed in the same sheet            |
| **Link**    | green (#008000) | cross-sheet references (`=Inputs!B3`)                |

Reading the colors lets a reviewer instantly distinguish what they may
edit (blue) from what they must not (black, green).

## Number formats

- Currency: `$#,##0` or `$#,##0.00` for cents (use the workbook's home
  currency consistently).
- Percentages: `0.0%` with one decimal for changes, `0%` for shares.
- Large numbers: `#,##0` with thousands separator. For summary
  reporting, `#,##0,"K"` for thousands and `#,##0,,"M"` for millions.
- Dates: ISO `YYYY-MM-DD` for cross-locale clarity; never `MM/DD/YYYY`.

## Structure

- One **topic per sheet**: `Inputs`, `Calculations`, `Outputs`. Never
  mix assumption cells and computed cells in the same sheet.
- **Header row in bold** with a subtle fill, frozen if the sheet
  scrolls (`ws.freeze_panes = "A2"`).
- **Right-align numeric columns**, left-align labels.
- **One data table per sheet**. Multiple stacked tables break sorting,
  filtering and pivots.

## Formulas

- Always use a formula for derived values — never hardcode a number
  that could be computed.
- Reference inputs by **named ranges** when the model is non-trivial:
  `=Discount_Rate * Year_1_Revenue` is more legible than `=B3 * D5`.
- Avoid volatile functions (`NOW`, `OFFSET`, `INDIRECT`) — they
  recalculate on every change and confuse reviewers.
- After writing a formula, the spreadsheet's cached values are stale
  until LibreOffice recomputes them. Both `create.py` and `edit.py`
  recalc by default (`recalc: true`); pass `false` to skip when no
  formulas were set or to chain another update later.

## Forbidden

- Hardcoded numbers in cells that should be formulas.
- Merged cells inside data ranges (they break sort / filter / pivot).
- Color used for emphasis without a documented semantic meaning.
- Multiple data tables on a single sheet.
- Volatile functions in performance-sensitive workbooks.

## Usage

The skill produces real `.xlsx` files via the LibreChat code
interpreter. The scripts are available at `office-xlsx/scripts/`
(the path the skill is mounted under inside the sandbox, per
[LibreChat PR #12831](https://github.com/danny-avila/LibreChat/pull/12831)).
Always invoke with `python3` — the code interpreter image does **not**
expose a `python` alias.

### Create a new workbook — `scripts/create.py`

Payload schema (stdin JSON):

```json
{
  "sheets": [
    {
      "name": "Inputs",
      "writes": [
        {"start_cell": "A1",
         "values": [["Region", "Units", "Price"]],
         "style": "header"},
        {"start_cell": "A2",
         "values": [["EMEA", 10, 5.0], ["AMER", 20, 4.5]],
         "style": "input"}
      ],
      "formulas": [
        {"cell": "D2", "formula": "=B2*C2"},
        {"cell": "D3", "formula": "=B3*C3"}
      ]
    }
  ],
  "recalc": true
}
```

`writes[].style` is one of `header`, `input`, `formula`, `link`.
`recalc` defaults to `true` and only fires when at least one formula
was set. **Build the entire workbook in one call** — pass every
sheet in the same payload so the LibreOffice recalc bakes formulas in
one pass.

Invocation from the code interpreter:

```python
import json, subprocess

payload = {
    "sheets": [
        {
            "name": "Sales",
            "writes": [
                {"start_cell": "A1",
                 "values": [["Region", "Units", "Price", "Revenue"]],
                 "style": "header"},
                {"start_cell": "A2",
                 "values": [["EMEA", 10, 5.0], ["AMER", 20, 4.5]],
                 "style": "input"},
            ],
            "formulas": [
                {"cell": "D2", "formula": "=B2*C2"},
                {"cell": "D3", "formula": "=B3*C3"},
            ],
        }
    ],
    "recalc": True,
}
result = subprocess.run(
    ["python3", "office-xlsx/scripts/create.py"],
    input=json.dumps(payload).encode(),
    capture_output=True,
    check=True,
)
with open("sales.xlsx", "wb") as f:
    f.write(result.stdout)
```

### Edit an existing workbook — `scripts/edit.py`

Same stdin-JSON pattern, only `update` is exposed:

```json
{
  "action": "update",
  "input_path": "/path/to/existing.xlsx",
  "params": {
    "sheets": [
      {"name": "Sales", "writes": [...], "formulas": [...]}
    ],
    "recalc": true
  }
}
```

Sheets that don't yet exist are created. The bytes of the updated
workbook are written to stdout.

```python
payload = {
    "action": "update",
    "input_path": "sales.xlsx",
    "params": {
        "sheets": [{
            "name": "Sales",
            "writes": [{"start_cell": "A4",
                        "values": [["APAC", 15, 6.0]],
                        "style": "input"}],
            "formulas": [{"cell": "D4", "formula": "=B4*C4"}],
        }],
        "recalc": True,
    },
}
result = subprocess.run(
    ["python3", "office-xlsx/scripts/edit.py"],
    input=json.dumps(payload).encode(),
    capture_output=True,
    check=True,
)
with open("sales.xlsx", "wb") as f:
    f.write(result.stdout)
```

Always use `python3`, never `python` — there is no symlink in the
code interpreter image.
