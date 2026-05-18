---
name: office-xlsx-rules
description: Use when creating or editing Excel (.xlsx) workbooks via the office-mcp tools (xlsx_create, xlsx_update, xlsx_read_range). Encodes the finance color-coding convention (blue inputs / black formulas / green links), number/date formats, structural rules (one topic per sheet, frozen headers), formula discipline (named ranges, no volatile functions), and the batched-by-intent calling pattern with a single end-of-call recalc.
invocationMode: auto
---

# Excel Spreadsheet Formatting Rules

These rules govern `.xlsx` files produced by office-mcp. The engine
applies the formatting bits (color coding, fonts) automatically; the
structural rules constrain how you should organise data through tool
calls.

## Color coding (finance convention)

Universal financial-modeling convention. The engine applies it
automatically based on the `style` field of each `writes` block in
`xlsx_create` / `xlsx_update` (`"input"` blue, `"formula"` black,
`"link"` green, `"header"` bold) and from the formula content
(`!` in the formula → cross-sheet link → green). Do **not** override
unless the workbook has its own house style.

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
  until LibreOffice recomputes them. `xlsx_create` and `xlsx_update`
  trigger a single recalc at the end by default (one pass for the
  whole workbook). Pass `recalc=False` if you'll chain another update
  and want to amortise the recalc cost.

## Forbidden

- Hardcoded numbers in cells that should be formulas.
- Merged cells inside data ranges (they break sort / filter / pivot).
- Color used for emphasis without a documented semantic meaning.
- Multiple data tables on a single sheet.
- Volatile functions in performance-sensitive workbooks.

## How to call the tools

- **Build the entire workbook in one `xlsx_create` call.** Pass the
  full content via `sheets=[…]`, where each sheet carries its
  `writes` (value blocks) and `formulas` (single-cell formula
  assignments). One LibreOffice recalc fires at the end and bakes
  every formula in one pass.
- **Editing existing workbooks is also batched.** A single
  `xlsx_update(sheets=[…])` adds writes/formulas across one or more
  sheets and runs one recalc.
- **Sheet schema** for both calls:

      {
          "name": "Inputs",
          "writes":   [{"start_cell": "A1",
                        "values": [["Region", "Units", "Price"]],
                        "style": "header"}],
          "formulas": [{"cell": "D2", "formula": "=B2*C2"}],
      }

## Templates

For recurring report types (monthly P&L, customer list, budget tracker)
prefer a template when one is available.
