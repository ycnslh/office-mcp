# Excel Spreadsheet Formatting Rules

These rules govern `.xlsx` files produced by office-mcp. The engine
applies the formatting bits (color coding, fonts) automatically; the
structural rules constrain how you should organise data through tool
calls.

## Color coding (finance convention)

Universal financial-modeling convention. The engine applies it
automatically when you use the `xlsx_write_range` and
`xlsx_apply_formula` tools — do **not** override unless the workbook
has its own house style.

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
  until LibreOffice recomputes them. The `xlsx_apply_formula` tool
  triggers a recalc by default; pass `recalc=False` only when you'll
  immediately apply more formulas, then recalc once at the end.

## Forbidden

- Hardcoded numbers in cells that should be formulas.
- Merged cells inside data ranges (they break sort / filter / pivot).
- Color used for emphasis without a documented semantic meaning.
- Multiple data tables on a single sheet.
- Volatile functions in performance-sensitive workbooks.

## Templates

For recurring report types (monthly P&L, customer list, budget tracker)
prefer a template. List available templates with
`office://templates/list` (V2).
