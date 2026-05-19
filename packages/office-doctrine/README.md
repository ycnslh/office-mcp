# office-doctrine

Pure bytes-in / bytes-out Python engines that apply office-mcp's formatting
doctrine to ``.docx``, ``.xlsx`` and ``.pptx`` payloads. Extracted from
[office-mcp](../../) so the same logic can be reused by:

- the office-mcp server (tools layer wraps these engines)
- a self-hosted Code Interpreter image consumed by LibreChat Agent Skills
- any other tool that wants to produce Office documents from a JSON spec

## Install

```bash
pip install office-doctrine
```

The package depends on ``python-docx``, ``openpyxl`` and ``python-pptx``.
``libreoffice`` is **optional** — it is only used by the xlsx engine to
recalculate formulas; when absent the engines log a warning and return
the workbook with un-computed formula cells.

## Quickstart

```python
from office_doctrine import docx, xlsx, pptx

# A minimal Word document (Arial 11, 1" margins enforced).
out = docx.compose(title="T", blocks=[
    {"type": "paragraph", "text": "Hello world"},
])

# A workbook with one sheet and a finance-coded formula.
wb = xlsx.compose(sheets=[
    {
        "name": "Inputs",
        "writes":   [{"start_cell": "A1", "values": [[10], [20], [30]]}],
        "formulas": [{"cell": "A4", "formula": "=SUM(A1:A3)"}],
    }
])

# A 1-slide deck with the corporate theme.
deck = pptx.compose(slides=[{"type": "title", "title": "Demo"}])
```

Each ``compose`` returns the raw bytes of the file; persistence is the
caller's concern.

## Module map

| Submodule | Public functions |
|---|---|
| `office_doctrine.docx` | `compose`, `read`, `append_blocks`, `replace_text`, `remove_sections` |
| `office_doctrine.xlsx` | `compose`, `apply_updates`, `read_range`, `list_sheets` |
| `office_doctrine.pptx` | `compose`, `append_slides`, `replace_slides`, `read`, `list_themes` |

See the docstring at the top of each submodule for the JSON-friendly
input schemas (blocks, sheets, slides).
