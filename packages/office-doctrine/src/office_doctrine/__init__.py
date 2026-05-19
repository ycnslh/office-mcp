"""Pure bytes-in / bytes-out engines for .docx, .xlsx and .pptx.

The three submodules — :mod:`docx`, :mod:`xlsx`, :mod:`pptx` — encapsulate
office-mcp's formatting doctrine without depending on any storage backend
or MCP plumbing. They are safe to import from a Code Interpreter sandbox.

Submodules are exposed here; their public symbols are NOT re-exported at
package level to keep import cycles trivial and namespaces clean::

    from office_doctrine import docx
    out = docx.compose(title="T", blocks=[])
"""

from __future__ import annotations

from . import docx, pptx, xlsx

__all__ = ["docx", "xlsx", "pptx"]
