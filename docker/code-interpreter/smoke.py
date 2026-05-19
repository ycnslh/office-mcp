#!/usr/bin/env python3
"""Smoke test for the office-code-interpreter image.

Verifies that the three vendor libs (python-docx / openpyxl / python-pptx)
and our `office_doctrine` wheel can be imported and that the docx engine
produces a sane payload — no LibreOffice recalc, no MCP plumbing, just
pure bytes-out from the lib that PR1 extracted.

Usage::

    docker run --rm office-code-interpreter:test python /smoke.py
    # → "smoke ok (docx=NNNN bytes, libs=docx,openpyxl,pptx,office_doctrine)"
"""

from __future__ import annotations

import sys


def _check_imports() -> list[str]:
    """Return the list of importable libs the sandbox needs for skills."""
    found: list[str] = []
    import docx  # noqa: F401 — python-docx

    found.append("docx")
    import openpyxl  # noqa: F401

    found.append("openpyxl")
    import pptx  # noqa: F401 — python-pptx

    found.append("pptx")
    import office_doctrine  # noqa: F401

    found.append("office_doctrine")
    return found


def main() -> int:
    libs = _check_imports()

    from office_doctrine import docx as od_docx

    payload = od_docx.compose(title="t", blocks=[])
    if not isinstance(payload, (bytes, bytearray)):
        print(f"FAIL  office_doctrine.docx.compose returned {type(payload).__name__}, expected bytes")
        return 1
    if len(payload) <= 1000:
        print(f"FAIL  office_doctrine.docx.compose returned {len(payload)} bytes (expected > 1000)")
        return 1

    print(f"smoke ok (docx={len(payload)} bytes, libs={','.join(libs)})")
    return 0


if __name__ == "__main__":
    sys.exit(main())
