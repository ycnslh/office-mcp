#!/usr/bin/env python3
"""Build a .xlsx from a JSON payload on stdin; emit bytes on stdout.

Invoke from the LibreChat code interpreter sandbox with ``python3``
(the image does not expose a ``python`` alias).

Payload schema::

    {
        "sheets": [
            {
                "name": "Inputs",
                "writes": [
                    {"start_cell": "A1",
                     "values": [["Region", "Units", "Price"]],
                     "style": "header"}
                ],
                "formulas": [
                    {"cell": "D2", "formula": "=B2*C2"}
                ]
            }
        ],
        "recalc": true
    }

``recalc`` defaults to true — LibreOffice bakes computed formula
values in one final pass. Set false to skip when no formulas are
present, or to chain another update later.
"""

from __future__ import annotations

import json
import sys

from office_doctrine import xlsx


def main() -> int:
    raw = sys.stdin.buffer.read()
    payload = json.loads(raw) if raw.strip() else {}
    if not isinstance(payload, dict):
        print("payload must be a JSON object", file=sys.stderr)
        return 2

    sheets = payload.get("sheets") or []
    recalc = payload.get("recalc", True)

    out = xlsx.compose(sheets=sheets, recalc=bool(recalc))
    sys.stdout.buffer.write(out)
    return 0


if __name__ == "__main__":
    sys.exit(main())
