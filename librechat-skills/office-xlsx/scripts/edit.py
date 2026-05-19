#!/usr/bin/env python3
"""Edit an existing .xlsx; emit the updated bytes on stdout.

Invoke from the LibreChat code interpreter sandbox with ``python3``
(the image does not expose a ``python`` alias).

Payload schema (only one action exposed — ``update``)::

    {
        "action": "update",
        "input_path": "/path/to/existing.xlsx",
        "params": {
            "sheets": [
                {
                    "name": "Inputs",
                    "writes": [...],
                    "formulas": [...]
                }
            ],
            "recalc": true
        }
    }

Sheets that don't yet exist are created. ``recalc`` defaults to true and
triggers a single LibreOffice recalc when at least one formula was set.
"""

from __future__ import annotations

import json
import sys
from pathlib import Path

from office_doctrine import xlsx


def main() -> int:
    raw = sys.stdin.buffer.read()
    payload = json.loads(raw) if raw.strip() else {}
    if not isinstance(payload, dict):
        print("payload must be a JSON object", file=sys.stderr)
        return 2

    action = payload.get("action")
    input_path = payload.get("input_path")
    params = payload.get("params") or {}

    if not isinstance(input_path, str) or not input_path:
        print("'input_path' must be a non-empty string", file=sys.stderr)
        return 2

    data = Path(input_path).read_bytes()

    if action == "update":
        out = xlsx.apply_updates(
            data,
            sheets=params.get("sheets") or [],
            recalc=bool(params.get("recalc", True)),
        )
    else:
        print(
            f"unknown action {action!r}; only 'update' is supported",
            file=sys.stderr,
        )
        return 2

    sys.stdout.buffer.write(out)
    return 0


if __name__ == "__main__":
    sys.exit(main())
