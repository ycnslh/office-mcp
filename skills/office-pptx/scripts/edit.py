#!/usr/bin/env python3
"""Edit an existing .pptx; emit the updated bytes on stdout.

Invoke from the LibreChat code interpreter sandbox with ``python3``
(the image does not expose a ``python`` alias).

Payload schema::

    {
        "action": "append" | "replace",
        "input_path": "/path/to/existing.pptx",
        "params": { ...action-specific..., "theme": "corporate" }
    }

Per action:
  - ``append``  → params: {"slides": [...], "theme": "corporate"}
  - ``replace`` → params: {"updates": [{"slide_index": 0,
                                       "title": "...",
                                       "bullets": [...],
                                       "left": [...],
                                       "right": [...]}, ...],
                           "theme": "corporate"}

``theme`` is optional in both cases and defaults to ``corporate``.
"""

from __future__ import annotations

import json
import sys
from pathlib import Path

from office_doctrine import pptx


def main() -> int:
    raw = sys.stdin.buffer.read()
    payload = json.loads(raw) if raw.strip() else {}
    if not isinstance(payload, dict):
        print("payload must be a JSON object", file=sys.stderr)
        return 2

    action = payload.get("action")
    input_path = payload.get("input_path")
    params = payload.get("params") or {}
    theme = params.get("theme", "corporate")

    if not isinstance(input_path, str) or not input_path:
        print("'input_path' must be a non-empty string", file=sys.stderr)
        return 2

    data = Path(input_path).read_bytes()

    if action == "append":
        out = pptx.append_slides(
            data, slides=params.get("slides") or [], theme=theme
        )
    elif action == "replace":
        out = pptx.replace_slides(
            data, updates=params.get("updates") or [], theme=theme
        )
    else:
        print(
            f"unknown action {action!r}; expected one of append|replace",
            file=sys.stderr,
        )
        return 2

    sys.stdout.buffer.write(out)
    return 0


if __name__ == "__main__":
    sys.exit(main())
