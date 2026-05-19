#!/usr/bin/env python3
"""Edit an existing .docx; emit the updated bytes on stdout.

Invoke from the LibreChat code interpreter sandbox with ``python3``
(the image does not expose a ``python`` alias).

Payload schema::

    {
        "action": "append" | "replace" | "remove",
        "input_path": "/path/to/existing.docx",
        "params": { ...action-specific... }
    }

Per action:
  - ``append``  → params: {"blocks": [...]}
  - ``replace`` → params: {"replacements": [{"find": "...",
                                            "replace": "...",
                                            "regex": false}, ...]}
  - ``remove``  → params: {"headings": ["heading text", ...]}

``replace`` writes ``{"count": int}`` to stderr (last line) so the
caller can report how many substitutions landed; ``remove`` does the
same. The output bytes always go to stdout regardless of action.
"""

from __future__ import annotations

import json
import sys
from pathlib import Path

from office_doctrine import docx


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

    if action == "append":
        out = docx.append_blocks(data, blocks=params.get("blocks") or [])
    elif action == "replace":
        out, count = docx.replace_text(
            data, replacements=params.get("replacements") or []
        )
        print(json.dumps({"count": count}), file=sys.stderr)
    elif action == "remove":
        out, count = docx.remove_sections(
            data, headings=params.get("headings") or []
        )
        print(json.dumps({"count": count}), file=sys.stderr)
    else:
        print(
            f"unknown action {action!r}; expected one of append|replace|remove",
            file=sys.stderr,
        )
        return 2

    sys.stdout.buffer.write(out)
    return 0


if __name__ == "__main__":
    sys.exit(main())
