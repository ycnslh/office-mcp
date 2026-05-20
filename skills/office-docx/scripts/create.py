#!/usr/bin/env python3
"""Build a .docx from a JSON payload on stdin; emit bytes on stdout.

Invoke from the LibreChat code interpreter sandbox with ``python3``
(the image does not expose a ``python`` alias).

Payload schema::

    {
        "title": "Optional document title or null",
        "blocks": [
            {"type": "heading",   "level": 1, "text": "..."},  # level 1..3
            {"type": "paragraph", "text": "..."},
            {"type": "table",     "headers": ["..."], "rows": [["..."]]}
        ]
    }

Both fields are optional; an empty payload yields a blank document.
Validation errors from ``office_doctrine.docx.compose`` propagate as
non-zero exit codes with the message on stderr — readable by the LLM.
"""

from __future__ import annotations

import json
import sys

from office_doctrine import docx


def main() -> int:
    raw = sys.stdin.buffer.read()
    payload = json.loads(raw) if raw.strip() else {}
    if not isinstance(payload, dict):
        print("payload must be a JSON object", file=sys.stderr)
        return 2

    title = payload.get("title")
    blocks = payload.get("blocks") or []

    out = docx.compose(title=title, blocks=blocks)
    sys.stdout.buffer.write(out)
    return 0


if __name__ == "__main__":
    sys.exit(main())
