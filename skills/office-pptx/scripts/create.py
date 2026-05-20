#!/usr/bin/env python3
"""Build a .pptx from a JSON payload on stdin; emit bytes on stdout.

Invoke from the LibreChat code interpreter sandbox with ``python3``
(the image does not expose a ``python`` alias).

Payload schema::

    {
        "theme": "corporate",
        "slides": [
            {"type": "title",       "title": "...", "subtitle": "..."},
            {"type": "content",     "title": "...", "bullets": ["...", "..."]},
            {"type": "two-content", "title": "...",
                                    "left": ["..."], "right": ["..."]},
            {"type": "section",     "title": "..."}
        ]
    }

``theme`` defaults to ``corporate``. Other supported themes:
``minimal``, ``cyberpunk``. Every slide spec must carry a non-empty
``title``; ``content`` slides need ``bullets``, ``two-content`` slides
need both ``left`` and ``right``.
"""

from __future__ import annotations

import json
import sys

from office_doctrine import pptx


def main() -> int:
    raw = sys.stdin.buffer.read()
    payload = json.loads(raw) if raw.strip() else {}
    if not isinstance(payload, dict):
        print("payload must be a JSON object", file=sys.stderr)
        return 2

    theme = payload.get("theme", "corporate")
    slides = payload.get("slides") or []

    out = pptx.compose(theme=theme, slides=slides)
    sys.stdout.buffer.write(out)
    return 0


if __name__ == "__main__":
    sys.exit(main())
