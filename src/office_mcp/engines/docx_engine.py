"""Compatibility shim — the docx engine now lives in ``office_doctrine.docx``.

Kept so that existing imports (``from office_mcp.engines import docx_engine``
or ``from office_mcp.engines.docx_engine import compose``) keep working
after the extraction. New code should depend on ``office_doctrine``
directly.
"""

from __future__ import annotations

from office_doctrine.docx import *  # noqa: F401, F403
