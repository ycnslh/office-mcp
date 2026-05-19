"""Compatibility shim — the xlsx engine now lives in ``office_doctrine.xlsx``.

Kept so that existing imports (``from office_mcp.engines import xlsx_engine``
or ``from office_mcp.engines.xlsx_engine import compose``) keep working
after the extraction. New code should depend on ``office_doctrine``
directly.
"""

from __future__ import annotations

from office_doctrine.xlsx import *  # noqa: F401, F403
