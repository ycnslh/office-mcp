"""Compatibility shim — the pptx engine now lives in ``office_doctrine.pptx``.

Kept so that existing imports (``from office_mcp.engines import pptx_engine``
or ``from office_mcp.engines.pptx_engine import compose``) keep working
after the extraction. New code should depend on ``office_doctrine``
directly.
"""

from __future__ import annotations

from office_doctrine.pptx import *  # noqa: F401, F403
