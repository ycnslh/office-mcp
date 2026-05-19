"""Compatibility shim — moved to ``office_doctrine._libreoffice``.

``docx_tools.docx_to_pdf`` still imports ``convert`` from here, so the
public surface (``is_available``, ``convert``, ``recalculate_xlsx``,
``LibreOfficeError``) is re-exported. New code should depend on
``office_doctrine`` directly when the dependency makes sense, or call
``convert`` here when the call site is office-mcp-specific (e.g. the
docx → pdf tool).
"""

from __future__ import annotations

from office_doctrine._libreoffice import (
    LibreOfficeError,
    convert,
    is_available,
    recalculate_xlsx,
)

__all__ = ["LibreOfficeError", "convert", "is_available", "recalculate_xlsx"]
