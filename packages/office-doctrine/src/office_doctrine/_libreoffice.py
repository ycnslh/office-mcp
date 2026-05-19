"""Headless LibreOffice wrapper used by the xlsx engine for formula recalc.

This is an internal module of ``office_doctrine``: it is not part of the
package's public surface (the leading underscore signals that).

``recalculate_xlsx`` degrades gracefully when no LibreOffice binary is on
PATH — it returns the input bytes unchanged and logs a warning. This keeps
the engine usable in lightweight environments (CI, dev laptops, slim
Docker images) where the formulas just won't have cached values.

``convert`` keeps its strict behaviour (raises ``LibreOfficeError`` if
the binary is missing), because callers asking for a format conversion
always need an output in the requested format.
"""

from __future__ import annotations

import logging
import shutil
import subprocess
import tempfile
from pathlib import Path

_BIN_CANDIDATES = ("soffice", "libreoffice")
_DEFAULT_TIMEOUT_S = 60

_log = logging.getLogger(__name__)


class LibreOfficeError(RuntimeError):
    """LibreOffice subprocess failed or produced no output."""


def is_available() -> bool:
    """Return True if a headless LibreOffice binary is on PATH."""
    return _resolve_binary() is not None


def _resolve_binary() -> str | None:
    for name in _BIN_CANDIDATES:
        path = shutil.which(name)
        if path is not None:
            return path
    return None


def convert(
    data: bytes,
    *,
    src_ext: str,
    dst_ext: str,
    timeout_s: int = _DEFAULT_TIMEOUT_S,
) -> bytes:
    """Convert ``data`` from ``src_ext`` to ``dst_ext`` and return the bytes.

    A round-trip ``xlsx → xlsx`` works as a recalc trick because LibreOffice
    re-computes every formula on open and bakes the cached values into the
    saved file.
    """
    binary = _resolve_binary()
    if binary is None:
        raise LibreOfficeError(
            "LibreOffice is not installed (looked for: "
            f"{', '.join(_BIN_CANDIDATES)})."
        )

    src_ext = src_ext.lstrip(".")
    dst_ext = dst_ext.lstrip(".")

    with tempfile.TemporaryDirectory(prefix="office-doctrine-soffice-") as tmp:
        tmp_path = Path(tmp)
        in_file = tmp_path / f"in.{src_ext}"
        out_dir = tmp_path / "out"
        out_dir.mkdir()
        in_file.write_bytes(data)

        # Binary path comes from shutil.which (trusted lookup) and every other
        # element is a known-safe literal or a tempfile path we own.
        result = subprocess.run(  # noqa: S603
            [
                binary,
                "--headless",
                "--norestore",
                "--nolockcheck",
                "--nofirststartwizard",
                f"-env:UserInstallation=file://{tmp_path}/uno",
                "--convert-to",
                dst_ext,
                "--outdir",
                str(out_dir),
                str(in_file),
            ],
            check=False,
            capture_output=True,
            timeout=timeout_s,
        )
        if result.returncode != 0:
            raise LibreOfficeError(
                f"libreoffice exited {result.returncode}: "
                f"{result.stderr.decode(errors='replace')[:500]}"
            )

        out_file = out_dir / f"in.{dst_ext}"
        if not out_file.exists():
            raise LibreOfficeError(
                f"libreoffice produced no output (stderr={result.stderr.decode(errors='replace')[:500]})"
            )
        return out_file.read_bytes()


def recalculate_xlsx(data: bytes, *, timeout_s: int = _DEFAULT_TIMEOUT_S) -> bytes:
    """Open the .xlsx in headless LibreOffice and persist computed values.

    Degrades gracefully: if no LibreOffice binary is on PATH, log a warning
    and return ``data`` unchanged so the calling engine still produces a
    usable workbook (formulas will simply lack cached values until the
    file is opened in a real spreadsheet client).
    """
    if _resolve_binary() is None:
        _log.warning(
            "LibreOffice not available on PATH (looked for: %s); "
            "skipping recalc — formula cells will lack cached values",
            ", ".join(_BIN_CANDIDATES),
        )
        return data
    return convert(data, src_ext="xlsx", dst_ext="xlsx", timeout_s=timeout_s)
