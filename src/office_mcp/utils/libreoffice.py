"""Headless LibreOffice wrapper.

Used by the xlsx engine to recalculate formulas (openpyxl writes formulas
but never computes their values) and by Phase 6's docx_to_pdf tool.
"""

from __future__ import annotations

import shutil
import subprocess
import tempfile
from pathlib import Path

_BIN_CANDIDATES = ("soffice", "libreoffice")
_DEFAULT_TIMEOUT_S = 60


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

    with tempfile.TemporaryDirectory(prefix="office-mcp-soffice-") as tmp:
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
    """Open the .xlsx in headless LibreOffice and persist computed values."""
    return convert(data, src_ext="xlsx", dst_ext="xlsx", timeout_s=timeout_s)
