"""Unit tests for the scheme → backend registry."""

from __future__ import annotations

from pathlib import Path

import pytest

from office_mcp.storage import LocalBackend, StorageRegistry
from office_mcp.utils.errors import InvalidUriError, UnknownSchemeError


def test_register_and_resolve_by_scheme(tmp_path: Path) -> None:
    registry = StorageRegistry()
    local = LocalBackend(root=tmp_path)
    registry.register(local)

    assert registry.get("local://reports/q3.docx") is local


def test_unknown_scheme_raises() -> None:
    registry = StorageRegistry()
    with pytest.raises(UnknownSchemeError):
        registry.get("sp://site/file.docx")


def test_uri_without_scheme_raises() -> None:
    registry = StorageRegistry()
    with pytest.raises(InvalidUriError):
        registry.get("not-a-uri")


def test_register_overrides_same_scheme(tmp_path: Path) -> None:
    registry = StorageRegistry()
    first = LocalBackend(root=tmp_path / "a")
    second = LocalBackend(root=tmp_path / "b")

    registry.register(first)
    registry.register(second)

    assert registry.get("local://x") is second
