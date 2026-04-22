"""Shared pytest fixtures."""

from __future__ import annotations

from collections.abc import Iterator

import pytest

from office_mcp.runtime import set_registry


@pytest.fixture(autouse=True)
def _reset_process_state() -> Iterator[None]:
    """Reset the process-wide storage registry after every test.

    Prevents registry leakage between tests that bootstrap different
    storage roots.
    """
    try:
        yield
    finally:
        set_registry(None)
