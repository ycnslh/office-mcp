"""Shared pytest fixtures."""

from __future__ import annotations

from collections.abc import Iterator

import pytest

from office_mcp.runtime import set_oauth_client, set_registry, set_token_store


@pytest.fixture(autouse=True)
def _reset_process_state() -> Iterator[None]:
    """Reset process-wide singletons after every test.

    Prevents leakage between tests that bootstrap different storage
    roots, OAuth credentials, or token stores.
    """
    try:
        yield
    finally:
        set_registry(None)
        set_oauth_client(None)
        set_token_store(None)
