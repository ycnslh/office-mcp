"""Shared pytest fixtures. Populated as features land."""

from __future__ import annotations

import pytest


@pytest.fixture
def anyio_backend() -> str:
    return "asyncio"
