"""Per-request user context — populated from LibreChat headers by middleware."""

from __future__ import annotations

from collections.abc import Iterator
from contextlib import contextmanager
from contextvars import ContextVar
from dataclasses import dataclass


@dataclass(frozen=True, slots=True)
class UserContext:
    """Identity and credentials scoped to a single MCP request.

    ``user_id`` is the mandatory scoping key — every storage operation is
    constrained to the caller's subtree. ``email`` is informational. The
    ``on_behalf_of_token`` field carries an OAuth token for Graph API calls
    when the backend authenticates as the user (Phase 7, V2).
    """

    user_id: str
    email: str | None = None
    on_behalf_of_token: str | None = None


_CURRENT: ContextVar[UserContext | None] = ContextVar(
    "office_mcp_current_user", default=None
)
_MCP_SESSION_ID: ContextVar[str | None] = ContextVar(
    "office_mcp_session_id", default=None
)


def current_user() -> UserContext | None:
    """Return the user context for the in-flight request, or ``None``."""
    return _CURRENT.get()


def require_current_user() -> UserContext:
    """Same as :func:`current_user` but raises if the context is missing."""
    user = _CURRENT.get()
    if user is None:
        raise RuntimeError(
            "No UserContext bound to the current request — "
            "ensure UserContextMiddleware is mounted and X-LibreChat-User-Id is sent."
        )
    return user


def current_mcp_session_id() -> str | None:
    """Return the ``Mcp-Session-Id`` header bound to the in-flight request.

    Set by :class:`~office_mcp.server.UserContextMiddleware` from the
    transport-layer session header. ``None`` outside of a request scope
    or before any /mcp POST has minted a session.
    """
    return _MCP_SESSION_ID.get()


@contextmanager
def bind_user(ctx: UserContext) -> Iterator[None]:
    """Bind ``ctx`` as the current request's user, restoring on exit."""
    token = _CURRENT.set(ctx)
    try:
        yield
    finally:
        _CURRENT.reset(token)


@contextmanager
def bind_mcp_session_id(session_id: str | None) -> Iterator[None]:
    """Bind ``session_id`` for the in-flight request, restoring on exit."""
    token = _MCP_SESSION_ID.set(session_id)
    try:
        yield
    finally:
        _MCP_SESSION_ID.reset(token)
