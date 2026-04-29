"""Per-request user context — populated from LibreChat headers by middleware.

Lookup strategy
---------------
Both :func:`current_user` and :func:`current_mcp_session_id` prefer FastMCP's
per-call ``get_http_request()`` over the module-level :class:`ContextVar`
fallbacks. Why: FastMCP runs every tool call inside a long-lived session
task whose context snapshot is taken at session-init time, so ContextVars
set by ASGI middleware on a *later* request do **not** propagate into the
tool. We saw this in production as: same ``request_id`` logged across
unrelated tool calls and ``Mcp-Session-Id`` always read as ``None`` from
inside ``instrumentation.py``.

The ContextVar path stays as the test fallback (``bind_user``) and as a
safety net for code paths that run outside a /mcp HTTP request — e.g. the
``/files/{token}`` download route, which builds its own UserContext.
"""

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

#: Default user id stamped when the LibreChat header is missing or carries
#: the literal ``{{LIBRECHAT_USER_ID}}`` placeholder. Mirrors the middleware
#: default so the HTTP-request fallback returns the same value as the
#: ContextVar path.
_DEFAULT_USER_ID = "demo-user"


def _looks_like_real_id(value: str) -> bool:
    """True if *value* is non-empty and not a literal ``{{...}}`` placeholder."""
    if not value:
        return False
    stripped = value.strip()
    return not (stripped.startswith("{{") and stripped.endswith("}}"))


def _build_from_http() -> tuple[UserContext, str | None] | None:
    """Construct ``(UserContext, session_id)`` from FastMCP's current HTTP request.

    Returns ``None`` when no HTTP request is in scope (tests, the
    ``/files/{token}`` download route, etc.) so callers can fall back to
    the ContextVar path.
    """
    try:
        from fastmcp.server.dependencies import get_http_request
    except ImportError:
        return None
    try:
        request = get_http_request()
    except RuntimeError:
        return None

    headers = request.headers
    raw_user = headers.get("x-librechat-user-id") or ""
    raw_email = headers.get("x-librechat-user-email") or ""
    session_id = headers.get("mcp-session-id") or None

    user_id = raw_user if _looks_like_real_id(raw_user) else _DEFAULT_USER_ID
    email = raw_email if _looks_like_real_id(raw_email) else None
    access_token = _lookup_session_token(session_id)

    return (
        UserContext(
            user_id=user_id,
            email=email,
            on_behalf_of_token=access_token,
        ),
        session_id,
    )


def _lookup_session_token(session_id: str | None) -> str | None:
    """Best-effort token lookup; never raises (e.g. store may be uninitialised)."""
    if not session_id:
        return None
    try:
        from office_mcp.runtime import get_token_store
    except ImportError:
        return None
    try:
        store = get_token_store()
    except RuntimeError:
        return None
    entry = store.get_token_for_session(session_id)
    if entry is None or entry.is_expired():
        return None
    return entry.access_token


def current_user() -> UserContext | None:
    """Return the user context for the in-flight request, or ``None``.

    Prefers per-call HTTP headers over the ContextVar fallback — see the
    module docstring for why ContextVars set by ASGI middleware are not
    reliable inside FastMCP tool execution.
    """
    fresh = _build_from_http()
    if fresh is not None:
        return fresh[0]
    return _CURRENT.get()


def require_current_user() -> UserContext:
    """Same as :func:`current_user` but raises if the context is missing."""
    user = current_user()
    if user is None:
        raise RuntimeError(
            "No UserContext bound to the current request — "
            "ensure UserContextMiddleware is mounted and X-LibreChat-User-Id is sent."
        )
    return user


def current_mcp_session_id() -> str | None:
    """Return the ``Mcp-Session-Id`` header for the in-flight request.

    Reads directly from FastMCP's per-call HTTP request. Falls back to
    the ContextVar (used by tests via :func:`bind_mcp_session_id`).
    ``None`` outside of a request scope.
    """
    fresh = _build_from_http()
    if fresh is not None:
        return fresh[1]
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
