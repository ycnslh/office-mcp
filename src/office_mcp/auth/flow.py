"""Helpers that drive the OAuth flow on behalf of a tool.

When a tool encounters :class:`~office_mcp.utils.errors.AuthRequiredError`
(typically from :class:`~office_mcp.storage.graph.GraphBackend`), it
needs to hand the caller a Microsoft consent URL that, on completion,
will leave a usable token attached to the **same** MCP transport
session so a re-invocation succeeds.

:func:`start_oauth_flow` does that in one shot: generate a server-side
``state`` (``secrets.token_urlsafe(32)``), register a pending claim
that *carries* the MCP session id, and build the upstream Microsoft
authorize URL. The user clicks it, completes consent, lands on
``/auth/callback`` — which sees the pending claim, exchanges the code,
stashes the token under ``state``, **and** records
``mcp_session_id → state`` so the next tool call in the same session
finds the token via the
:meth:`~office_mcp.auth.store.TokenStore.get_token_for_session` index.

Tools never touch the state value directly — the wrapper around them
in :mod:`office_mcp.instrumentation` calls
:func:`auth_required_response` which packages everything into the
shape an LLM can render to the user.
"""

from __future__ import annotations

import secrets
import time
from typing import Any

import structlog

from office_mcp.auth.store import PendingAuth, PendingStoreFull
from office_mcp.runtime import get_oauth_client, get_token_store

log = structlog.get_logger(__name__)


def start_oauth_flow(mcp_session_id: str | None) -> tuple[str, str]:
    """Mint a state, register the pending claim, return ``(state, login_url)``.

    The returned ``login_url`` is the Microsoft authorize URL — the user
    should open it directly. office-mcp's ``/auth/login`` is **not**
    on the path of this flow; it stays as an ad-hoc browser entry point.

    Raises :class:`RuntimeError` if the OAuth client / token store have
    not been initialised (Azure creds missing) — callers map that to a
    clear "feature disabled" message instead of an opaque 500.
    """
    client = get_oauth_client()
    store = get_token_store()
    state = secrets.token_urlsafe(32)
    try:
        store.remember_pending(
            PendingAuth(
                state=state,
                created_at=time.time(),
                redirect_uri=client.redirect_uri,
                mcp_session_id=mcp_session_id,
            )
        )
    except PendingStoreFull:
        log.warning("oauth.flow.store_full")
        raise
    login_url = client.build_auth_url(state=state)
    log.info(
        "oauth.flow.started",
        state_prefix=state[:8],
        bound_session=bool(mcp_session_id),
    )
    return state, login_url


def auth_required_response(mcp_session_id: str | None) -> dict[str, Any]:
    """Package an "auth required" response shape for tools to return.

    The :mod:`office_mcp.instrumentation` wrapper calls this when a tool
    raises :class:`~office_mcp.utils.errors.AuthRequiredError`. The shape
    is deliberately small so the LLM has no excuse not to surface the
    URL verbatim to the user:

    .. code-block:: json

        {
          "auth_required": true,
          "login_url": "https://login.microsoftonline.com/.../authorize?...",
          "summary": "Authentification Microsoft requise. ..."
        }

    Raises :class:`RuntimeError` (via :func:`start_oauth_flow`) when
    Azure is not configured. The wrapper turns that into a different
    user-facing error so we don't lie about why a feature is unavailable.
    """
    _, login_url = start_oauth_flow(mcp_session_id)
    return {
        "auth_required": True,
        "login_url": login_url,
        # Render as a Markdown link on its own line so the LLM can't
        # paraphrase it into a sentence and silently truncate the URL.
        "summary": (
            "Authentification Microsoft requise pour accéder à OneDrive.\n\n"
            f"[Connectez-vous ici]({login_url})\n\n"
            "Puis demandez-moi à nouveau la même opération."
        ),
    }
