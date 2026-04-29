"""HTTP endpoints driving the Azure AD authorization-code flow.

* ``GET /auth/login``     — server generates a ``state`` and redirects to
                            Microsoft's consent page. The state IS a
                            capability (Phase 7c uses it to find a token
                            for a given chat session), so it is never
                            caller-supplied — only :func:`secrets.token_urlsafe`
                            output crosses the wire.
* ``GET /auth/callback``  — Microsoft sends the user back here with a
                            code; we exchange it on a worker thread and
                            stash the resulting tokens keyed by ``state``.
* ``GET /auth/status``    — liveness probe: returns ``{"authenticated": bool}``.
                            Identity claims (oid/upn/email) are deliberately
                            **not** returned here — possession of a state
                            must not disclose user identity to anyone who
                            happens to know it. The HTML success page
                            shows the UPN to the actual user in their
                            browser, which is the only place it belongs.

Routes are registered as side-effects of importing this module, which
:func:`office_mcp.server._bootstrap` does once. They are exempted from
the bearer middleware in ``server.py`` because the user driving the
flow has not yet authenticated against the office-mcp shared secret —
the OAuth ``state`` and the upstream Microsoft login do the gating
instead.
"""

from __future__ import annotations

import asyncio
import secrets
import time
from html import escape

import structlog
from starlette.requests import Request
from starlette.responses import HTMLResponse, JSONResponse, RedirectResponse, Response

from office_mcp.auth.oauth import OAuthError
from office_mcp.auth.store import PendingAuth, PendingStoreFull
from office_mcp.runtime import get_oauth_client, get_token_store, mcp

log = structlog.get_logger(__name__)


@mcp.custom_route("/auth/login", methods=["GET"])
async def _auth_login(_: Request) -> Response:
    """Kick off the OAuth code flow with a freshly minted state."""
    try:
        client = get_oauth_client()
    except RuntimeError:
        return JSONResponse(
            {"error": "OAuth not configured (set AZURE_TENANT_ID/CLIENT_ID/CLIENT_SECRET)"},
            status_code=503,
        )
    store = get_token_store()

    state = secrets.token_urlsafe(32)
    try:
        store.remember_pending(
            PendingAuth(
                state=state,
                created_at=time.time(),
                redirect_uri=client.redirect_uri,
            )
        )
    except PendingStoreFull as exc:
        log.warning("oauth.login.store_full", error=str(exc))
        return JSONResponse(
            {"error": "auth backend busy, retry shortly"},
            status_code=503,
        )

    auth_url = client.build_auth_url(state=state)
    log.info("oauth.login", state_prefix=state[:8])
    return RedirectResponse(auth_url, status_code=302)


@mcp.custom_route("/auth/callback", methods=["GET"])
async def _auth_callback(request: Request) -> Response:
    """Receive Microsoft's redirect after consent and exchange the code."""
    try:
        client = get_oauth_client()
    except RuntimeError:
        return JSONResponse({"error": "OAuth not configured"}, status_code=503)
    store = get_token_store()

    error = request.query_params.get("error")
    if error:
        desc = request.query_params.get("error_description", "")
        log.warning("oauth.callback.error", error=error, description=desc)
        message = f"{error}: {desc}" if desc else error
        return _result_page(success=False, message=message)

    code = request.query_params.get("code")
    state = request.query_params.get("state")
    if not code or not state:
        return JSONResponse(
            {"error": "missing 'code' or 'state' query parameter"},
            status_code=400,
        )

    pending = store.consume_pending(state)
    if pending is None:
        log.warning("oauth.callback.unknown_state", state_prefix=state[:8])
        return JSONResponse(
            {"error": "unknown or expired state"},
            status_code=400,
        )

    try:
        # MSAL hits Microsoft over HTTPS synchronously (~hundreds of ms);
        # off-load it so the event loop keeps serving other requests.
        entry = await asyncio.to_thread(client.exchange_code, code=code)
    except OAuthError as exc:
        return _result_page(success=False, message=str(exc))

    store.put_token(state, entry)
    log.info(
        "oauth.callback.ok",
        state_prefix=state[:8],
        user_oid=entry.user_oid,
        user_upn=entry.user_upn,
    )
    return _result_page(
        success=True,
        message=(
            f"Connecté en tant que {entry.user_upn}. "
            "Vous pouvez fermer cet onglet et retourner dans LibreChat."
        ),
    )


@mcp.custom_route("/auth/status", methods=["GET"])
async def _auth_status(request: Request) -> Response:
    """Liveness probe — ``{"authenticated": bool}``, no identity claims.

    We deliberately do not return ``oid`` / ``upn`` / ``email`` here:
    the ``state`` is a capability that may flow through URLs and chat
    transcripts, and disclosing user identity to anyone who guesses or
    intercepts a state would be a confidentiality leak. The actual UPN
    is shown only on the success page rendered to the consenting user's
    browser.
    """
    try:
        get_oauth_client()
    except RuntimeError:
        return JSONResponse({"error": "OAuth not configured"}, status_code=503)
    store = get_token_store()

    state = request.query_params.get("state")
    if not state:
        return JSONResponse({"error": "missing 'state' query parameter"}, status_code=400)

    entry = store.get_token(state)
    if entry is None or entry.is_expired():
        return JSONResponse({"authenticated": False})
    return JSONResponse({"authenticated": True})


def _result_page(*, success: bool, message: str) -> HTMLResponse:
    """Tiny self-contained HTML page rendered after the callback."""
    color = "#16a34a" if success else "#dc2626"
    title = "Connexion réussie" if success else "Échec de la connexion"
    body = f"""<!doctype html>
<html lang="fr">
<head>
<meta charset="utf-8">
<title>{escape(title)}</title>
<style>
  body {{
    font-family: -apple-system, BlinkMacSystemFont, system-ui, sans-serif;
    background: #0b0f17; color: #e5e7eb;
    display: grid; place-items: center;
    min-height: 100vh; margin: 0;
  }}
  .card {{
    max-width: 480px; padding: 32px;
    border-radius: 12px; background: #111827;
    border-left: 4px solid {color};
  }}
  h1 {{ margin: 0 0 16px; color: {color}; font-size: 1.4rem; }}
  p  {{ margin: 0; line-height: 1.6; }}
</style>
</head>
<body>
  <div class="card">
    <h1>{escape(title)}</h1>
    <p>{escape(message)}</p>
  </div>
</body>
</html>"""
    return HTMLResponse(body, status_code=200 if success else 400)
