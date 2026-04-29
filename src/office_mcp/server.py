"""FastMCP streamable-http server with auth + user context middleware.

Routes:
  GET  /health   — unauthenticated liveness probe (used by Docker, CI, ops).
  ANY  /mcp      — MCP JSON-RPC over streamable HTTP (LibreChat connects here).

Middleware (outermost → innermost):
  BearerAuthMiddleware     — rejects every non-/health request when
                             ``OFFICE_MCP_TOKEN`` is set and the bearer
                             does not match.
  UserContextMiddleware    — extracts ``X-LibreChat-User-Id`` /
                             ``X-LibreChat-User-Email`` into a ContextVar
                             that tools read via :func:`current_user`.

Tools and resources self-register through ``@mcp.tool`` /
``@mcp.resource`` decorators in ``office_mcp.tools`` and
``office_mcp.resources``. ``build_app()`` imports those modules so the
registrations happen before the FastMCP HTTP app is materialised.
"""

from __future__ import annotations

import time
import uuid
from collections import defaultdict
from dataclasses import dataclass, field

import structlog
import uvicorn
from prometheus_client import CONTENT_TYPE_LATEST, generate_latest
from starlette.requests import Request
from starlette.responses import JSONResponse, Response
from starlette.types import ASGIApp, Receive, Scope, Send

from office_mcp import __version__, metrics
from office_mcp.auth import OAuthClient, OAuthConfig, TokenStore
from office_mcp.config import Settings
from office_mcp.context import (
    UserContext,
    bind_mcp_session_id,
    bind_user,
    current_user,
)
from office_mcp.instrumentation import setup_logging
from office_mcp.runtime import (
    get_registry,
    get_token_store,
    mcp,
    set_oauth_client,
    set_registry,
    set_settings,
    set_token_store,
)
from office_mcp.storage import GraphBackend, LocalBackend, StorageRegistry
from office_mcp.utils import signing
from office_mcp.utils.errors import StorageNotFoundError

# ── HTTP routes ────────────────────────────────────────────────────────


@mcp.custom_route("/health", methods=["GET"])
async def _health(_: Request) -> JSONResponse:
    backends: dict[str, str] = {}
    overall = "ok"
    try:
        registry = get_registry()
    except RuntimeError:
        backends = {}
    else:
        for scheme, backend in registry.items():
            try:
                backends[scheme] = await backend.health_check()
            except Exception as exc:  # noqa: BLE001 — surface any failure
                backends[scheme] = f"error: {exc}"
                overall = "degraded"
    return JSONResponse(
        {
            "status": overall,
            "version": __version__,
            "phase": "10-hardening",
            "mcp_ready": True,
            "backends": backends,
        }
    )


@mcp.custom_route("/metrics", methods=["GET"])
async def _prometheus_metrics(_: Request) -> Response:
    return Response(
        content=generate_latest(metrics.REGISTRY),
        media_type=CONTENT_TYPE_LATEST,
    )


@mcp.custom_route("/files/{token}", methods=["GET"])
async def _download(request: Request) -> Response:
    """Serve a file referenced by a signed, user-scoped download token.

    For cloud-backed URIs (``od://``) the signed claim must carry the
    OAuth ``state`` minted at URL-creation time; we resolve the
    matching access token from the in-process :class:`TokenStore` and
    stamp it on the constructed :class:`UserContext` so the Graph
    backend can call ``/me/drive``. URLs minted before this state-aware
    flow (or for ``local://`` URIs that don't need a token) work
    unchanged.
    """
    from office_mcp.runtime import get_settings, get_token_store
    from office_mcp.utils.errors import AuthRequiredError

    token = request.path_params["token"]
    settings = get_settings()
    if not settings.signing_secret:
        metrics.downloads_total.labels(status="disabled").inc()
        return JSONResponse(
            {"error": "downloads disabled (OFFICE_MCP_SIGNING_SECRET unset)"},
            status_code=404,
        )

    try:
        claims = signing.verify(token, secret=settings.signing_secret)
    except signing.InvalidTokenError as exc:
        metrics.downloads_total.labels(status="forbidden").inc()
        return JSONResponse({"error": f"invalid token: {exc}"}, status_code=403)

    on_behalf_of_token: str | None = None
    if claims.state is not None:
        try:
            store = get_token_store()
        except RuntimeError:
            store = None
        if store is not None:
            entry = store.get_token(claims.state)
            if entry is not None and not entry.is_expired():
                on_behalf_of_token = entry.access_token

    user_ctx = UserContext(
        user_id=claims.user_id,
        on_behalf_of_token=on_behalf_of_token,
    )
    try:
        backend = get_registry().get(claims.uri)
        data = await backend.read(claims.uri, user_ctx=user_ctx)
        meta = await backend.stat(claims.uri, user_ctx=user_ctx)
    except StorageNotFoundError:
        metrics.downloads_total.labels(status="not_found").inc()
        return JSONResponse({"error": "not found"}, status_code=404)
    except AuthRequiredError:
        # The signed link's embedded state has expired or never existed.
        # Returning 401 here means "click the original link again from
        # LibreChat after re-authenticating", which is what the user
        # has to do anyway.
        metrics.downloads_total.labels(status="auth_expired").inc()
        return JSONResponse(
            {"error": "OAuth session expired; reopen the file from chat"},
            status_code=401,
        )

    filename = claims.uri.rsplit("/", 1)[-1] or "download"
    metrics.downloads_total.labels(status="ok").inc()
    return Response(
        content=data,
        media_type=meta.mime_type or "application/octet-stream",
        headers={
            "Content-Disposition": f'attachment; filename="{filename}"',
            "Content-Length": str(meta.size_bytes),
        },
    )


# ── ASGI middlewares ───────────────────────────────────────────────────
# Pure ASGI (not BaseHTTPMiddleware) so SSE streams emitted by the MCP
# transport are not buffered into memory.


#: Routes served without the bearer token. Exact-match (not prefix) so
#: a future router upgrade or an oddly-encoded path can't slip past
#: gating into ``/mcp``. ``/files/`` keeps the prefix check below
#: because it carries a path parameter.
_PUBLIC_PATHS = frozenset(
    {
        "/health",
        "/metrics",
        "/auth/login",
        "/auth/callback",
        "/auth/status",
    }
)


class BearerAuthMiddleware:
    """Reject requests missing a valid ``Authorization: Bearer <token>``.

    No-ops when ``expected_token`` is ``None``. Lets a closed set of
    public paths through (probes, signed downloads, OAuth flow) — see
    :data:`_PUBLIC_PATHS` and the ``/files/`` prefix below.
    """

    def __init__(self, app: ASGIApp, expected_token: str | None) -> None:
        self.app = app
        self._expected = expected_token

    async def __call__(self, scope: Scope, receive: Receive, send: Send) -> None:
        if scope["type"] != "http" or self._expected is None:
            await self.app(scope, receive, send)
            return
        path = scope.get("path", "")
        # /health, /metrics: probes / scrapers don't send the bearer.
        # /auth/*: drives the user-facing OAuth flow before any bearer
        #          token is in scope; gating is via the OAuth `state`.
        # /files/*: its own per-request auth lives in the HMAC-signed token.
        if (
            path in _PUBLIC_PATHS
            or path.startswith("/files/")
            or path.startswith("/auth/start/")
        ):
            await self.app(scope, receive, send)
            return
        auth = _header(scope, b"authorization")
        if not auth.startswith("Bearer "):
            await _reject(send, 401, b'{"error":"unauthorized"}')
            return
        if auth.removeprefix("Bearer ").strip() != self._expected:
            await _reject(send, 401, b'{"error":"unauthorized"}')
            return
        await self.app(scope, receive, send)


@dataclass
class _TokenBucket:
    """Refilling token bucket — capacity tokens, ``capacity`` per minute."""

    capacity: int
    tokens: float = field(init=False)
    last_refill: float = field(default_factory=time.monotonic)

    def __post_init__(self) -> None:
        self.tokens = float(self.capacity)

    def take(self) -> bool:
        now = time.monotonic()
        rate_per_second = self.capacity / 60.0
        self.tokens = min(
            float(self.capacity),
            self.tokens + (now - self.last_refill) * rate_per_second,
        )
        self.last_refill = now
        if self.tokens >= 1:
            self.tokens -= 1
            return True
        return False

    def retry_after_seconds(self) -> int:
        if self.tokens >= 1:
            return 0
        rate_per_second = self.capacity / 60.0
        return max(1, int((1 - self.tokens) / rate_per_second) + 1)


class RateLimitMiddleware:
    """Token-bucket rate limiter scoped per user (LibreChat user_id).

    Only `/mcp*` paths are rate-limited; `/health`, `/metrics` and
    `/files/{token}` are excluded so probes, scrapers and downloads
    don't burn the user's budget.
    """

    def __init__(self, app: ASGIApp, requests_per_minute: int) -> None:
        self.app = app
        self._capacity = requests_per_minute
        self._buckets: dict[str, _TokenBucket] = defaultdict(
            lambda: _TokenBucket(capacity=requests_per_minute)
        )

    async def __call__(self, scope: Scope, receive: Receive, send: Send) -> None:
        if scope["type"] != "http" or self._capacity <= 0:
            await self.app(scope, receive, send)
            return
        path = scope.get("path", "")
        if not path.startswith("/mcp"):
            await self.app(scope, receive, send)
            return
        user = current_user()
        user_key = user.user_id if user else "anonymous"
        bucket = self._buckets[user_key]
        if not bucket.take():
            metrics.rate_limited_total.labels(user_id=user_key).inc()
            retry_after = bucket.retry_after_seconds()
            await send(
                {
                    "type": "http.response.start",
                    "status": 429,
                    "headers": [
                        (b"content-type", b"application/json"),
                        (b"retry-after", str(retry_after).encode()),
                    ],
                }
            )
            await send(
                {
                    "type": "http.response.body",
                    "body": (
                        f'{{"error":"rate limit exceeded","retry_after":{retry_after}}}'
                    ).encode(),
                }
            )
            return
        await self.app(scope, receive, send)


class RequestIdMiddleware:
    """Bind a per-request correlation ID into structlog's contextvars.

    Honours an inbound ``X-Request-Id`` header, otherwise generates a
    UUID4. The ID is appended to every log line emitted while the
    request is in flight and echoed back in the response header so
    callers can grep across logs.
    """

    def __init__(self, app: ASGIApp) -> None:
        self.app = app

    async def __call__(self, scope: Scope, receive: Receive, send: Send) -> None:
        if scope["type"] != "http":
            await self.app(scope, receive, send)
            return
        request_id = _header(scope, b"x-request-id") or uuid.uuid4().hex
        token = structlog.contextvars.bind_contextvars(request_id=request_id)

        async def _send_with_header(message: dict[str, object]) -> None:
            if message["type"] == "http.response.start":
                headers = list(message.get("headers", []))  # type: ignore[arg-type]
                headers.append((b"x-request-id", request_id.encode()))
                message = {**message, "headers": headers}
            await send(message)  # type: ignore[arg-type]

        try:
            await self.app(scope, receive, _send_with_header)
        finally:
            structlog.contextvars.unbind_contextvars(*token)


class UserContextMiddleware:
    """Bind a :class:`UserContext` from LibreChat headers for the request.

    Falls back to ``fallback_user_id`` when the incoming header is missing or
    is a literal ``{{...}}`` placeholder — LibreChat 0.8.4 does not substitute
    header placeholders in chat mode (upstream bug #9910), so without this
    fallback the whole MCP surface is unreachable from a plain chat.

    Also reads the ``Mcp-Session-Id`` transport header and, when the
    :class:`~office_mcp.auth.store.TokenStore` has a token bound to that
    session, stamps ``user_ctx.on_behalf_of_token`` so cloud backends
    (Graph, future Drive) can act on the user's behalf without each
    tool re-implementing the lookup. Phase 7c.

    Trust model
    -----------
    ``Mcp-Session-Id`` is treated as a **capability**: any caller able to
    present both a valid bearer (when ``OFFICE_MCP_TOKEN`` is set) and a
    known session id is granted Graph access on behalf of the human
    Microsoft user previously bound to that session via ``/auth/callback``.
    The bearer + session-id combination is therefore the user-impersonation
    boundary on this server — never log a session id at INFO level, never
    accept session ids from sources other than the MCP transport header,
    and rotate the bearer immediately if it ever leaks.
    """

    def __init__(self, app: ASGIApp, fallback_user_id: str = "demo-user") -> None:
        self.app = app
        self._fallback = fallback_user_id

    async def __call__(self, scope: Scope, receive: Receive, send: Send) -> None:
        if scope["type"] != "http":
            await self.app(scope, receive, send)
            return
        raw_user = _header(scope, b"x-librechat-user-id")
        raw_email = _header(scope, b"x-librechat-user-email")
        session_id = _header(scope, b"mcp-session-id") or None
        user_id = raw_user if _looks_like_real_id(raw_user) else self._fallback
        email = raw_email if _looks_like_real_id(raw_email) else None

        access_token = _lookup_session_token(session_id)

        ctx = UserContext(
            user_id=user_id,
            email=email,
            on_behalf_of_token=access_token,
        )
        with bind_user(ctx), bind_mcp_session_id(session_id):
            await self.app(scope, receive, send)


def _lookup_session_token(session_id: str | None) -> str | None:
    """Best-effort lookup of an OAuth access token for an MCP session.

    Returns ``None`` when no session id was supplied, the token store
    is not initialised (e.g. in tests that bypass ``build_app``), the
    session has never been linked to a state, or the token has expired.
    Never raises — failures here must not break unrelated requests.
    """
    if not session_id:
        return None
    try:
        store = get_token_store()
    except RuntimeError:
        return None
    entry = store.get_token_for_session(session_id)
    if entry is None or entry.is_expired():
        return None
    return entry.access_token


def _looks_like_real_id(value: str) -> bool:
    """True if *value* is non-empty and not a literal ``{{...}}`` placeholder."""
    if not value:
        return False
    stripped = value.strip()
    return not (stripped.startswith("{{") and stripped.endswith("}}"))


def _header(scope: Scope, name: bytes) -> str:
    for key, value in scope.get("headers", []):
        if key.lower() == name:
            return value.decode("latin-1")
    return ""


async def _reject(send: Send, status: int, body: bytes) -> None:
    await send(
        {
            "type": "http.response.start",
            "status": status,
            "headers": [(b"content-type", b"application/json")],
        }
    )
    await send({"type": "http.response.body", "body": body})


# ── Bootstrap ──────────────────────────────────────────────────────────


_BOOTSTRAPPED = False


def _bootstrap(settings: Settings) -> None:
    """Wire the default registry and import tool/resource modules once.

    Idempotent: safe to call from both :func:`run` and :func:`build_app`.
    """
    global _BOOTSTRAPPED
    setup_logging(settings.log_level)
    set_settings(settings)
    registry = StorageRegistry()
    registry.register(LocalBackend(root=settings.local_root))
    if settings.oauth_configured:
        # Only register the Graph backend when Azure AD is wired —
        # otherwise `od://` URIs raise UnknownSchemeError, which is
        # the right signal in dev where creds are absent.
        registry.register(GraphBackend())
    set_registry(registry)

    set_token_store(TokenStore())
    set_oauth_client(_build_oauth_client(settings))

    if _BOOTSTRAPPED:
        return
    # Importing these modules triggers the @mcp.tool / @mcp.resource /
    # @mcp.custom_route decorators that register the surface with the
    # FastMCP instance.
    from office_mcp.auth import routes as _auth_routes  # noqa: F401
    from office_mcp.resources import rules  # noqa: F401
    from office_mcp.tools import (  # noqa: F401
        docx_tools,
        fs_tools,
        pptx_tools,
        xlsx_tools,
    )

    _BOOTSTRAPPED = True


def _build_oauth_client(settings: Settings) -> OAuthClient | None:
    """Instantiate the Azure AD client when fully configured, else ``None``.

    Returning ``None`` makes ``/auth/*`` answer 503 instead of crashing
    at startup — useful in dev where Azure creds may be absent.
    """
    tenant_id = settings.azure_tenant_id
    client_id = settings.azure_client_id
    client_secret = settings.azure_client_secret
    base_url = settings.public_base_url
    if not (tenant_id and client_id and client_secret and base_url):
        return None
    return OAuthClient(
        OAuthConfig(
            tenant_id=tenant_id,
            client_id=client_id,
            client_secret=client_secret,
            redirect_uri=f"{base_url.rstrip('/')}/auth/callback",
        )
    )


# ── App factory ────────────────────────────────────────────────────────


def build_app(settings: Settings) -> ASGIApp:
    _bootstrap(settings)
    app = mcp.http_app(path="/mcp")
    # add_middleware prepends — last call wraps everything, so the
    # request flow is: RequestId → BearerAuth → UserContext → RateLimit → MCP.
    app.add_middleware(RateLimitMiddleware, requests_per_minute=settings.rate_limit_per_min)
    app.add_middleware(UserContextMiddleware, fallback_user_id=settings.default_user_id)
    app.add_middleware(BearerAuthMiddleware, expected_token=settings.token)
    app.add_middleware(RequestIdMiddleware)
    return app


def run() -> None:
    settings = Settings()
    uvicorn.run(
        build_app(settings),
        host=settings.host,
        port=settings.port,
        log_level=settings.log_level.lower(),
    )
