"""In-memory cache of pending OAuth claims and exchanged tokens.

Two indexed kinds of entry:

* :class:`PendingAuth` — a ``state`` (claim code) that has been issued to
  a caller and redirected toward Azure AD, but has not come back through
  ``/auth/callback`` yet. Expires after ``pending_ttl_seconds`` so a
  user who never finishes the consent doesn't leak slots forever.

* :class:`TokenEntry` — the access + refresh token pair plus minimal
  identity claims, attached to the same ``state`` once the callback
  completes. Tokens are also TTL-bound (default 24 h) so a long-running
  process doesn't accumulate refresh tokens indefinitely.

Both maps are protected by a single :class:`threading.Lock`. The auth
flow is low-volume (one or two flips per chat session) so coarse
locking is fine; we don't expect concurrent callbacks for the same
``state``.

To prevent unbounded growth from an unauthenticated ``/auth/login``
flood, :meth:`TokenStore.remember_pending` enforces a hard cap and
raises :class:`PendingStoreFull` past it — callers map that to a 503.

Persistence is intentionally absent in phase 7a — every restart wipes
the store. Refresh-token persistence and per-``oid`` reuse will land in
a follow-up phase once the basic flow is validated end-to-end.
"""

from __future__ import annotations

import time
from dataclasses import dataclass
from threading import Lock


@dataclass(frozen=True, slots=True)
class TokenEntry:
    """An access + refresh token pair plus the minimal identity claims
    needed to scope Graph operations and surface a useful UPN in logs."""

    access_token: str
    refresh_token: str | None
    expires_at: float
    user_oid: str
    user_upn: str
    user_email: str | None = None

    def is_expired(self, *, now: float | None = None, leeway_seconds: int = 60) -> bool:
        """True if the access token has expired (or will within the leeway).

        The leeway protects against clock skew and against the token
        expiring between the freshness check and the actual Graph call.
        """
        return (now if now is not None else time.time()) + leeway_seconds >= self.expires_at


@dataclass(frozen=True, slots=True)
class PendingAuth:
    """A ``state`` that has been issued to a caller but not yet redeemed."""

    state: str
    created_at: float
    redirect_uri: str
    mcp_session_id: str | None = None


class PendingStoreFull(Exception):
    """The pending-claim cap has been reached — refuse new logins."""


class TokenStore:
    """Thread-safe in-memory store of pending claims and exchanged tokens."""

    #: Hard cap on concurrent un-redeemed pending claims. Stops a flood
    #: on /auth/login from filling memory with state entries.
    MAX_PENDING: int = 1024

    def __init__(
        self,
        *,
        pending_ttl_seconds: int = 600,
        token_ttl_seconds: int = 24 * 3600,
    ) -> None:
        self._pending_ttl = pending_ttl_seconds
        self._token_ttl = token_ttl_seconds
        self._pending: dict[str, PendingAuth] = {}
        self._tokens: dict[str, TokenEntry] = {}
        self._token_inserted_at: dict[str, float] = {}
        self._lock = Lock()

    # ── pending claims ─────────────────────────────────────────────────

    def remember_pending(self, claim: PendingAuth) -> None:
        """Record a ``state`` that has been redirected toward Azure AD.

        Raises :class:`PendingStoreFull` if the store is at capacity even
        after purging expired entries.
        """
        with self._lock:
            self._purge_expired_pending_locked()
            if len(self._pending) >= self.MAX_PENDING:
                raise PendingStoreFull(
                    f"pending claim store full ({self.MAX_PENDING} entries)"
                )
            self._pending[claim.state] = claim

    def consume_pending(self, state: str) -> PendingAuth | None:
        """Pop and return the pending claim, or ``None`` if unknown/expired."""
        with self._lock:
            self._purge_expired_pending_locked()
            return self._pending.pop(state, None)

    def pending_count(self) -> int:
        """Current number of un-redeemed pending claims (after purge)."""
        with self._lock:
            self._purge_expired_pending_locked()
            return len(self._pending)

    # ── exchanged tokens ───────────────────────────────────────────────

    def put_token(self, state: str, entry: TokenEntry) -> None:
        with self._lock:
            self._purge_expired_tokens_locked()
            self._tokens[state] = entry
            self._token_inserted_at[state] = time.time()

    def get_token(self, state: str) -> TokenEntry | None:
        with self._lock:
            self._purge_expired_tokens_locked()
            return self._tokens.get(state)

    def delete_token(self, state: str) -> None:
        with self._lock:
            self._tokens.pop(state, None)
            self._token_inserted_at.pop(state, None)

    def token_count(self) -> int:
        with self._lock:
            self._purge_expired_tokens_locked()
            return len(self._tokens)

    # ── internals ──────────────────────────────────────────────────────

    def _purge_expired_pending_locked(self) -> None:
        cutoff = time.time() - self._pending_ttl
        self._pending = {
            state: claim
            for state, claim in self._pending.items()
            if claim.created_at >= cutoff
        }

    def _purge_expired_tokens_locked(self) -> None:
        cutoff = time.time() - self._token_ttl
        stale = [s for s, t in self._token_inserted_at.items() if t < cutoff]
        for state in stale:
            self._tokens.pop(state, None)
            self._token_inserted_at.pop(state, None)
