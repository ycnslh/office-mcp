"""Unit tests for the in-memory token store."""

from __future__ import annotations

import time

from office_mcp.auth.store import PendingAuth, TokenEntry, TokenStore


def _entry(**overrides: object) -> TokenEntry:
    base: dict[str, object] = {
        "access_token": "at",
        "refresh_token": "rt",
        "expires_at": time.time() + 3600,
        "user_oid": "oid-1",
        "user_upn": "alice@example.com",
        "user_email": None,
    }
    base.update(overrides)
    return TokenEntry(**base)  # type: ignore[arg-type]


def _claim(state: str, *, age_seconds: float = 0) -> PendingAuth:
    return PendingAuth(
        state=state,
        created_at=time.time() - age_seconds,
        redirect_uri="http://localhost/auth/callback",
    )


def test_remember_and_consume_pending_round_trip() -> None:
    store = TokenStore()
    store.remember_pending(_claim("abc"))

    pending = store.consume_pending("abc")

    assert pending is not None
    assert pending.state == "abc"
    # Claim is one-shot
    assert store.consume_pending("abc") is None


def test_pending_purged_after_ttl() -> None:
    store = TokenStore(pending_ttl_seconds=10)
    store.remember_pending(_claim("old", age_seconds=20))

    assert store.consume_pending("old") is None
    assert store.pending_count() == 0


def test_unknown_pending_state_returns_none() -> None:
    store = TokenStore()

    assert store.consume_pending("nope") is None


def test_pending_count_excludes_expired() -> None:
    store = TokenStore(pending_ttl_seconds=10)
    store.remember_pending(_claim("old", age_seconds=20))
    store.remember_pending(_claim("new"))

    assert store.pending_count() == 1


def test_token_get_after_put() -> None:
    store = TokenStore()
    store.put_token("s1", _entry(access_token="hello"))

    got = store.get_token("s1")

    assert got is not None
    assert got.access_token == "hello"


def test_token_delete_removes_it() -> None:
    store = TokenStore()
    store.put_token("s1", _entry())

    store.delete_token("s1")

    assert store.get_token("s1") is None


def test_token_get_unknown_state_returns_none() -> None:
    store = TokenStore()

    assert store.get_token("missing") is None


def test_token_count_grows_and_shrinks() -> None:
    store = TokenStore()
    assert store.token_count() == 0

    store.put_token("a", _entry())
    store.put_token("b", _entry())
    assert store.token_count() == 2

    store.delete_token("a")
    assert store.token_count() == 1


def test_is_expired_with_leeway() -> None:
    fresh = _entry(expires_at=time.time() + 3600)
    assert not fresh.is_expired()

    near = _entry(expires_at=time.time() + 30)  # within 60s leeway
    assert near.is_expired(leeway_seconds=60)
    assert not near.is_expired(leeway_seconds=10)

    past = _entry(expires_at=time.time() - 1)
    assert past.is_expired()
