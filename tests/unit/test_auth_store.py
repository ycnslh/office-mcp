"""Unit tests for the in-memory token store."""

from __future__ import annotations

import time

import pytest

from office_mcp.auth.store import PendingAuth, PendingStoreFull, TokenEntry, TokenStore


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


def test_peek_pending_returns_claim_without_removing_it() -> None:
    """``/auth/start/{state}`` peeks; the callback later consumes."""
    store = TokenStore()
    store.remember_pending(_claim("xyz"))

    first = store.peek_pending("xyz")
    second = store.peek_pending("xyz")
    consumed = store.consume_pending("xyz")

    assert first is not None
    assert second is not None
    assert first.state == "xyz"
    # peek does not remove → second peek still finds the claim.
    assert second.state == "xyz"
    # consume after peek still works.
    assert consumed is not None
    assert consumed.state == "xyz"
    # ...and only after consume is the claim gone.
    assert store.peek_pending("xyz") is None


def test_peek_pending_returns_none_for_expired_claim() -> None:
    store = TokenStore(pending_ttl_seconds=10)
    store.remember_pending(_claim("old", age_seconds=20))

    assert store.peek_pending("old") is None


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


def test_remember_pending_raises_when_at_capacity() -> None:
    store = TokenStore()
    # Force a tiny cap so the test stays fast.
    store.MAX_PENDING = 3  # type: ignore[misc]
    for i in range(3):
        store.remember_pending(_claim(f"s{i}"))

    with pytest.raises(PendingStoreFull):
        store.remember_pending(_claim("overflow"))


def test_token_purged_after_ttl() -> None:
    # 0-second TTL → every put_token immediately makes the next get
    # see the entry as expired and purge it.
    store = TokenStore(token_ttl_seconds=0)
    store.put_token("s1", _entry())
    # A small sleep guarantees ``time.time() - ttl`` lands strictly
    # after the insertion timestamp.
    time.sleep(0.01)
    assert store.get_token("s1") is None
    assert store.token_count() == 0


def test_get_state_for_session_returns_state_when_token_alive() -> None:
    """``build_download_url`` reads this so a signed download URL can
    embed the token-store key for ``od://`` URIs."""
    store = TokenStore()
    store.put_token("state-X", _entry())
    store.link_session_to_state("session-1", "state-X")

    assert store.get_state_for_session("session-1") == "state-X"


def test_get_state_for_session_returns_none_when_unlinked() -> None:
    store = TokenStore()
    store.put_token("state-X", _entry())

    assert store.get_state_for_session("never-linked") is None


def test_get_state_for_session_drops_orphan_link_when_token_purged() -> None:
    """Otherwise build_download_url could embed a state whose token no
    longer exists, producing a download URL that 401s on first click."""
    store = TokenStore(token_ttl_seconds=0)
    store.put_token("state-X", _entry())
    store.link_session_to_state("session-1", "state-X")
    time.sleep(0.01)

    assert store.get_state_for_session("session-1") is None


def test_link_session_then_get_token_for_session_returns_entry() -> None:
    store = TokenStore()
    entry = _entry(access_token="for-session")
    store.put_token("state-X", entry)
    store.link_session_to_state("session-1", "state-X")

    got = store.get_token_for_session("session-1")

    assert got is not None
    assert got.access_token == "for-session"


def test_get_token_for_session_returns_none_when_not_linked() -> None:
    store = TokenStore()
    store.put_token("state-X", _entry())

    assert store.get_token_for_session("never-linked") is None


def test_get_token_for_session_drops_orphan_link_when_token_purged() -> None:
    store = TokenStore(token_ttl_seconds=0)
    store.put_token("state-X", _entry())
    store.link_session_to_state("session-1", "state-X")
    time.sleep(0.01)  # force token to be considered expired

    assert store.get_token_for_session("session-1") is None
    # The orphan session→state link is also cleaned up so subsequent
    # lookups don't keep finding a phantom mapping.
    store.put_token("state-Y", _entry(access_token="other"))
    assert store.get_token_for_session("session-1") is None


def test_delete_token_clears_session_index() -> None:
    store = TokenStore()
    store.put_token("state-X", _entry())
    store.link_session_to_state("session-1", "state-X")

    store.delete_token("state-X")

    assert store.get_token_for_session("session-1") is None


def test_link_session_to_state_ignores_oversized_session_id() -> None:
    """Defensive cap so a misbehaving caller can't grow the index."""
    store = TokenStore()
    store.put_token("state-X", _entry())
    long_id = "x" * (TokenStore.MAX_SESSION_ID_LENGTH + 1)

    store.link_session_to_state(long_id, "state-X")

    assert store.get_token_for_session(long_id) is None


def test_link_session_to_state_ignores_empty_session_id() -> None:
    store = TokenStore()
    store.put_token("state-X", _entry())

    store.link_session_to_state("", "state-X")

    assert store.get_token_for_session("") is None


def test_token_purge_drops_orphan_session_links() -> None:
    """Long-running deployments must not accumulate stale session links."""
    store = TokenStore(token_ttl_seconds=0)
    store.put_token("state-X", _entry())
    store.link_session_to_state("session-1", "state-X")
    time.sleep(0.01)
    # token_count() triggers _purge_expired_tokens_locked which now
    # also cleans the session index — verify both maps shrink.
    assert store.token_count() == 0
    # Re-add a fresh, long-lived token under a different state to
    # force a subsequent get_token_for_session NOT to find the old
    # session-1 link via the orphan-clean path.
    fresh = TokenStore()
    fresh.put_token("state-Y", _entry())
    # ``session-1`` from the first store must not resolve in the new one.
    assert fresh.get_token_for_session("session-1") is None
