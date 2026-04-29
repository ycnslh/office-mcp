"""Unit tests for HMAC-signed download tokens."""

from __future__ import annotations

import time

import pytest

from office_mcp.utils import signing


def test_sign_then_verify_roundtrip() -> None:
    token = signing.sign(
        user_id="yacine", uri="local://reports/q3.docx", secret="k3y"
    )
    claims = signing.verify(token, secret="k3y")

    assert claims.user_id == "yacine"
    assert claims.uri == "local://reports/q3.docx"
    assert claims.expires_at > int(time.time())
    # Old shape: no state field — claims should reflect that.
    assert claims.state is None


def test_sign_then_verify_carries_state_when_supplied() -> None:
    """The download route uses ``state`` to look up the user's OAuth
    token at request time — pin the round-trip so ``od://`` URLs work."""
    token = signing.sign(
        user_id="yacine",
        uri="od://Decks/pitch.pptx",
        secret="k3y",
        state="opaque-state-abc",
    )
    claims = signing.verify(token, secret="k3y")

    assert claims.state == "opaque-state-abc"
    assert claims.uri == "od://Decks/pitch.pptx"


def test_old_format_token_without_state_still_verifies() -> None:
    """Backward compatibility: tokens minted before the ``state`` field
    landed are still accepted; the route then treats them as
    ``state=None`` (works for ``local://``, fails for ``od://``)."""
    token = signing.sign(user_id="bob", uri="local://x.docx", secret="k3y")
    claims = signing.verify(token, secret="k3y")

    assert claims.state is None


def test_verify_rejects_wrong_secret() -> None:
    token = signing.sign(user_id="alice", uri="local://a", secret="original")
    with pytest.raises(signing.InvalidTokenError, match="signature"):
        signing.verify(token, secret="tampered")


def test_verify_rejects_tampered_payload() -> None:
    token = signing.sign(user_id="alice", uri="local://a", secret="k")
    head, sig = token.split(".", 1)
    # swap payload to point at another user
    fake = signing.sign(user_id="attacker", uri="local://a", secret="k")
    fake_head = fake.split(".", 1)[0]
    forged = fake_head + "." + sig
    with pytest.raises(signing.InvalidTokenError):
        signing.verify(forged, secret="k")


def test_verify_rejects_expired_token() -> None:
    token = signing.sign(
        user_id="alice", uri="local://a", secret="k", ttl_seconds=-1
    )
    with pytest.raises(signing.InvalidTokenError, match="expired"):
        signing.verify(token, secret="k")


def test_verify_rejects_malformed_token() -> None:
    with pytest.raises(signing.InvalidTokenError, match="malformed"):
        signing.verify("not-a-dot-separated-token", secret="k")


def test_sign_rejects_empty_secret() -> None:
    with pytest.raises(ValueError, match="secret"):
        signing.sign(user_id="alice", uri="local://a", secret="")


def test_verify_rejects_when_signing_disabled() -> None:
    token = signing.sign(user_id="alice", uri="local://a", secret="k")
    with pytest.raises(signing.InvalidTokenError, match="disabled"):
        signing.verify(token, secret="")


def test_tokens_are_url_safe() -> None:
    token = signing.sign(
        user_id="yacine@netwatch.fr",
        uri="local://reports/with spaces & weird/file.docx",
        secret="k",
    )
    # URL-safe base64 uses only [A-Za-z0-9_-] plus our separating dot
    assert all(c.isalnum() or c in "._-" for c in token)
