"""Per-request user context — populated from LibreChat headers in Phase 2."""

from __future__ import annotations

from dataclasses import dataclass


@dataclass(frozen=True, slots=True)
class UserContext:
    """Identity and credentials scoped to a single MCP request.

    ``user_id`` is the mandatory scoping key; every storage operation is
    constrained to the caller's subtree. ``email`` is informational. The
    ``on_behalf_of_token`` field carries an OAuth token for Graph API calls
    when the backend authenticates as the user (Phase 7, V2).
    """

    user_id: str
    email: str | None = None
    on_behalf_of_token: str | None = None
