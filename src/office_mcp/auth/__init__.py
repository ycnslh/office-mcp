"""OAuth flow that lets an LLM-driven session attach a real Microsoft user.

LibreChat 0.8.4 forwards literal ``{{LIBRECHAT_USER_ID}}`` placeholders
instead of substituting them, so office-mcp cannot rely on the upstream
chat layer to identify the caller. Instead we drive an out-of-band
authorization-code flow against Azure AD: a tool returns a one-time
``state`` (claim code) and a URL; the user opens it, completes the
Microsoft consent, lands on ``/auth/callback`` which exchanges the code
for delegated tokens and stashes them keyed by ``state``. Subsequent
tool calls can then read OneDrive on the user's behalf.
"""

from office_mcp.auth.oauth import OAuthClient, OAuthConfig, OAuthError
from office_mcp.auth.store import PendingAuth, TokenEntry, TokenStore

__all__ = [
    "OAuthClient",
    "OAuthConfig",
    "OAuthError",
    "PendingAuth",
    "TokenEntry",
    "TokenStore",
]
