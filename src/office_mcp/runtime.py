"""Process-wide runtime state: the FastMCP instance, the storage registry,
and the active :class:`Settings`.

Kept in its own module so tool/resource modules can import the ``mcp``
instance and shared state without creating a cycle through
``office_mcp.server``.
"""

from __future__ import annotations

from fastmcp import FastMCP

from office_mcp.auth import OAuthClient, TokenStore
from office_mcp.config import Settings
from office_mcp.storage import StorageRegistry
from office_mcp.storage.base import FileMetadata
from office_mcp.utils import signing

_INSTRUCTIONS = """\
office-mcp creates and edits Word, Excel and PowerPoint documents.

ADDRESSING FILES
  Every tool takes a `file_uri` like `<scheme>://path/to/file.ext`.
  Two schemes are available in this deployment:
    - `local://`  — the user's per-account scratch space on the server.
                    Always available, no auth required. Use it for
                    drafts, scratch files, and anything the user did not
                    explicitly ask to put in the cloud.
    - `od://`     — the user's Microsoft OneDrive (delegated OAuth).
                    Use when the user mentions OneDrive, Office 365,
                    "my files in the cloud", or asks to save somewhere
                    they can open from their Microsoft account.
  Path components are scoped per user, so a leading `/` is unnecessary.
  Examples: `local://reports/q3.docx`, `od://AI Generated/q3.docx`.

  For OneDrive: when the user does not specify a folder, default to
  `od://AI Generated/<name>` so generated files cluster in one place.
  Bare names (`od://q3.docx`, no folder) are rewritten server-side to
  the same `AI Generated/` subfolder; the response carries the canonical
  `file_uri` — use it verbatim for any follow-up call.

DO NOT EXPLORE FIRST
  Never call `fs_list` or `fs_stat` to "scope" OneDrive, "check if a
  file already exists", or "verify access" before a `*_create`. These
  calls add latency and tokens for nothing. Go straight to the create
  tool with your chosen `file_uri`:

    - Auth: the create call surfaces `auth_required` exactly the same
      way `fs_list` would, so you do not need a probe.
    - Folder normalization: the server rewrites bare-name OneDrive
      writes into a default folder automatically.
    - Existence: create tools overwrite by design — no need to pre-check.

  Use `fs_list` ONLY when the user explicitly asks to see what is in
  a folder, or when an edit tool needs to find an existing file.

AUTHENTICATION FOR od://
  The first tool call that touches `od://` in a fresh chat returns
  `{"auth_required": true, "login_url": "...", "summary": "..."}`
  instead of the normal payload. When you see that:
    1. Surface the `login_url` to the user EXACTLY as a Markdown link —
       do not paraphrase, do not strip query parameters.
    2. Stop and wait for the user to confirm they have signed in.
    3. Once they say they are back, RETRY the same tool call with the
       same arguments — the server picks up the new token automatically.
  Never try to bypass `auth_required` by switching to `local://` unless
  the user explicitly asks you to.

WORKFLOW
  1. Before creating a document, consult the matching `office://rules/<format>`
     resource (docx, xlsx, pptx) for formatting conventions.
  2. Build the entire document body in a single create call (pass the
     full content as the create tool's batched argument — e.g. `blocks`
     for `docx_create`). Do not chain a create + multiple appends when
     you already know the content.
  3. Every mutating tool returns a `download_url` — render it as a
     Markdown link in your answer so the user can click directly from
     the conversation. The URL points to the file's *native* viewer
     when applicable: OneDrive web for `od://` (the user opens the
     file in their Microsoft account), the server's signed `/files/`
     endpoint for `local://`. Example:
       `→ [q3.docx](https://office-mcp.reho.dev/files/...)`
  4. Always echo the resulting `file_uri` alongside the download link.

ERROR HANDLING
  If a tool returns `UnknownSchemeError` or `InvalidUriError`, the URI
  is malformed — fix the scheme/path and retry. Don't silently fall
  back to a different storage location.
"""

mcp: FastMCP = FastMCP("office-mcp", instructions=_INSTRUCTIONS)

_registry: StorageRegistry | None = None


def get_registry() -> StorageRegistry:
    """Return the process-wide storage registry or raise if not initialised."""
    if _registry is None:
        raise RuntimeError(
            "Storage registry not initialised — build_app() or set_registry() "
            "must run before tools can access storage."
        )
    return _registry


def set_registry(registry: StorageRegistry | None) -> None:
    """Install (or clear, with ``None``) the process-wide registry."""
    global _registry
    _registry = registry


_settings: Settings | None = None


def get_settings() -> Settings:
    if _settings is None:
        raise RuntimeError(
            "Settings not initialised — build_app() must run first."
        )
    return _settings


def set_settings(settings: Settings | None) -> None:
    global _settings
    _settings = settings


_oauth_client: OAuthClient | None = None
_token_store: TokenStore | None = None


def get_oauth_client() -> OAuthClient:
    """Return the OAuth client or raise — call after :func:`build_app`.

    Raises :class:`RuntimeError` when the Azure credentials are not configured;
    callers should map that to a 503 for the ``/auth/*`` endpoints.
    """
    if _oauth_client is None:
        raise RuntimeError(
            "OAuthClient not initialised — Azure credentials missing or "
            "build_app() did not run."
        )
    return _oauth_client


def set_oauth_client(client: OAuthClient | None) -> None:
    global _oauth_client
    _oauth_client = client


def get_token_store() -> TokenStore:
    """Return the process-wide token store. Always set by :func:`build_app`."""
    if _token_store is None:
        raise RuntimeError("TokenStore not initialised — build_app() must run first.")
    return _token_store


def set_token_store(store: TokenStore | None) -> None:
    global _token_store
    _token_store = store


def download_url_for(meta: FileMetadata, user_id: str) -> str | None:
    """Return the user-facing URL for a file, preferring its native viewer.

    For backends that surface a browser URL (``meta.web_url`` — e.g.
    OneDrive's ``webUrl``), that is returned as-is so the user lands in
    the file's native viewer with their existing session. For backends
    without one (``local://``), falls back to a signed ``/files/{token}``
    URL via :func:`build_download_url`.
    """
    if meta.web_url:
        return meta.web_url
    return build_download_url(meta.uri, user_id)


def build_download_url(uri: str, user_id: str) -> str | None:
    """Return a ``/files/{token}`` URL if the server is configured for it.

    Requires both ``signing_secret`` and ``download_base_url`` in
    :class:`Settings`; returns ``None`` otherwise so tools can silently
    omit the ``download_url`` field when the operator has not enabled it.

    For cloud-backed URIs (currently ``od://``) the download route
    needs an OAuth token to fetch bytes from Microsoft Graph. We embed
    the user's current OAuth ``state`` (token-store key) into the
    signed claim so ``/files/{token}`` can resolve the right token at
    request time. If no state is bound to the current MCP session the
    download URL is still returned, but it will only succeed for URIs
    whose backend doesn't need a per-user token (i.e. ``local://``).
    """
    settings = _settings
    if settings is None or not settings.signing_secret or not settings.download_base_url:
        return None
    state: str | None = None
    if uri.startswith("od://"):
        # Lazy lookup so unit tests that don't wire context/store still
        # exercise the local:// path.
        from office_mcp.context import current_mcp_session_id

        session_id = current_mcp_session_id()
        if session_id and _token_store is not None:
            state = _token_store.get_state_for_session(session_id)
    token = signing.sign(
        user_id=user_id,
        uri=uri,
        secret=settings.signing_secret,
        ttl_seconds=settings.download_url_ttl_seconds,
        state=state,
    )
    return f"{settings.download_base_url.rstrip('/')}/files/{token}"
