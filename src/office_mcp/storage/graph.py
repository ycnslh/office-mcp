"""Microsoft Graph storage backend — OneDrive (and SharePoint later).

Talks to ``https://graph.microsoft.com/v1.0`` using a delegated access
token taken from :attr:`UserContext.on_behalf_of_token`. The lookup of
that token from the in-process :class:`~office_mcp.auth.store.TokenStore`
is the *caller's* responsibility (Phase 7c) — this backend is purely
the HTTP wiring against Graph and the URI / metadata mapping.

URI scheme is ``od://`` for OneDrive (the user's own drive). Examples:

* ``od://Documents/q3.docx``
* ``od://Apps/office-mcp/report.pdf``

User scoping comes from the access token itself: Microsoft Graph maps
``/me/drive/...`` to whichever user the token was issued for. We never
need to embed a ``user_id`` in the URI — a request with user A's token
cannot observe user B's drive.

Small files only in this first pass: writes go through PUT to the
``:/content`` endpoint, which Graph caps at 4 MB. Larger files require
an upload session — TODO for a follow-up phase.
"""

from __future__ import annotations

import asyncio
from collections.abc import AsyncIterator
from typing import Any
from urllib.parse import quote

import httpx
import structlog

from office_mcp.context import UserContext
from office_mcp.storage.base import FileMetadata, StorageBackend
from office_mcp.utils.errors import (
    AuthRequiredError,
    BackendUnavailableError,
    InvalidUriError,
    PathTraversalError,
    PreconditionFailedError,
    StorageNotFoundError,
)

log = structlog.get_logger(__name__)

_SCHEME = "od"
_GRAPH_BASE = "https://graph.microsoft.com/v1.0"
_DEFAULT_TIMEOUT = httpx.Timeout(connect=10.0, read=60.0, write=60.0, pool=5.0)
#: Graph rejects PUT to ``:/content`` past this size. Larger files need
#: an upload session, which we'll wire in a follow-up phase.
_MAX_SIMPLE_PUT_BYTES = 4 * 1024 * 1024

#: Microsoft Graph briefly locks an Office file after a write while
#: SharePoint propagates and any open Office Online preview releases
#: it (HTTP 423 ``resourceLocked``). Rapid successive edits from the
#: LLM trip this in practice — we retry transparently with backoff
#: rather than asking the model to recover, which it does poorly.
#: The previous (1, 2, 4) schedule (~7 s window) was too short: in
#: production we observed locks persisting 8–30 s after a fresh
#: create while SharePoint runs AV scan + preview generation. The
#: cumulative ~31 s window stays well under LibreChat's 120 s MCP
#: tool timeout while covering most real lock durations.
_RETRY_LOCKED_DELAYS_S: tuple[float, ...] = (1.0, 2.0, 4.0, 8.0, 16.0)


class GraphBackend(StorageBackend):
    """OneDrive backend using delegated tokens via Microsoft Graph."""

    scheme = _SCHEME

    def __init__(
        self,
        *,
        client: httpx.AsyncClient | None = None,
        timeout: httpx.Timeout = _DEFAULT_TIMEOUT,
        default_folder: str | None = None,
    ) -> None:
        # Tests inject a pre-configured client (e.g. with pytest-httpx
        # transport hooks). Production gets a default httpx instance
        # with sane connect / read / pool timeouts.
        self._client = client if client is not None else httpx.AsyncClient(timeout=timeout)
        self._owns_client = client is None
        # When set, writes targeting the OneDrive root (``od://foo.docx``,
        # no folder segment) get rewritten into ``od://<default>/foo.docx``
        # so LLM-generated documents land in a predictable subfolder
        # rather than at the user's drive root. Reads/lists/deletes
        # respect the URI as-is — only the create/write path normalizes.
        self._default_folder = (default_folder or "").strip("/") or None

    async def aclose(self) -> None:
        """Close the underlying HTTP client if we created it."""
        if self._owns_client:
            await self._client.aclose()

    # ── URI helpers ─────────────────────────────────────────────────

    @staticmethod
    def _path_from_uri(uri: str) -> str:
        if not uri.startswith(f"{_SCHEME}://"):
            raise InvalidUriError(f"expected {_SCHEME}:// URI, got {uri!r}")
        path = uri.removeprefix(f"{_SCHEME}://").lstrip("/")
        if not path or "\x00" in path:
            raise InvalidUriError(f"invalid URI path: {uri!r}")
        # Reject ``..`` segments and empty segments (``foo//bar``). Empty
        # segments collapse silently in PurePosixPath.parts but Graph
        # rejects them with a confusing 404, so fail closed at parse time.
        segments = path.split("/")
        if any(s == ".." or s == "" for s in segments):
            raise PathTraversalError(
                f"URI contains '..' or empty segment: {uri!r}"
            )
        return path

    @staticmethod
    def _drive_url(path: str, *, suffix: str = "") -> str:
        # /me/drive/root:/<encoded-path>:/<suffix>
        # `/` between path components must NOT be encoded — encode each
        # segment individually so file names with spaces or accents work.
        encoded = "/".join(quote(part, safe="") for part in path.split("/"))
        if suffix:
            return f"{_GRAPH_BASE}/me/drive/root:/{encoded}:/{suffix}"
        return f"{_GRAPH_BASE}/me/drive/root:/{encoded}"

    @staticmethod
    def _require_token(user_ctx: UserContext) -> str:
        token = user_ctx.on_behalf_of_token
        if not token:
            raise AuthRequiredError(
                "Microsoft Graph operation requires user authentication. "
                "Direct the user to /auth/login to start the OAuth flow."
            )
        return token

    @staticmethod
    def _to_metadata(uri: str, item: dict[str, Any]) -> FileMetadata:
        file_facet = item.get("file") or {}
        # Graph returns both ``eTag`` (item-wide, includes metadata changes)
        # and ``cTag`` (content-only). For optimistic concurrency on writes
        # we use eTag — that's what ``If-Match`` matches against.
        web_url = item.get("webUrl")
        return FileMetadata(
            uri=uri,
            size_bytes=int(item.get("size", 0)),
            modified_at=str(item.get("lastModifiedDateTime", "")),
            etag=str(item.get("eTag")) if item.get("eTag") else None,
            mime_type=file_facet.get("mimeType"),
            web_url=str(web_url) if web_url else None,
        )

    def _normalize_write_uri(self, uri: str) -> str:
        """Promote a root-level ``od://foo.ext`` URI into the default folder.

        Only fires on URIs that have no path separator (i.e. the LLM
        forgot to include a folder). Anything with an explicit folder
        is passed through unchanged.
        """
        if not self._default_folder or not uri.startswith(f"{_SCHEME}://"):
            return uri
        raw = uri.removeprefix(f"{_SCHEME}://").lstrip("/")
        if "/" in raw or not raw:
            return uri
        return f"{_SCHEME}://{self._default_folder}/{raw}"

    @staticmethod
    def _raise_for_status(response: httpx.Response, uri: str) -> None:
        status = response.status_code
        if status < 400:
            return
        if status == 404:
            raise StorageNotFoundError(uri)
        if status == 412:
            raise PreconditionFailedError(f"ETag mismatch for {uri}")
        if status in (401, 403):
            raise AuthRequiredError(
                f"Microsoft Graph rejected the access token for {uri} "
                f"(status {status})"
            )
        if status == 429:
            # Graph throttles delegated /me/drive aggressively. Surface
            # the Retry-After hint so the LLM can wait or back off.
            retry_after = response.headers.get("Retry-After", "unspecified")
            raise BackendUnavailableError(
                f"Microsoft Graph throttled the request for {uri} "
                f"(retry after {retry_after} seconds)"
            )
        if status == 507:
            raise BackendUnavailableError(
                f"OneDrive quota exhausted for {uri}"
            )
        if status == 423:
            # The bounded retry in ``_retry_on_423`` exhausted itself —
            # OneDrive is still holding the lock (typically AV scan or
            # Office Online preview generation after a fresh write).
            # Spell out the cause so the LLM waits-and-retries instead
            # of cloning the file under a v2 name.
            raise BackendUnavailableError(
                f"OneDrive is still indexing {uri} (post-write lock held >30 s). "
                f"Wait ~30 seconds and retry the same file_uri — do not rename "
                f"or create a v2 copy."
            )
        raise BackendUnavailableError(
            f"Microsoft Graph returned {status} for {uri}: {response.text[:200]}"
        )

    # ── StorageBackend API ──────────────────────────────────────────

    async def read(self, uri: str, *, user_ctx: UserContext) -> bytes:
        token = self._require_token(user_ctx)
        path = self._path_from_uri(uri)
        url = self._drive_url(path, suffix="content")
        response = await self._client.get(
            url,
            headers={"Authorization": f"Bearer {token}"},
            # Graph 302s the bytes onto a short-lived CDN URL.
            follow_redirects=True,
        )
        self._raise_for_status(response, uri)
        return response.content

    async def write(
        self,
        uri: str,
        data: bytes,
        *,
        user_ctx: UserContext,
        if_match: str | None = None,
    ) -> FileMetadata:
        # Order matches read/stat/delete: token check, then URI parse,
        # then size check. Keeps the error class surfaced for an
        # unauthenticated bad-URI call consistent across operations.
        token = self._require_token(user_ctx)
        effective_uri = self._normalize_write_uri(uri)
        if effective_uri != uri:
            log.info(
                "graph.write.default_folder_applied",
                requested_uri=uri,
                effective_uri=effective_uri,
            )
        path = self._path_from_uri(effective_uri)
        if len(data) > _MAX_SIMPLE_PUT_BYTES:
            raise BackendUnavailableError(
                f"file size {len(data)} exceeds simple-PUT cap "
                f"({_MAX_SIMPLE_PUT_BYTES} bytes); upload sessions not yet wired"
            )
        url = self._drive_url(path, suffix="content")
        headers: dict[str, str] = {
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/octet-stream",
        }
        if if_match is not None:
            headers["If-Match"] = if_match
        response = await self._retry_on_423(
            "PUT",
            url,
            effective_uri,
            content=data,
            headers=headers,
        )
        self._raise_for_status(response, effective_uri)
        return self._to_metadata(effective_uri, response.json())

    async def _retry_on_423(
        self,
        method: str,
        url: str,
        uri: str,
        **request_kwargs: Any,
    ) -> httpx.Response:
        """Issue a Graph request, retrying through the brief 423 lock window.

        OneDrive/SharePoint locks an Office file for a few seconds
        after each write while it propagates and any open Office Online
        preview lets go. The lock returns HTTP 423 ``resourceLocked``;
        a small bounded retry transparently waits it out instead of
        bubbling a confusing error up to the LLM, which would otherwise
        clone the file and rename it (we observed that with MiniMax in
        production).
        """
        for delay in (0.0, *_RETRY_LOCKED_DELAYS_S):
            if delay:
                await asyncio.sleep(delay)
            response = await self._client.request(method, url, **request_kwargs)
            if response.status_code != 423:
                return response
            log.info(
                "graph.resource_locked.retry",
                uri=uri,
                method=method,
                next_delay_s=delay,
            )
        return response

    async def stat(self, uri: str, *, user_ctx: UserContext) -> FileMetadata:
        token = self._require_token(user_ctx)
        path = self._path_from_uri(uri)
        response = await self._client.get(
            self._drive_url(path),
            headers={"Authorization": f"Bearer {token}"},
        )
        self._raise_for_status(response, uri)
        return self._to_metadata(uri, response.json())

    async def list(
        self,
        uri_prefix: str,
        *,
        user_ctx: UserContext,
    ) -> AsyncIterator[FileMetadata]:
        token = self._require_token(user_ctx)
        starting_path = (
            ""
            if uri_prefix in (f"{_SCHEME}://", f"{_SCHEME}:///")
            else self._path_from_uri(uri_prefix)
        )

        # BFS through folders so the result is recursive (matches the
        # StorageBackend contract). Files are yielded as we go; folders
        # are queued so the next iteration descends into them.
        folders_to_visit: list[str] = [starting_path]
        while folders_to_visit:
            folder_path = folders_to_visit.pop(0)
            url = (
                f"{_GRAPH_BASE}/me/drive/root/children"
                if not folder_path
                else self._drive_url(folder_path, suffix="children")
            )
            while url:
                response = await self._client.get(
                    url, headers={"Authorization": f"Bearer {token}"}
                )
                self._raise_for_status(response, uri_prefix)
                payload = response.json()
                for item in payload.get("value", []):
                    name = item["name"]
                    child_path = f"{folder_path}/{name}" if folder_path else name
                    if "folder" in item:
                        folders_to_visit.append(child_path)
                    elif "file" in item:
                        yield self._to_metadata(f"{_SCHEME}://{child_path}", item)
                url = payload.get("@odata.nextLink", "")

    async def delete(self, uri: str, *, user_ctx: UserContext) -> None:
        token = self._require_token(user_ctx)
        path = self._path_from_uri(uri)
        # OneDrive holds the same post-write 423 lock on DELETE as on
        # PUT — the file the LLM just created stays locked for tens of
        # seconds while SharePoint indexes it, and a delete issued in
        # that window fails. Reuse the same retry strategy as write().
        response = await self._retry_on_423(
            "DELETE",
            self._drive_url(path),
            uri,
            headers={"Authorization": f"Bearer {token}"},
        )
        self._raise_for_status(response, uri)

    async def health_check(self) -> str:
        # /v1.0/$metadata is anonymous and confirms Graph is reachable.
        # We don't have a user token at health-check time (it's a
        # process-wide probe), so we can only verify the endpoint is up.
        try:
            response = await self._client.get(
                f"{_GRAPH_BASE}/$metadata", timeout=5.0
            )
        except httpx.HTTPError as exc:
            raise RuntimeError(f"Graph unreachable: {exc}") from exc
        if not (200 <= response.status_code < 300):
            raise RuntimeError(f"Graph $metadata returned {response.status_code}")
        return "ok"
