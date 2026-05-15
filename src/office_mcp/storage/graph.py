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
from datetime import UTC, datetime
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


def _extract_bearer(headers: Any) -> str | None:
    """Pull the bearer token out of an Authorization header dict.

    Used to keep the diagnostic GET in ``_retry_on_423`` self-contained
    without changing the signature of every caller.
    """
    if not isinstance(headers, dict):
        return None
    auth = headers.get("Authorization")
    if not isinstance(auth, str) or not auth.startswith("Bearer "):
        return None
    return auth.removeprefix("Bearer ").strip() or None

_SCHEME = "od"
_GRAPH_BASE = "https://graph.microsoft.com/v1.0"
_DEFAULT_TIMEOUT = httpx.Timeout(connect=10.0, read=60.0, write=60.0, pool=5.0)

#: Cap on single-shot upload-session uploads. Microsoft documents
#: 60 MiB as the max for a single ``Content-Range`` PUT against an
#: ``uploadUrl``; beyond that we'd need to split the bytes across
#: multiple ranged PUTs (TODO when a real use case lands). Until then
#: anything over the cap is rejected at the API boundary with a clear
#: error rather than a vague "Request entity too large" from Graph.
_MAX_UPLOAD_SESSION_BYTES = 60 * 1024 * 1024

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

#: A 423 on a file whose ``lastModifiedDateTime`` is older than this
#: cannot be the post-write indexing lock (that window is tens of
#: seconds at worst). The remaining cause is an external coauthoring
#: session — Word Online, OneDrive desktop, Office Mobile — holding
#: an exclusive handle. No amount of server-side retry will clear it:
#: we short-circuit with an actionable message instead of burning the
#: 31 s retry budget for nothing.
_EXTERNAL_LOCK_AGE_S: float = 90.0


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
        """Write via the Graph upload-session API.

        We deliberately use ``POST /createUploadSession`` + a ranged
        ``PUT`` to the returned ``uploadUrl`` *even for tiny files*.
        The simple ``PUT /me/drive/root:/<path>:/content`` endpoint
        races against Office for the web: if the user has the same
        file open in Word Online, every simple-PUT returns HTTP 423
        for the entire duration of the WOPI coauthoring session.

        The upload-session API takes the WOPI-aware code path on
        Microsoft's side and is the official way to push bytes into a
        file that may be concurrently open in Office. Trades 1 extra
        roundtrip (the session creation) for not failing when the
        user is looking at the file.
        """
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
        size = len(data)
        if size == 0:
            # The ``Content-Range`` header for a 0-byte upload would be
            # ``bytes 0--1/0`` which Graph rejects. We never expect a
            # legitimate zero-byte write from the engines (even an
            # empty .docx is several KB), so fail closed at the API
            # boundary with a clear cause.
            raise BackendUnavailableError(
                f"refusing to upload zero bytes to {effective_uri}"
            )
        if size > _MAX_UPLOAD_SESSION_BYTES:
            raise BackendUnavailableError(
                f"file size {size} exceeds single-shot upload-session "
                f"cap ({_MAX_UPLOAD_SESSION_BYTES} bytes); chunked upload "
                f"not yet wired"
            )

        # Step 1 — open the upload session. Still wrapped in _retry_on_423
        # because Graph can occasionally 423 the session-create endpoint
        # itself; the same classifier triggers if it's a coauthoring lock.
        session_url = self._drive_url(path, suffix="createUploadSession")
        session_body: dict[str, Any] = {
            "item": {
                # "replace" matches the previous simple-PUT semantics.
                # The upload-session call doesn't honor ``If-Match``
                # itself — we apply it on the final ranged PUT below.
                "@microsoft.graph.conflictBehavior": "replace",
                "name": path.split("/")[-1],
            },
        }
        session_response = await self._retry_on_423(
            "POST",
            session_url,
            effective_uri,
            json=session_body,
            headers={"Authorization": f"Bearer {token}"},
        )
        self._raise_for_status(session_response, effective_uri)
        upload_url = session_response.json().get("uploadUrl")
        if not isinstance(upload_url, str) or not upload_url:
            raise BackendUnavailableError(
                f"Graph createUploadSession for {effective_uri} returned "
                f"no uploadUrl"
            )

        # Step 2 — single-shot ranged PUT to the pre-signed uploadUrl.
        # CRITICAL: the uploadUrl is pre-signed; do NOT add an
        # Authorization header (Microsoft docs explicitly forbid it).
        upload_headers: dict[str, str] = {
            "Content-Length": str(size),
            "Content-Range": f"bytes 0-{size - 1}/{size}",
        }
        if if_match is not None:
            upload_headers["If-Match"] = if_match
        upload_response = await self._client.put(
            upload_url,
            content=data,
            headers=upload_headers,
        )
        self._raise_for_status(upload_response, effective_uri)
        return self._to_metadata(effective_uri, upload_response.json())

    async def _retry_on_423(
        self,
        method: str,
        url: str,
        uri: str,
        **request_kwargs: Any,
    ) -> httpx.Response:
        """Issue a Graph request, retrying through the brief 423 lock window.

        Two distinct lock conditions surface as HTTP 423 from Graph:

        1. **Transient post-write lock** — SharePoint runs AV scan and
           preview generation after every successful write. Lasts tens
           of seconds; absorbed transparently via the retry schedule.
        2. **External coauthoring lock** — Word Online, OneDrive
           desktop, or Office Mobile holds an exclusive handle on the
           file. Lasts as long as the external session stays open
           (minutes to hours). No server-side retry will release it.

        On the first 423 we probe the file's ``lastModifiedDateTime``
        to distinguish the two: a file modified seconds ago is in (1),
        a file modified minutes ago and still 423-ing is in (2). For
        (2) we short-circuit with an actionable error so the LLM
        relays the real fix ("close the open Word tab") instead of
        wasting the retry budget.
        """
        token = _extract_bearer(request_kwargs.get("headers"))
        delays = (0.0, *_RETRY_LOCKED_DELAYS_S)
        last_idx = len(delays) - 1
        for idx, delay in enumerate(delays):
            if delay:
                await asyncio.sleep(delay)
            response = await self._client.request(method, url, **request_kwargs)
            if response.status_code != 423:
                return response
            # Probe once on the first 423. Subsequent retries don't
            # re-probe — if it was transient on the first check it stays
            # transient. Short-circuits when either condition before the
            # await is false, so the GET only fires when really needed.
            if (
                idx == 0
                and token is not None
                and await self._is_external_lock(uri, token, response)
            ):
                raise BackendUnavailableError(
                    f"{uri} is currently open in Word Online, OneDrive "
                    f"desktop, or another Office client (coauthoring "
                    f"lock). Close the open editor session and retry — "
                    f"the server cannot release this lock for you."
                )
            if idx == last_idx:
                # Retry budget exhausted — surface the final 423 to the
                # caller so ``_raise_for_status`` turns it into the
                # "still indexing" message.
                return response
            log.info(
                "graph.resource_locked.retry",
                uri=uri,
                method=method,
                next_delay_s=delay,
            )
        # Unreachable: the loop always returns on either success, external
        # lock, or exhausted budget. This satisfies the type checker.
        raise RuntimeError("retry loop completed without returning")

    async def _is_external_lock(
        self,
        uri: str,
        token: str,
        triggering_response: httpx.Response,
    ) -> bool:
        """Classify a 423 as external (coauthoring) vs transient (post-write).

        Heuristic: fetch the driveItem and inspect
        ``lastModifiedDateTime``. If the file was modified more than
        ``_EXTERNAL_LOCK_AGE_S`` seconds ago, the post-write window has
        long passed — the only remaining cause for a 423 is an external
        lock. Anything younger gets classified as transient (retry).

        Probe failures (404, network blip, malformed payload) fall back
        to "transient" so we don't break the normal retry path on a
        flaky lookup.
        """
        try:
            path = self._path_from_uri(uri)
        except (InvalidUriError, PathTraversalError):
            return False
        try:
            probe = await self._client.get(
                self._drive_url(path),
                headers={"Authorization": f"Bearer {token}"},
            )
        except httpx.RequestError:
            return False
        if probe.status_code != 200:
            return False
        try:
            payload = probe.json()
        except ValueError:
            return False
        modified_raw = payload.get("lastModifiedDateTime")
        if not isinstance(modified_raw, str):
            return False
        try:
            modified_at = datetime.fromisoformat(
                modified_raw.replace("Z", "+00:00")
            )
        except ValueError:
            return False
        age_s = (datetime.now(UTC) - modified_at).total_seconds()
        if age_s <= _EXTERNAL_LOCK_AGE_S:
            return False
        log.info(
            "graph.resource_locked.external_classified",
            uri=uri,
            modified_age_s=age_s,
            response_body=triggering_response.text[:300],
        )
        return True

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

    async def preview_url(self, uri: str, *, user_ctx: UserContext) -> str | None:
        """Return a read-only Office for the web URL via Graph's ``/preview``.

        ``POST /me/drive/root:/<path>:/preview`` returns ``{"getUrl": ...}``
        — a long-lived URL that opens the document in Word/Excel/PowerPoint
        for the web in **read-only** mode. Read-only sessions do not
        acquire a WOPI write lock, so handing this URL to the user does
        NOT block subsequent server-side edits (the way ``meta.web_url``,
        which is the *edit* URL, does).

        If the user wants to switch to edit mode they can click "Edit
        Document" inside the viewer — that acquires the lock and our
        edits are blocked until they close the tab. Documented limitation.

        Returns ``None`` on any failure (no token, parse error, Graph
        non-200) so callers can fall back to a signed download URL.
        """
        token = user_ctx.on_behalf_of_token
        if not token:
            return None
        try:
            path = self._path_from_uri(uri)
        except (InvalidUriError, PathTraversalError):
            return None
        try:
            response = await self._client.post(
                self._drive_url(path, suffix="preview"),
                json={},
                headers={"Authorization": f"Bearer {token}"},
            )
        except httpx.RequestError:
            return None
        if response.status_code != 200:
            return None
        try:
            payload = response.json()
        except ValueError:
            return None
        get_url = payload.get("getUrl")
        if isinstance(get_url, str) and get_url:
            return get_url
        return None

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
