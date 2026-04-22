"""Scheme → backend routing."""

from __future__ import annotations

from office_mcp.storage.base import StorageBackend
from office_mcp.utils.errors import InvalidUriError, UnknownSchemeError


class StorageRegistry:
    """Looks up the storage backend that owns a given URI."""

    def __init__(self) -> None:
        self._backends: dict[str, StorageBackend] = {}

    def register(self, backend: StorageBackend) -> None:
        self._backends[backend.scheme] = backend

    def get(self, uri: str) -> StorageBackend:
        if "://" not in uri:
            raise InvalidUriError(f"URI missing scheme: {uri!r}")
        scheme = uri.split("://", 1)[0]
        backend = self._backends.get(scheme)
        if backend is None:
            raise UnknownSchemeError(scheme)
        return backend
