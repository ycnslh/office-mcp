"""Error hierarchy for office-mcp.

All runtime errors surfaced by this package inherit from :class:`OfficeMcpError`.
Storage errors have their own branch so the MCP layer can map them to clean
tool responses (404 / 409 / 400 equivalents) without leaking internals.
"""

from __future__ import annotations


class OfficeMcpError(Exception):
    """Base class for every error raised by office-mcp."""


class StorageError(OfficeMcpError):
    """Base class for storage-backend errors."""


class InvalidUriError(StorageError):
    """The supplied URI is malformed or uses an unsupported scheme."""


class UnknownSchemeError(StorageError):
    """No storage backend is registered for this URI scheme."""


class PathTraversalError(StorageError):
    """The resolved path escapes the user's storage scope."""


class StorageNotFoundError(StorageError):
    """The requested resource does not exist."""


class PreconditionFailedError(StorageError):
    """Optimistic concurrency check (``if_match`` / ETag) did not match."""
