"""Runtime configuration loaded from environment variables (and optional .env)."""

from __future__ import annotations

from pathlib import Path

from pydantic_settings import BaseSettings, SettingsConfigDict


class Settings(BaseSettings):
    """All ``OFFICE_MCP_*`` environment variables.

    LibreChat-specific variables (``CREDS_KEY``, ``MINIMAX_*`` …) live in the
    LibreChat container's environment and are intentionally ignored here.
    """

    host: str = "0.0.0.0"  # noqa: S104 — listening on all interfaces inside the container is intentional
    port: int = 8080
    log_level: str = "INFO"

    token: str | None = None
    """Optional shared bearer required on every non-/health request."""

    local_root: Path = Path("/data")
    max_file_size_mb: int = 50
    rate_limit_per_min: int = 60

    default_user_id: str = "demo-user"
    """Fallback user scope used when the client did not send a usable
    ``X-LibreChat-User-Id`` header — or sent a literal ``{{...}}`` placeholder
    because LibreChat 0.8.4 does not substitute headers in chat mode
    (see danny-avila/LibreChat#9910). Set to a non-generic value in multi-user
    deployments to force clients to supply a real user id.
    """

    signing_secret: str | None = None
    """HMAC key for signed ``/files/{token}`` download URLs. When unset, no
    ``download_url`` is returned in tool responses. Typically set to the
    same value as ``OFFICE_MCP_TOKEN``."""

    download_base_url: str | None = None
    """Externally-reachable base URL for the ``/files/{token}`` endpoint.
    Example: ``http://pixel.lan:18080``. When unset, no ``download_url``
    is returned."""

    download_url_ttl_seconds: int = 900
    """Lifetime of signed download tokens. Short values reduce replay risk;
    long values tolerate slow readers. Default: 15 min."""

    model_config = SettingsConfigDict(
        env_prefix="OFFICE_MCP_",
        env_file=".env",
        env_file_encoding="utf-8",
        extra="ignore",
        case_sensitive=False,
    )
