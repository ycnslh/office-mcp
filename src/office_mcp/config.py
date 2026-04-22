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

    model_config = SettingsConfigDict(
        env_prefix="OFFICE_MCP_",
        env_file=".env",
        env_file_encoding="utf-8",
        extra="ignore",
        case_sensitive=False,
    )
