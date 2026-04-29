"""Logging + metrics instrumentation around MCP tools.

Tool modules call :func:`register_tool` instead of ``mcp.tool(fn)``
directly. The wrapper measures duration, increments Prometheus
counters, and emits a structured log line — all without forcing every
tool to know about logging or metrics.

The original function reference stays callable from tests (we only
register the wrapper with FastMCP).
"""

from __future__ import annotations

import functools
import logging
import time
from collections.abc import Callable
from typing import Any, TypeVar

import structlog

from office_mcp import metrics
from office_mcp.context import current_mcp_session_id, current_user
from office_mcp.runtime import mcp
from office_mcp.utils.errors import AuthRequiredError

F = TypeVar("F", bound=Callable[..., Any])


_LOGGING_CONFIGURED = False


def setup_logging(level: str = "INFO") -> None:
    """Configure structlog + stdlib logging for JSON output once per process."""
    global _LOGGING_CONFIGURED
    if _LOGGING_CONFIGURED:
        return

    log_level = getattr(logging, level.upper(), logging.INFO)
    logging.basicConfig(format="%(message)s", level=log_level)
    structlog.configure(
        processors=[
            structlog.contextvars.merge_contextvars,
            structlog.processors.add_log_level,
            structlog.processors.TimeStamper(fmt="iso"),
            structlog.processors.StackInfoRenderer(),
            structlog.dev.set_exc_info,
            structlog.processors.JSONRenderer(),
        ],
        wrapper_class=structlog.make_filtering_bound_logger(log_level),
        cache_logger_on_first_use=True,
    )
    _LOGGING_CONFIGURED = True


def _logger() -> Any:
    return structlog.get_logger("office_mcp.tool")


def instrumented(fn: F) -> F:
    """Wrap ``fn`` so each call logs + records Prometheus metrics.

    Also intercepts :class:`AuthRequiredError` from cloud backends and
    converts it into a structured ``auth_required`` response shape with
    a Microsoft login URL. Phase 7c — without this, every Graph-touching
    tool would surface a raw exception to the LLM instead of the URL
    the user needs to click.
    """
    tool_name = fn.__name__

    @functools.wraps(fn)
    async def wrapper(*args: Any, **kwargs: Any) -> Any:
        user = current_user()
        log = _logger().bind(
            tool=tool_name,
            user_id=user.user_id if user else None,
            args=_sanitise_args(kwargs),
        )
        start = time.perf_counter()
        try:
            result = await fn(*args, **kwargs)
        except AuthRequiredError as exc:
            duration = time.perf_counter() - start
            metrics.tool_calls.labels(tool=tool_name, status="auth_required").inc()
            metrics.tool_duration_seconds.labels(tool=tool_name).observe(duration)
            log.info(
                "tool_auth_required",
                duration_ms=round(duration * 1000, 2),
                detail=str(exc)[:200],
            )
            # Local imports: ``flow`` pulls runtime helpers, which import
            # the auth package — keeping these lazy avoids the cycle.
            from office_mcp.auth.flow import auth_required_response
            from office_mcp.auth.store import PendingStoreFull

            try:
                return auth_required_response(current_mcp_session_id())
            except RuntimeError:
                # OAuth not configured (no Azure creds) — surface a clear
                # message instead of a misleading "click here" URL.
                return {
                    "auth_required": True,
                    "summary": (
                        "Cloud backend access requires Microsoft auth, "
                        "but Azure AD is not configured on this server."
                    ),
                }
            except PendingStoreFull:
                # Pending-claim store at capacity — same shape as the
                # other auth_required branches so the LLM doesn't have
                # to learn a third response variant.
                return {
                    "auth_required": True,
                    "summary": (
                        "Trop de tentatives d'authentification simultanées "
                        "sur ce serveur. Réessayez dans quelques instants."
                    ),
                }
        except Exception as exc:
            duration = time.perf_counter() - start
            metrics.tool_calls.labels(tool=tool_name, status="error").inc()
            metrics.tool_duration_seconds.labels(tool=tool_name).observe(duration)
            log.warning(
                "tool_failed",
                duration_ms=round(duration * 1000, 2),
                error_type=type(exc).__name__,
                error_message=str(exc)[:400],
            )
            raise
        duration = time.perf_counter() - start
        metrics.tool_calls.labels(tool=tool_name, status="ok").inc()
        metrics.tool_duration_seconds.labels(tool=tool_name).observe(duration)
        log.info(
            "tool_succeeded",
            duration_ms=round(duration * 1000, 2),
        )
        return result

    return wrapper  # type: ignore[return-value]


#: Cap on the visible portion of a string-valued arg in the structured log.
#: Long enough to keep URIs and short prompts intact, short enough that a
#: stray document body doesn't blow up the log line.
_ARG_PREVIEW_CHARS = 120


def _sanitise_args(kwargs: dict[str, Any]) -> dict[str, Any]:
    """Render a structured-log-friendly preview of a tool's kwargs.

    The goal is diagnostic visibility — *what* the LLM passed and roughly
    how big it was — without dumping document text or base64 images into
    the logs. Strings get truncated; lists are summarised by length and a
    sample of element types; bytes-like values are summarised by length;
    other shapes show their type name.
    """
    out: dict[str, Any] = {}
    for key, value in kwargs.items():
        if value is None:
            out[key] = None
        elif isinstance(value, bool):
            out[key] = value
        elif isinstance(value, int | float):
            out[key] = value
        elif isinstance(value, str):
            if len(value) <= _ARG_PREVIEW_CHARS:
                out[key] = value
            else:
                out[key] = f"{value[:_ARG_PREVIEW_CHARS]}…(+{len(value) - _ARG_PREVIEW_CHARS} chars)"
        elif isinstance(value, list | tuple):
            element_types = sorted({type(v).__name__ for v in value})
            out[key] = f"<{type(value).__name__} len={len(value)} types={element_types}>"
        elif isinstance(value, bytes | bytearray):
            out[key] = f"<{type(value).__name__} len={len(value)}>"
        elif isinstance(value, dict):
            out[key] = f"<dict keys={sorted(value)}>"
        else:
            out[key] = f"<{type(value).__name__}>"
    return out


def register_tool(fn: F) -> F:
    """Register ``fn`` with FastMCP through the instrumentation wrapper.

    Returns the unwrapped function so the module-level name stays
    directly callable from tests.
    """
    mcp.tool(instrumented(fn))
    return fn
