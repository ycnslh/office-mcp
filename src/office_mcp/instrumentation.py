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
from office_mcp.context import current_user
from office_mcp.runtime import mcp

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
    """Wrap ``fn`` so each call logs + records Prometheus metrics."""
    tool_name = fn.__name__

    @functools.wraps(fn)
    async def wrapper(*args: Any, **kwargs: Any) -> Any:
        user = current_user()
        log = _logger().bind(
            tool=tool_name,
            user_id=user.user_id if user else None,
        )
        start = time.perf_counter()
        try:
            result = await fn(*args, **kwargs)
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


def register_tool(fn: F) -> F:
    """Register ``fn`` with FastMCP through the instrumentation wrapper.

    Returns the unwrapped function so the module-level name stays
    directly callable from tests.
    """
    mcp.tool(instrumented(fn))
    return fn
