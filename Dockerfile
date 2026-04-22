# syntax=docker/dockerfile:1.7
# ─────────────────────────────────────────────────────────────────────
# office-mcp — multi-stage build
#  - builder: install poetry + project deps into a venv
#  - runtime: slim python image with the venv copied in
# LibreOffice headless is added in Phase 4 (xlsx formula recalc + pdf).
# ─────────────────────────────────────────────────────────────────────

ARG PYTHON_VERSION=3.12

FROM python:${PYTHON_VERSION}-slim AS builder

ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    PIP_NO_CACHE_DIR=1 \
    POETRY_VERSION=1.8.4 \
    POETRY_VIRTUALENVS_IN_PROJECT=true \
    POETRY_NO_INTERACTION=1

RUN apt-get update \
 && apt-get install -y --no-install-recommends build-essential curl \
 && rm -rf /var/lib/apt/lists/*

RUN pip install "poetry==${POETRY_VERSION}"

WORKDIR /app
COPY pyproject.toml ./
COPY README.md ./
RUN poetry install --only main --no-root --no-ansi

COPY src ./src
RUN poetry install --only main --no-ansi

# ─────────────────────────────────────────────────────────────────────

FROM python:${PYTHON_VERSION}-slim AS runtime

ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    PATH="/app/.venv/bin:${PATH}" \
    OFFICE_MCP_HOST=0.0.0.0 \
    OFFICE_MCP_PORT=8080 \
    OFFICE_MCP_LOCAL_ROOT=/data

RUN apt-get update \
 && apt-get install -y --no-install-recommends curl \
 && rm -rf /var/lib/apt/lists/* \
 && groupadd --system app \
 && useradd --system --gid app --create-home --home-dir /home/app app \
 && mkdir -p /data \
 && chown -R app:app /data

WORKDIR /app
COPY --from=builder /app /app
RUN chown -R app:app /app

USER app
VOLUME ["/data"]
EXPOSE 8080

HEALTHCHECK --interval=30s --timeout=5s --retries=3 \
  CMD curl -fsS http://localhost:${OFFICE_MCP_PORT}/health || exit 1

ENTRYPOINT ["python", "-m", "office_mcp"]
