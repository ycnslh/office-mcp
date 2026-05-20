# syntax=docker/dockerfile:1.7
# ─────────────────────────────────────────────────────────────────────
# office-code-interpreter — self-hosted LibreCodeInterpreter, with
# office-doctrine preinstalled in the sandbox runtime.
#
# Why a custom image:
#   The upstream `ghcr.io/usnavy13/librecodeinterpreter:main` already
#   ships python-docx, python-pptx, openpyxl AND a full LibreOffice
#   headless stack (writer + calc + impress). What it does NOT ship is
#   our `office-doctrine` lib (the bytes-in / bytes-out engines extracted
#   in PR1, packages/office-doctrine/). This image adds that wheel — and
#   nothing else — so LibreChat skills can `import office_doctrine` from
#   inside the sandbox without bundling the lib in every skill zip.
#
# Build context: MUST be the repo root, not deployment/docker/code-interpreter/.
# The first stage needs access to packages/office-doctrine/ to build
# the wheel. Build command:
#
#   docker build -f deployment/docker/code-interpreter/Dockerfile \
#                -t office-code-interpreter:test .
#
# (This is what the compose service does via
#  `context: ../.., dockerfile: deployment/docker/code-interpreter/Dockerfile`.)
# ─────────────────────────────────────────────────────────────────────

ARG PYTHON_VERSION=3.12

# ============================================
# Stage 1 — Build the office-doctrine wheel
# ============================================
# Pure Poetry build, no runtime deps installed. Output is a single .whl
# under /build/dist/, copied into the runtime stage below.
FROM python:${PYTHON_VERSION}-slim AS wheel-builder

ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    PIP_NO_CACHE_DIR=1 \
    POETRY_VERSION=1.8.4 \
    POETRY_NO_INTERACTION=1

RUN pip install "poetry==${POETRY_VERSION}"

WORKDIR /build
COPY packages/office-doctrine /build/

RUN poetry build --format wheel --no-ansi \
 && ls -la /build/dist/

# ============================================
# Stage 2 — Runtime: LibreCodeInterpreter + our wheel
# ============================================
# We deliberately pin to `:main` (stable channel) rather than a sha-…
# tag because upstream does not cut versioned releases yet — see
# https://github.com/usnavy13/LibreCodeInterpreter/releases (empty as
# of 2026-05-19). Worst case for the PoC; we revisit once they tag.
#
# python-docx, python-pptx, openpyxl + libreoffice are already in the
# base — see https://github.com/usnavy13/LibreCodeInterpreter/blob/main/Dockerfile
# (runtime-core stage installs python-core.txt, python-documents.txt
# and `libreoffice-impress libreoffice-writer libreoffice-calc`).
FROM ghcr.io/usnavy13/librecodeinterpreter:main AS runtime

# Upstream installs Python packages with --break-system-packages because
# Ubuntu 24.04 marks the system Python as externally-managed (PEP 668).
# We follow the same convention so office_doctrine is importable from
# the same site-packages the sandbox uses.
COPY --from=wheel-builder /build/dist/*.whl /wheels/

RUN pip install --break-system-packages --no-cache-dir /wheels/*.whl \
 && rm -rf /wheels /root/.cache/pip

# Smoke script for local validation: `docker run --rm IMAGE python /smoke.py`
COPY deployment/docker/code-interpreter/smoke.py /smoke.py

# Inherit ENTRYPOINT / CMD / EXPOSE / HEALTHCHECK from the upstream image
# (CMD ["python3", "-m", "src.main"], EXPOSE 8000). Do not redefine them
# here — that would silently fork from upstream if they ever change.
