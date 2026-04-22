# office-mcp

MCP server to create and edit Word, Excel, and PowerPoint documents.
Designed for **LibreChat 0.8.4** integration over the `streamable-http` transport,
with a pluggable storage backend (Local volume, SharePoint/OneDrive via Microsoft
Graph, S3, Azure Blob).

> See [CLAUDE.md](./CLAUDE.md) for the full implementation plan and rationale.

## Status

Phase 0 — bootstrap. Project skeleton, Poetry config, Docker image, and a dev
docker-compose stack including LibreChat v0.8.4 are in place. No tools are
implemented yet.

## Quickstart (dev)

```bash
# 1. Clone and configure
cp .env.example .env
# Generate the four LibreChat secrets (each with `openssl rand -hex 32`)
# and paste into .env: CREDS_KEY, CREDS_IV, JWT_SECRET, JWT_REFRESH_SECRET.

# 2. Standalone office-mcp (no LibreChat)
docker compose up --build

# 3. Full dev stack with LibreChat 0.8.4
docker compose -f docker-compose.yml -f docker-compose.librechat.yml up --build
#   → office-mcp  : http://localhost:8080
#   → LibreChat   : http://localhost:3080  (register a user, then chat)
```

## Local development without Docker

```bash
# Requires Python 3.12 and Poetry
poetry install
poetry run office-mcp                     # serves on :8080
poetry run pytest                         # runs the test suite
poetry run ruff check . && poetry run mypy src
```

## Repo layout

```
src/office_mcp/
  server.py          FastMCP app + lifecycle
  config.py          pydantic-settings
  context.py         per-request UserContext (extracted from LibreChat headers)
  storage/           StorageBackend ABC + Local / Graph / S3 / Azure Blob
  engines/           python-docx, openpyxl, python-pptx wrappers
  tools/             @mcp.tool decorators (docx / xlsx / pptx / fs)
  resources/         @mcp.resource (rules, templates listing)
  templates/         starter .docx / .pptx templates
  utils/             libreoffice headless, validation, errors
tests/               unit / integration / e2e
docs/                architecture, storage backends, LibreChat integration
```

## LibreChat integration

A working example is provided as
[`librechat.yaml.example`](./librechat.yaml.example). Mount it into the
LibreChat container as `/app/librechat.yaml` (the `docker-compose.librechat.yml`
file does this automatically for the dev stack).

Required pieces:

* **`mcpSettings.allowedDomains`** must include `office-mcp` — without this,
  LibreChat refuses internal addresses as a SSRF guard.
* The shared bearer token (`OFFICE_MCP_TOKEN` env var) is forwarded via the
  `Authorization` header.
* Per-user isolation relies on the `X-LibreChat-User-Id` header, which the
  `streamable-http` transport injects via the `{{LIBRECHAT_USER_ID}}`
  placeholder.

## License

MIT — see [LICENSE](./LICENSE).
