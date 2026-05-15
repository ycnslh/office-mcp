# Integrating office-mcp with LibreChat

This guide walks through plugging `office-mcp` into an existing
LibreChat deployment (tested against v0.8.4). The end result: a user
chats with OpenRouter (or any other LLM LibreChat supports), asks for an
Office document, the LLM invokes the MCP tools, and a clickable
download link appears in the conversation.

## Prerequisites

- LibreChat v0.8.4 running in Docker (self-host or an existing
  project).
- `docker` and `docker compose` on the host.
- A secret value generated with `openssl rand -hex 32`
  (used for both the bearer token and the signing secret).
- Optionally, an OpenRouter API key — any LibreChat-supported provider works
  once the MCP is wired.

## 1. Build (or pull) the office-mcp image

From this repo:

```bash
docker build -t office-mcp:latest .
```

Or pull it from your internal registry if you publish it.

## 2. Drop-in override for the LibreChat stack

Copy [`docker-compose.librechat.override.yml.example`](../docker-compose.librechat.override.yml.example)
into the directory that hosts the client's LibreChat
`docker-compose.yml` and rename it `docker-compose.override.yml`.

Set three environment variables before `docker compose up`:

| Variable                   | What it is                                                       |
|----------------------------|------------------------------------------------------------------|
| `OFFICE_MCP_TOKEN`         | Bearer shared between LibreChat and the MCP, and the signing key |
| `OFFICE_MCP_PUBLIC_URL`    | Base URL a user's **browser** can reach — used in download links |
| `OPENROUTER_API_KEY`       | Only if you wire OpenRouter in LibreChat                         |

`OFFICE_MCP_PUBLIC_URL` example:

- `http://pixel.lan:18080` — LAN host
- `https://office.example.com` — behind a reverse proxy with TLS

If your LibreChat `docker-compose.yml` uses a network name other than
`librechat_default`, edit the `networks:` section at the bottom of the
override to match.

## 3. Merge the librechat.yaml snippet

In the LibreChat project's `librechat.yaml`, append the blocks from
[`librechat.yaml.example`](../librechat.yaml.example):

- `endpoints.custom` — adds OpenRouter as an OpenAI-compatible endpoint.
  Omit if OpenRouter isn't needed.
- `mcpSettings.allowedDomains` — **mandatory**: add `office-mcp` to the
  allow-list, otherwise LibreChat refuses internal addresses as SSRF.
- `mcpServers.office` — the streamable-http MCP declaration.

After editing, restart LibreChat: `docker compose restart api`.

If LibreChat runs with Redis caching, flush the config cache first:
`docker exec <redis-container> redis-cli FLUSHALL`.

## 4. Verify

Hit `/health` on the office-mcp container from the host:

```bash
curl http://localhost:18080/health
# → {"status":"ok","version":"0.1.0","phase":"9-librechat-integration","mcp_ready":true}
```

In LibreChat's UI, open a new chat, click the MCP menu. You should see
**Office Documents** listed as connected (a small green dot). The LLM
now has access to the 17 tools and 3 rules resources.

## 5. Known LibreChat 0.8.4 limitations

- **Header placeholders are not substituted in chat mode.** When the
  LLM calls a tool from a normal chat, LibreChat forwards the literal
  string `{{LIBRECHAT_USER_ID}}` instead of the user's ID (see
  [#9910](https://github.com/danny-avila/LibreChat/issues/9910)).
  The MCP server detects this and falls back to the
  `OFFICE_MCP_DEFAULT_USER_ID` scope (`demo-user` by default) so the
  demo still works. Per-user isolation kicks in automatically once
  LibreChat ships the fix — or if the MCP is called from
  LibreChat's **agents** flow, where substitution already works.
- **MCP tools cannot return binary files natively.** That's why this
  MCP exposes a signed `/files/{token}` endpoint and returns a
  `download_url` in every tool response: the LLM renders it as a
  Markdown link and the user downloads the file straight from the
  conversation.

## 6. Troubleshooting

- **`[MCP][office] Error POSTing to endpoint: Not Found`** in LibreChat
  logs — the office-mcp container isn't running, or `librechat.yaml`
  points at the wrong host/port. Inside the docker network the URL is
  `http://office-mcp:8080/mcp`.
- **`UnknownSchemeError: 's3'`** in office-mcp logs — the LLM invented
  an unsupported URI scheme. The server's error message now lists the
  available schemes and the LLM self-corrects on retry; if it persists,
  restate the constraint in the chat.
- **Download link returns 403** — the token expired (15 min default).
  Call the tool again to get a fresh link, or raise
  `OFFICE_MCP_DOWNLOAD_URL_TTL_SECONDS`.
- **Download link returns 404 with "not found"** — the file was in a
  different user scope (common when LibreChat's header substitution
  flips between real user ID and `demo-user`). Regenerate the file in
  the current session and retry.

## 7. What's still on the roadmap

- SharePoint / OneDrive backend (Phase 7) and S3 / Azure Blob
  (Phase 8) — waiting on the client's cloud choice.
- Structured JSON logging + Prometheus `/metrics` (Phase 10).
- CI-published image on GHCR for zero-build installs (Phase 11).
