#!/usr/bin/env bash
# ─────────────────────────────────────────────────────────────────────
# migrate-librechat-to-pixel.sh
#
# One-shot helper to install the LibreChat stack at
#   $HOME/librechat-stack/
# on the pixel host, in a way that survives an office-mcp redeploy.
#
# Why this script exists:
#   The office-mcp CI used to deploy LibreChat alongside office-mcp from
#   docker-compose.librechat.yml. With the repo going public the
#   LibreChat compose has been removed from git (it's not the MCP
#   server's concern). The orphaned librechat / mongo / meilisearch
#   containers — still running from the last full-stack deploy — must
#   be re-adopted into a stand-alone compose project so they (1) keep
#   their existing data volumes, (2) survive future office-mcp
#   deploys, (3) keep reaching office-mcp via the shared
#   `proxy-network`.
#
# Run as the operator user (e.g. yacine) via:
#   scp scripts/migrate-librechat-to-pixel.sh pixel:/tmp/
#   ssh pixel "bash /tmp/migrate-librechat-to-pixel.sh"
#
# Stack lands in $HOME/librechat-stack so the operator owns it (no
# sudo needed for upkeep). The CI workspace
# /home/marvin/actions-runner-omcp/_work/office-mcp/office-mcp/.env
# is only READ for secret extraction (world-readable).
#
# Idempotent: re-running it just refreshes the compose/yaml/.env and
# `docker compose up -d` again. Existing volumes are reused.
# ─────────────────────────────────────────────────────────────────────
set -euo pipefail

STACK_DIR="$HOME/librechat-stack"
CI_ENV="/home/marvin/actions-runner-omcp/_work/office-mcp/office-mcp/.env"

echo "==> Preparing $STACK_DIR"
mkdir -p "$STACK_DIR"
cd "$STACK_DIR"

# ── Source the secrets ───────────────────────────────────────────────
# OFFICE_MCP_TOKEN + AZURE_* come from the CI-generated .env (the
# office-mcp workflow continues to write them on every deploy).
#
# LibreChat-only secrets (CREDS_*, JWT_*, OPENID_SESSION_SECRET,
# OPENROUTER_API_KEY) live in the *currently running* `librechat`
# container's environment — they were baked in by the previous
# full-stack deploy and aren't in the post-split CI .env anymore.
# Reading them back from the container preserves the existing JWT
# signing key (no forced logout) and avoids re-generating an OPENID
# session secret that would invalidate active SSO cookies.

if [ ! -r "$CI_ENV" ]; then
  echo "ERROR: cannot read $CI_ENV — has the office-mcp CI ever run on this host?"
  exit 1
fi

if ! docker inspect librechat >/dev/null 2>&1; then
  echo "ERROR: no running librechat container to read the legacy secrets from."
  echo "       Either run this script while the orphan stack is still up, or"
  echo "       seed the env values manually."
  exit 1
fi

get_ci() { grep -E "^${1}=" "$CI_ENV" | head -1 | cut -d= -f2-; }
get_lc() { docker inspect librechat --format '{{range .Config.Env}}{{println .}}{{end}}' | grep -E "^${1}=" | head -1 | cut -d= -f2-; }

OFFICE_MCP_TOKEN=$(get_ci OFFICE_MCP_TOKEN)
AZURE_CLIENT_ID=$(get_ci AZURE_CLIENT_ID)
AZURE_CLIENT_SECRET=$(get_ci AZURE_CLIENT_SECRET)
AZURE_TENANT_ID=$(get_ci AZURE_TENANT_ID)

CREDS_KEY=$(get_lc CREDS_KEY)
CREDS_IV=$(get_lc CREDS_IV)
JWT_SECRET=$(get_lc JWT_SECRET)
JWT_REFRESH_SECRET=$(get_lc JWT_REFRESH_SECRET)
OPENID_SESSION_SECRET=$(get_lc OPENID_SESSION_SECRET)
OPENROUTER_API_KEY=$(get_lc OPENROUTER_API_KEY)

for var in OFFICE_MCP_TOKEN AZURE_CLIENT_ID AZURE_CLIENT_SECRET AZURE_TENANT_ID; do
  if [ -z "${!var}" ]; then
    echo "ERROR: $var missing from $CI_ENV — refusing to write incomplete config"
    exit 1
  fi
done

for var in CREDS_KEY CREDS_IV JWT_SECRET JWT_REFRESH_SECRET \
           OPENID_SESSION_SECRET OPENROUTER_API_KEY; do
  if [ -z "${!var}" ]; then
    echo "ERROR: $var not found in the running librechat container env"
    exit 1
  fi
done

echo "==> Writing .env"
cat > .env <<EOF
LIBRECHAT_VERSION=v0.8.5
LIBRECHAT_PORT=3080

# Secrets reused from the office-mcp CI-generated .env
OFFICE_MCP_TOKEN=$OFFICE_MCP_TOKEN
CREDS_KEY=$CREDS_KEY
CREDS_IV=$CREDS_IV
JWT_SECRET=$JWT_SECRET
JWT_REFRESH_SECRET=$JWT_REFRESH_SECRET

# OIDC SSO (Azure AD) — same app registration as office-mcp
ALLOW_REGISTRATION=false
ALLOW_SOCIAL_LOGIN=true
ALLOW_SOCIAL_REGISTRATION=true
OPENID_CLIENT_ID=$AZURE_CLIENT_ID
OPENID_CLIENT_SECRET=$AZURE_CLIENT_SECRET
OPENID_ISSUER=https://login.microsoftonline.com/$AZURE_TENANT_ID/v2.0
OPENID_SESSION_SECRET=$OPENID_SESSION_SECRET
OPENID_SCOPE=openid profile email offline_access Files.ReadWrite
OPENID_REUSE_TOKENS=true
OPENID_CALLBACK_URL=/oauth/openid/callback
OPENID_BUTTON_LABEL=Continuer avec Microsoft

# OpenRouter LLM
OPENROUTER_API_KEY=$OPENROUTER_API_KEY
EOF
chmod 600 .env

echo "==> Writing librechat.yaml"
cat > librechat.yaml <<'EOF'
version: 1.2.0
cache: true

endpoints:
  custom:
    - name: "OpenRouter"
      apiKey: "${OPENROUTER_API_KEY}"
      baseURL: "https://openrouter.ai/api/v1"
      models:
        default: ["nvidia/nemotron-3-super-120b-a12b:free"]
        fetch: true
      titleConvo: true
      titleModel: "nvidia/nemotron-3-super-120b-a12b:free"
      modelDisplayLabel: "OpenRouter"
      headers:
        HTTP-Referer: "https://office-mcp.reho.dev"
        X-Title: "office-mcp"

mcpSettings:
  allowedDomains:
    - "office-mcp"

mcpServers:
  office:
    type: streamable-http
    url: http://office-mcp:8080/mcp
    timeout: 120000
    initTimeout: 15000
    chatMenu: true
    title: "Office Documents"
    description: "Create and edit Word, Excel, PowerPoint documents."
    headers:
      Authorization: "Bearer ${OFFICE_MCP_TOKEN}"
      X-User-Authorization: "Bearer {{LIBRECHAT_OPENID_ACCESS_TOKEN}}"
      X-LibreChat-User-Id: "{{LIBRECHAT_USER_ID}}"
      X-LibreChat-User-Email: "{{LIBRECHAT_USER_EMAIL}}"
    serverInstructions: |
      This MCP creates and edits Office documents (.docx, .xlsx, .pptx) on the user's
      Microsoft OneDrive. Files are addressed by `od://` URIs; bare names like
      `od://q3.docx` are rewritten server-side to `od://AI Generated/q3.docx`.
      Build the whole document in a single create call (blocks=[...] / sheets=[...]
      / slides=[...]). Always echo the resulting file_uri back to the user.
EOF

echo "==> Writing docker-compose.yml"
cat > docker-compose.yml <<'EOF'
# LibreChat stack — maintained hors-gitops on pixel.
# Adopts the librechat / mongo / meili containers that were previously
# managed by the office-mcp CI. External volumes and networks preserve
# data and connectivity to office-mcp across redeploys.

services:
  mongodb:
    image: mongo:7.0
    container_name: office-mcp-mongodb
    restart: unless-stopped
    volumes:
      - mongodb-data:/data/db
    networks:
      - office-mcp
    command: mongod --noauth

  meilisearch:
    image: getmeili/meilisearch:v1.12.3
    container_name: office-mcp-meilisearch
    restart: unless-stopped
    environment:
      - MEILI_HOST=http://meilisearch:7700
      - MEILI_NO_ANALYTICS=true
    volumes:
      - meili-data:/meili_data
    networks:
      - office-mcp

  librechat:
    image: ghcr.io/danny-avila/librechat:${LIBRECHAT_VERSION:-v0.8.5}
    container_name: librechat
    restart: unless-stopped
    depends_on:
      - mongodb
      - meilisearch
    ports:
      - "${LIBRECHAT_PORT:-3080}:3080"
    environment:
      - HOST=0.0.0.0
      - MONGO_URI=mongodb://mongodb:27017/LibreChat
      - MEILI_HOST=http://meilisearch:7700
      - MEILI_NO_ANALYTICS=true
      - CREDS_KEY=${CREDS_KEY}
      - CREDS_IV=${CREDS_IV}
      - JWT_SECRET=${JWT_SECRET}
      - JWT_REFRESH_SECRET=${JWT_REFRESH_SECRET}
      - OFFICE_MCP_TOKEN=${OFFICE_MCP_TOKEN}
      - ALLOW_REGISTRATION=${ALLOW_REGISTRATION:-false}
      - OPENROUTER_API_KEY=${OPENROUTER_API_KEY}
      - ALLOW_SOCIAL_LOGIN=${ALLOW_SOCIAL_LOGIN:-true}
      - ALLOW_SOCIAL_REGISTRATION=${ALLOW_SOCIAL_REGISTRATION:-true}
      - OPENID_CLIENT_ID=${OPENID_CLIENT_ID}
      - OPENID_CLIENT_SECRET=${OPENID_CLIENT_SECRET}
      - OPENID_ISSUER=${OPENID_ISSUER}
      - OPENID_SESSION_SECRET=${OPENID_SESSION_SECRET}
      - OPENID_SCOPE=${OPENID_SCOPE}
      - OPENID_REUSE_TOKENS=${OPENID_REUSE_TOKENS:-true}
      - OPENID_CALLBACK_URL=${OPENID_CALLBACK_URL:-/oauth/openid/callback}
      - OPENID_BUTTON_LABEL=${OPENID_BUTTON_LABEL:-Continuer avec Microsoft}
      - DOMAIN_CLIENT=https://librechat.reho.dev
      - DOMAIN_SERVER=https://librechat.reho.dev
    volumes:
      - ./librechat.yaml:/app/librechat.yaml:ro
      - librechat-images:/app/client/public/images
      - librechat-uploads:/app/uploads
    networks:
      - office-mcp
      - proxy-network
    labels:
      - "traefik.enable=true"
      - "traefik.docker.network=proxy-network"
      - "traefik.http.routers.librechat.rule=Host(`librechat.reho.dev`)"
      - "traefik.http.routers.librechat.entrypoints=websecure"
      - "traefik.http.routers.librechat.tls.certresolver=letsencrypt"
      - "traefik.http.services.librechat.loadbalancer.server.port=3080"

# External volumes: reuse the data created by the previous office-mcp
# CI-managed full stack (project name was `office-mcp`).
volumes:
  mongodb-data:
    external: true
    name: office-mcp_mongodb-data
  meili-data:
    external: true
    name: office-mcp_meili-data
  librechat-images:
    external: true
    name: office-mcp_librechat-images
  librechat-uploads:
    external: true
    name: office-mcp_librechat-uploads

# External networks: `office-mcp` is created by the office-mcp stack
# CI deploy; `proxy-network` is created by the Traefik stack.
networks:
  office-mcp:
    external: true
  proxy-network:
    external: true
EOF

echo "==> Stopping orphaned containers (project=office-mcp)"
# These were created by the previous full-stack deploy and still carry
# the office-mcp compose project labels. Remove them so the new
# librechat-stack project can recreate them with its own labels,
# reusing the external volumes.
for c in librechat office-mcp-mongodb office-mcp-meilisearch; do
  if docker ps -a --format '{{.Names}}' | grep -qx "$c"; then
    echo "  stopping & removing $c"
    docker stop "$c" >/dev/null 2>&1 || true
    docker rm  "$c" >/dev/null 2>&1 || true
  fi
done

echo "==> Bringing up librechat-stack"
docker compose up -d

echo "==> Waiting for LibreChat to be healthy"
for i in $(seq 1 30); do
  if curl -sf "http://localhost:3080/api/config" >/dev/null; then
    echo "  LibreChat healthy (attempt $i)"
    break
  fi
  echo "  attempt $i/30..."
  sleep 10
  if [ $i -eq 30 ]; then
    echo "  ERROR: LibreChat did not become healthy in 5 minutes"
    docker compose logs librechat | tail -50
    exit 1
  fi
done

echo
echo "Done. Stack running under project=librechat-stack:"
docker compose ps
echo
echo "Verify externally: curl -sf https://librechat.reho.dev/api/config | head -c 200"
