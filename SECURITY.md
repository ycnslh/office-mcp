# Security Policy

## Reporting a vulnerability

If you find a security issue in office-mcp, please **do not open a public
issue**. Instead:

- **Preferred**: open a private report via
  [GitHub's "Report a vulnerability" flow](https://github.com/soulkill3r/office-mcp/security/advisories/new).
- **Alternative**: email the maintainer at the address listed in
  `pyproject.toml` (`[tool.poetry] authors`). Use a clear subject line
  starting with `[office-mcp security]`.

Please include enough detail to reproduce the issue (affected version
or commit SHA, request payload, expected vs observed behaviour). A
proof-of-concept is welcome but not required.

## Response expectations

This is a personal / small-team project, not a funded security
programme. Best-effort acknowledgement within **7 days** of receipt.
Fixes for confirmed issues land on `main` and are documented in the
release notes once available; coordinated disclosure timelines can be
agreed on a case-by-case basis.

## Scope

In scope:

- The MCP server source code in `src/office_mcp/` and the engines it
  embeds (docx / xlsx / pptx).
- The Dockerfile and the published image when running with documented
  configuration (`.env.example`).

Out of scope:

- Vulnerabilities in upstream dependencies — please report those to
  the respective project (we will track and update as fixes ship).
- Misconfiguration on a specific deployment (e.g. leaked
  `OFFICE_MCP_TOKEN`, exposed `/data` volume) — those are operational
  concerns, not server bugs.
- Denial-of-service via legitimate but expensive operations
  (`docx_to_pdf` on huge documents, `fs_list` on deep trees). The
  built-in rate limiter mitigates abuse; tighter quotas are a
  per-deployment choice.

## Hardening already in place

- Bearer auth on every MCP call (`OFFICE_MCP_TOKEN`).
- Per-user scoping enforced by the storage backends (`local://`
  prefixes paths by `user_id`; `od://` calls Graph as the end user via
  their delegated token).
- Path-traversal rejection on `local://` URIs.
- HMAC-signed download URLs with a short TTL (`OFFICE_MCP_SIGNING_SECRET`).
- Rate limiting per `user_id` (`OFFICE_MCP_RATE_LIMIT_PER_MIN`).
- File-size cap on writes (`OFFICE_MCP_MAX_FILE_SIZE_MB`).
