# office-code-interpreter

Custom container image derived from
[LibreCodeInterpreter](https://github.com/usnavy13/LibreCodeInterpreter)
that adds the `office-doctrine` Python library to the sandbox runtime.

It powers the **Code Interpreter** button on `librechat-staging.reho.dev`
so the new self-contained LibreChat **Skills** (PR3) can produce `.docx`,
`.xlsx` and `.pptx` payloads through plain `import office_doctrine` from
inside the sandbox — no MCP, no OneDrive round-trip, native rich-artifact
previews instead.

## What is preinstalled

| Provided by | Packages |
| --- | --- |
| Upstream `librecodeinterpreter:main` | `python-docx`, `python-pptx`, `openpyxl`, `libreoffice-{writer,calc,impress}`, plus the full LibreCodeInterpreter polyglot toolchain (Node, Go, R, Java, Rust, …) |
| This image (delta) | `office_doctrine` wheel built from `packages/office-doctrine/` |

The upstream image is pinned to the floating `:main` tag — there are no
versioned releases yet ([releases page is empty](https://github.com/usnavy13/LibreCodeInterpreter/releases)).
Acceptable trade-off for the PoC; switch to an immutable `sha-…` tag once
upstream cuts releases or once we observe drift in staging.

## Build locally

The Dockerfile is multi-stage and the build context **must be the repo
root** so stage 1 can read `packages/office-doctrine/`.

```bash
# From the office-mcp repo root
docker build \
  -f deployment/docker/code-interpreter/Dockerfile \
  -t office-code-interpreter:test \
  .
```

This is exactly what the staging compose does via
`build: { context: ../.., dockerfile: deployment/docker/code-interpreter/Dockerfile }`
in [`../compose/docker-compose.librechat-staging.yml`](../compose/docker-compose.librechat-staging.yml).

## Smoke test

```bash
docker run --rm office-code-interpreter:test python /smoke.py
# → smoke ok (docx=NNNN bytes, libs=docx,openpyxl,pptx,office_doctrine)
```

## Updating the bundled office-doctrine

`office-doctrine` is bumped via PRs against `packages/office-doctrine/`.
When you change the lib:

1. Bump the `version` in `packages/office-doctrine/pyproject.toml`.
2. Rebuild this image. The multi-stage rebuilds the wheel from source on
   every build — no manual wheel copy needed.
3. Push to `main`; the CI workflow rebuilds and restarts
   `code-interpreter-staging` as part of the staging deploy job.

No code is bundled in skill zips — they `import office_doctrine` from the
sandbox site-packages, kept in sync by this image alone.

## Runtime dependencies (compose level)

LibreCodeInterpreter requires two sidecars that the staging override
also brings up: **Redis** (session state) and **Garage** (S3-compatible
object storage for sandbox files). Both run on the private
`office-mcp` network and are **not** exposed on the host or via Traefik.

## Why no port exposure / Traefik label

The container speaks plain HTTP on `8000` and is consumed only by
`librechat-staging` over the internal `office-mcp` Docker network. No
public ingress, no `traefik.enable=true` label. Staging users hit it
indirectly through LibreChat (`LIBRECHAT_CODE_BASEURL=http://code-interpreter-staging:8000`).
