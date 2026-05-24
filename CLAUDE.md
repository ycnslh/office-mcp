# office-mcp — Plan d'implémentation pour Claude Code

> MCP server pour créer et modifier des documents Word, Excel, PowerPoint.
> Transport : **streamable-http** sur `/mcp` (compatible n'importe quel client MCP).
> Stockage abstrait : **Local / SharePoint-OneDrive (Graph) / S3-Blob**.

---

## 0. Contexte et objectifs

Le client a challengé Yacine sur la réalisation d'un MCP server capable de créer/modifier des documents Office dans le cadre d'un pivot vers l'équipe IA. Le livrable doit :

- S'intégrer proprement à LibreChat 0.8.4 (transport **streamable-http**, recommandé en production par la doc LibreChat).
- Tourner en container Docker à côté de LibreChat, sur le même réseau interne.
- Abstraire le stockage des fichiers (local volume, SharePoint/OneDrive via Microsoft Graph, S3/Azure Blob).
- Exposer des outils métier clairs, pas une simple API CRUD.
- Embarquer une vraie "doctrine" de production (les règles Anthropic pour docx/xlsx/pptx — marges, polices, styles de tables, color coding finance pour xlsx, etc.) — exposées comme **MCP resources** pour que l'LLM les consulte à la demande.

### Critères de réussite (démo client)

1. Depuis LibreChat, un utilisateur demande "crée-moi un rapport Q3 avec un sommaire, un tableau de ventes et une conclusion" → un `.docx` propre est généré et accessible.
2. Même chose pour un .xlsx avec formules qui se calculent, et un .pptx avec slides cohérentes.
3. L'outil `office_read` permet de lire un document existant (Graph ou local) puis `*_edit` le modifie avec tracked-changes visibles.
4. Multi-user isolation : chaque utilisateur LibreChat voit uniquement ses fichiers (scoping par `LIBRECHAT_USER_ID`).
5. Zéro secret en clair dans le repo.

---

## 1. Stack technique

| Couche | Choix | Justification |
|---|---|---|
| Langage | Python 3.12 | Écosystème Office mature, cohérent avec ENGRAM |
| Framework MCP | **FastMCP 2.x** | Support natif streamable-http, déjà utilisé sur ENGRAM |
| Dep manager | **Poetry** | Cohérent avec le reste de la stack Yacine |
| Office libs | `python-docx`, `openpyxl`, `python-pptx` | Standards, lecture/écriture sans dépendance Office |
| Rendu/validation | `libreoffice --headless` | Recalcul formules xlsx, conversion PDF pour preview |
| Graph API | `msgraph-sdk` + `msal` | SDK officiel, OAuth on-behalf-of |
| S3/Blob | `boto3` / `azure-storage-blob` | Standards |
| Observabilité | `structlog` + OpenTelemetry | Même pattern que ENGRAM/PYTHIA |
| Tests | `pytest` + `pytest-asyncio` | |
| Container | Python 3.12-slim + LibreOffice headless | Image ~800MB, acceptable |

### Pourquoi pas le skill docx Anthropic en l'état

Le skill Anthropic utilise `docx-js` (Node) et un workflow unpack-edit-repack via scripts Python. C'est excellent pour un agent qui a accès à un shell, moins adapté à un MCP server exposant des outils. On va donc :

- **Réutiliser la connaissance du skill** (règles de mise en forme, pièges, patterns XML) et la matérialiser en :
  - **Ressources MCP** (`office://rules/docx`, `office://rules/xlsx`, etc.) servant les bonnes pratiques à l'LLM.
  - **Guardrails** dans le code (ex: refus d'insérer des bullets unicode, application automatique des doubles widths sur les tables, etc.).
- **Porter la stack en Python pur** (`python-docx` au lieu de `docx-js`) pour rester dans un seul runtime.

---

## 2. Architecture

```
                  ┌──────────────────────────────┐
                  │  LibreChat (api container)   │
                  │  librechat.yaml:             │
                  │    mcpServers.office:        │
                  │      type: streamable-http   │
                  │      url: http://office-mcp:8080/mcp │
                  └──────────────┬───────────────┘
                                 │ streamable-http + headers
                                 │   X-LibreChat-User-Id: {{LIBRECHAT_USER_ID}}
                                 │   X-LibreChat-User-Email: {{LIBRECHAT_USER_EMAIL}}
                  ┌──────────────▼───────────────┐
                  │  office-mcp (this repo)      │
                  │                              │
                  │  ┌────────────────────────┐  │
                  │  │ FastMCP server :8080   │  │
                  │  └──────────┬─────────────┘  │
                  │             │                │
                  │  ┌──────────▼─────────────┐  │
                  │  │ Tools layer            │  │
                  │  │  (docx/xlsx/pptx/fs)   │  │
                  │  └──────────┬─────────────┘  │
                  │             │                │
                  │  ┌──────────▼─────────────┐  │
                  │  │ Engines layer          │  │
                  │  │  python-docx/openpyxl  │  │
                  │  │  python-pptx           │  │
                  │  └──────────┬─────────────┘  │
                  │             │                │
                  │  ┌──────────▼─────────────┐  │
                  │  │ StorageBackend (ABC)   │  │
                  │  │  Local / Graph / S3    │  │
                  │  └────────────────────────┘  │
                  └──────────────┬───────────────┘
                                 │
              ┌──────────────────┼──────────────────┐
              ▼                  ▼                  ▼
        Docker volume      MS Graph API         S3 / Azure Blob
        local://           sp:// / od://        s3:// / az://
```

### URI scheme stockage

Les outils manipulent un `file_uri` abstrait. Le backend est résolu par le scheme :

| Scheme | Backend | Exemple |
|---|---|---|
| `local://` | LocalBackend (volume Docker) | `local://reports/q3-2026.docx` |
| `sp://` | SharePointBackend (Graph) | `sp://site-name/drive/Docs/report.docx` |
| `od://` | OneDriveBackend (Graph) | `od://me/Documents/pitch.pptx` |
| `s3://` | S3Backend | `s3://bucket/path/file.xlsx` |
| `az://` | AzureBlobBackend | `az://container/path/file.xlsx` |

Chaque backend implémente `StorageBackend` (voir §5).

---

## 3. Surface d'outils MCP

### Nommage et granularité

Convention : `<format>_<action>` (ex: `docx_create`). Les outils retournent toujours un dict structuré avec `file_uri`, `size_bytes`, et un `summary` textuel.

### Docx (6 outils, batched-by-intent)

Chaque tool est conçu pour faire **toute son intention en un seul appel**
afin que l'LLM ne chaîne plus de séquences `create + N×append + M×insert`.

| Tool | Description |
|---|---|
| `docx_create` | Crée un docx complet en un appel : titre + `blocks` (liste ordonnée de `heading`/`paragraph`/`table`) |
| `docx_read` | Extrait texte + structure (headings, tables) en JSON LLM-friendly |
| `docx_append` | Ajoute N blocs à la fin d'un doc existant en un appel |
| `docx_replace` | Applique N find/replace (literal ou regex) en une passe, retourne le compte |
| `docx_remove` | Supprime N sections (par texte de heading), retourne le compte |
| `docx_to_pdf` | Convertit en PDF via LibreOffice headless |

> Pas de `docx_delete` (côté MCP) : supprimer le fichier passe par `fs_delete`,
> qui couvre tous les backends storage de manière cohérente.

### Xlsx (3 outils, batched-by-intent)

| Tool | Description |
|---|---|
| `xlsx_create` | Crée un xlsx complet en un appel : `sheets=[{name, writes, formulas}]`, recalc LibreOffice unique à la fin |
| `xlsx_read_range` | Lit une plage (sheet + range A1:C10) |
| `xlsx_update` | Applique N writes + N formulas sur N sheets en un appel (recalc unique batché) |

### Pptx (4 outils, batched-by-intent)

| Tool | Description |
|---|---|
| `pptx_create` | Crée un deck complet en un appel : `theme` + `slides=[…]` (types : title/content/two-content/section) |
| `pptx_read` | Extrait texte + structure des slides |
| `pptx_append` | Ajoute N slides à la fin en un appel |
| `pptx_replace` | Patche N slides par index en un appel (title / bullets / left / right) |

### Filesystem abstrait (4 outils)

| Tool | Description |
|---|---|
| `fs_list` | Liste les fichiers pour un URI (ex: `sp://site/drive/folder/`) |
| `fs_stat` | Métadonnées d'un fichier (taille, mtime, etag) |
| `fs_copy` | Copie entre backends (ex: `local://` → `sp://`) |
| `fs_delete` | Supprime un fichier |

**Total : 17 outils** (docx 6 + xlsx 3 + pptx 4 + fs 4). La doctrine "batched-by-intent" — appliquée aux trois formats — réduit le surface area sans perdre d'usage : chaque create accepte le contenu complet en un seul appel, et les éditions batch N opérations en un round-trip.

### Ressources MCP (pas des tools, à servir par le server)

| Resource URI | Contenu |
|---|---|
| `office://rules/docx` | Règles de mise en forme Word (marges, polices, tables, bullets) |
| `office://rules/xlsx` | Conventions finance (color coding, formats, %, currencies) |
| `office://rules/pptx` | Design rules (palette, dominance, motif visuel) |
| `office://templates/list` | Liste des templates dispo dans le backend |

Ces ressources sont directement dérivées des skills Anthropic — on capitalise sur la doctrine sans la réinventer.

---

## 4. Structure du repo

```
office-mcp/
├── pyproject.toml              # Poetry
├── poetry.lock
├── README.md                   # quickstart + intégration LibreChat
├── LICENSE
├── .env.example
├── .gitignore
├── .dockerignore
├── Dockerfile
├── docker-compose.yml          # service standalone (dev)
├── docker-compose.librechat.override.yml.example  # à copier dans le projet LibreChat
├── librechat.yaml.example      # snippet à merger dans librechat.yaml
│
├── src/
│   └── office_mcp/
│       ├── __init__.py
│       ├── __main__.py         # entry point: python -m office_mcp
│       ├── server.py           # FastMCP app + lifecycle
│       ├── config.py           # pydantic-settings
│       ├── logging.py          # structlog setup
│       ├── context.py          # user context extraction (headers LibreChat)
│       │
│       ├── storage/
│       │   ├── __init__.py
│       │   ├── base.py         # StorageBackend ABC
│       │   ├── registry.py     # resolve scheme → backend
│       │   ├── local.py        # LocalBackend
│       │   ├── graph.py        # SharePoint + OneDrive via msgraph-sdk
│       │   ├── s3.py           # S3Backend (boto3)
│       │   └── azure_blob.py   # AzureBlobBackend
│       │
│       ├── engines/
│       │   ├── __init__.py
│       │   ├── docx_engine.py  # wraps python-docx, enforces rules
│       │   ├── xlsx_engine.py  # wraps openpyxl + libreoffice recalc
│       │   └── pptx_engine.py  # wraps python-pptx, themes
│       │
│       ├── tools/
│       │   ├── __init__.py
│       │   ├── docx_tools.py   # @mcp.tool decorators
│       │   ├── xlsx_tools.py
│       │   ├── pptx_tools.py
│       │   └── fs_tools.py
│       │
│       ├── resources/
│       │   ├── __init__.py
│       │   ├── rules.py        # @mcp.resource for rules URIs
│       │   └── content/
│       │       ├── docx_rules.md
│       │       ├── xlsx_rules.md
│       │       └── pptx_rules.md
│       │
│       ├── templates/
│       │   ├── docx/
│       │   │   ├── report.docx
│       │   │   └── memo.docx
│       │   └── pptx/
│       │       ├── cyberpunk.pptx
│       │       └── corporate.pptx
│       │
│       └── utils/
│           ├── __init__.py
│           ├── libreoffice.py  # headless invocation
│           ├── validation.py   # uri parsing, path safety
│           └── errors.py       # OfficeMcpError hierarchy
│
├── tests/
│   ├── conftest.py
│   ├── unit/
│   │   ├── test_storage_local.py
│   │   ├── test_storage_registry.py
│   │   ├── test_docx_engine.py
│   │   ├── test_xlsx_engine.py
│   │   └── test_pptx_engine.py
│   ├── integration/
│   │   ├── test_docx_tools.py
│   │   ├── test_xlsx_tools.py
│   │   └── test_pptx_tools.py
│   └── e2e/
│       └── test_mcp_client.py  # vrai client MCP sur streamable-http
│
└── docs/
    ├── architecture.md
    ├── storage-backends.md
    ├── librechat-integration.md
    └── security.md
```

---

## 5. Interfaces clés (contrats)

### StorageBackend ABC

```python
# src/office_mcp/storage/base.py
from abc import ABC, abstractmethod
from dataclasses import dataclass
from typing import AsyncIterator

@dataclass
class FileMetadata:
    uri: str
    size_bytes: int
    modified_at: str  # ISO8601
    etag: str | None
    mime_type: str | None

class StorageBackend(ABC):
    scheme: str  # "local", "sp", "od", "s3", "az"

    @abstractmethod
    async def read(self, uri: str, *, user_ctx: "UserContext") -> bytes: ...

    @abstractmethod
    async def write(self, uri: str, data: bytes, *, user_ctx: "UserContext",
                    if_match: str | None = None) -> FileMetadata: ...

    @abstractmethod
    async def stat(self, uri: str, *, user_ctx: "UserContext") -> FileMetadata: ...

    @abstractmethod
    async def list(self, uri_prefix: str, *, user_ctx: "UserContext") -> AsyncIterator[FileMetadata]: ...

    @abstractmethod
    async def delete(self, uri: str, *, user_ctx: "UserContext") -> None: ...
```

Le paramètre `if_match` sur `write` permet l'optimistic concurrency (ETag sur Graph, version id sur S3) — crucial pour éviter les écrasements concurrents en multi-user.

### UserContext

```python
# src/office_mcp/context.py
from dataclasses import dataclass

@dataclass(frozen=True)
class UserContext:
    user_id: str          # LibreChat user ID
    email: str | None
    on_behalf_of_token: str | None  # pour Graph API si OAuth user
```

Extrait des headers HTTP injectés par LibreChat via `{{LIBRECHAT_USER_ID}}` etc.

### Registry de backends

```python
# src/office_mcp/storage/registry.py
def get_backend(uri: str) -> StorageBackend:
    scheme = uri.split("://", 1)[0]
    backend = _REGISTRY.get(scheme)
    if not backend:
        raise UnknownSchemeError(scheme)
    return backend
```

---

## 6. Sécurité

### À appliquer dès la V1

1. **Path traversal** : valider que les URI `local://` résolvent sous le volume mappé, rejeter `..`.
2. **URL SSRF** : refuser `file://`, `http(s)://` arbitraires en entrée si jamais un tool accepte une URL source.
3. **Auth MCP** : ajouter un header `Authorization: Bearer <token>` optionnel côté server, token partagé avec LibreChat via `.env`. LibreChat le transmet via le champ `headers` de `mcpServers`.
4. **Scoping user** : le `LocalBackend` préfixe tous les chemins par `/data/{user_id}/`. Aucun tool ne peut accéder aux données d'un autre user.
5. **Secrets** : pas de clé Graph/S3 dans le code. Tout via env vars, documentées dans `.env.example`.
6. **Allowed domains LibreChat** : bien documenter l'ajout de `office-mcp` dans `mcpSettings.allowedDomains` du `librechat.yaml`.
7. **Taille max fichiers** : 50 MB par défaut (configurable), 413 au-delà.
8. **Rate limiting** : middleware simple token-bucket par `user_id`, 60 req/min par défaut.

### À prévoir en V2

- Graph API en mode "on-behalf-of" (reuse du token OAuth LibreChat) plutôt qu'app permissions.
- Audit log structuré (qui a créé/modifié quoi) vers stdout JSON, exploitable par Loki.
- Signature des fichiers générés (hash SHA-256 retourné dans la réponse).

---

## 7. Intégration côté client MCP

office-mcp parle **streamable-http** sur `/mcp`. Tout client MCP (LibreChat, Claude Desktop, etc.) peut s'y brancher. La configuration côté client (yaml LibreChat, json MCP, etc.) **ne vit plus dans ce repo** — elle est maintenue séparément par chaque utilisateur du serveur.

Pour l'instance pixel de Yacine, la config LibreChat vit hors-gitops sous `~yacine/librechat-stack/` (voir `.claude/agents/pixel-deployer.md` pour le détail).

Le seul contrat exposé par office-mcp :

| Header HTTP | Rôle |
|---|---|
| `Authorization: Bearer <token>` | auth server-level entre le client MCP et office-mcp (partagé via `OFFICE_MCP_TOKEN`) |
| `X-User-Authorization: Bearer <graph-token>` | token Graph délégué de l'utilisateur final, aud=Graph + scope Files.ReadWrite, utilisé tel quel par office-mcp pour appeler `/me/drive/*` |
| `X-User-Id`, `X-User-Email` | identité du user pour le scoping multi-tenant et l'audit log |

---

## 8. Phases de delivery (pour Claude Code)

Chaque phase finit sur un commit propre avec Conventional Commits, tests verts, et une note de démo.

### Phase 0 — Bootstrap (30 min)

- `poetry init`, dépendances : `mcp[cli]`, `fastmcp`, `python-docx`, `openpyxl`, `python-pptx`, `pydantic-settings`, `structlog`, `anyio`.
- Dependencies dev : `pytest`, `pytest-asyncio`, `ruff`, `mypy`.
- Squelette de repo (dossiers vides + `__init__.py`).
- `Dockerfile` multi-stage + `docker-compose.yml` minimal (service unique + volume local).
- `.env.example`, `.gitignore`, `.dockerignore`, `README.md` (intro + quickstart).
- Commit : `chore: bootstrap project with poetry and docker`.

### Phase 1 — Storage abstraction + LocalBackend (1h)

- `StorageBackend` ABC, `FileMetadata`, `UserContext`.
- `LocalBackend` complet avec scoping `/data/{user_id}/` et protection path traversal.
- Registry + resolver par scheme.
- Tests unitaires `test_storage_local.py` (read/write/stat/list/delete + path traversal rejection + user scoping).
- Commit : `feat(storage): add StorageBackend ABC with LocalBackend implementation`.

### Phase 2 — FastMCP server + context extraction (45 min)

- `server.py` : app FastMCP, transport streamable-http, endpoint `/health`.
- Middleware d'extraction `UserContext` depuis headers HTTP (`X-LibreChat-User-Id`, etc.).
- Auth bearer token optionnelle.
- Endpoint `/mcp` exposé.
- Test e2e basique : un client MCP Python se connecte, liste les tools (vide pour l'instant).
- Commit : `feat(server): FastMCP streamable-http server with user context middleware`.

### Phase 3 — Docx engine + 4 tools de base (2h)

- `docx_engine.py` : wrapper `python-docx` avec application automatique des règles (police Arial, marges 1", heading styles, tables en DXA double width, etc.).
- Tools : `docx_create` (avec `blocks` batché), `docx_read`, `docx_append`, `docx_replace`, `docx_remove`.
- Ressource MCP `office://rules/docx` qui sert `docx_rules.md` (reprise adaptée du skill Anthropic).
- Tests : create → read → assert contenu, insert table → assert double widths présentes.
- Commit : `feat(docx): add docx engine and 4 core tools with formatting rules`.

### Phase 4 — Xlsx engine + 4 tools (2h)

- `xlsx_engine.py` : wrapper `openpyxl`, invocation LibreOffice pour recalc des formules, color coding finance auto-applicable.
- Tools : `xlsx_create` (avec `sheets=[{name, writes, formulas}]` batché), `xlsx_read_range`, `xlsx_update`.
- Ressource `office://rules/xlsx`.
- Tests : create avec sheets, write range, apply formula avec recalc vérifié.
- Commit : `feat(xlsx): add xlsx engine with 4 tools and formula recalc`.

### Phase 5 — Pptx engine + 3 tools (1h30)

- `pptx_engine.py` : wrapper `python-pptx` avec thèmes (cyberpunk, corporate, minimal).
- Templates `.pptx` dans `src/office_mcp/templates/pptx/`.
- Tools : `pptx_create` (avec `slides=[…]` batché), `pptx_read`, `pptx_append`, `pptx_replace`.
- Ressource `office://rules/pptx`.
- Tests.
- Commit : `feat(pptx): add pptx engine with 3 tools and themes`.

### Phase 6 — Filesystem tools + docx_to_pdf (45 min)

- Tools : `fs_list`, `fs_stat`, `fs_copy`, `fs_delete`, `docx_to_pdf`.
- Tests (copy local → local, copy cross-backend mocké).
- Commit : `feat(fs): add filesystem tools and pdf export`.

### Phase 7 — Graph backend (SharePoint + OneDrive) (2h)

- `graph.py` : implémente `StorageBackend` via `msgraph-sdk`.
- Auth : app permissions en V1 (client credentials), on-behalf-of en V2.
- ETag support pour `write(if_match=...)`.
- Tests d'intégration avec mock Graph (via `pytest-httpx`).
- Commit : `feat(storage): add Graph backend for SharePoint and OneDrive`.

### Phase 8 — S3 + Azure Blob backends (1h)

- `s3.py` et `azure_blob.py` (implémentations similaires, simples).
- Tests avec `moto` (S3) et mocks Azure.
- Commit : `feat(storage): add S3 and Azure Blob backends`.

### Phase 9 — Intégration LibreChat + démo scripts (1h30)

- `librechat.yaml.example` finalisé.
- `docker-compose.librechat.override.yml.example` finalisé.
- Script `scripts/demo_scenarios.md` : 3 scénarios clé-en-main à dérouler pendant la démo (rapport docx, modèle xlsx de budget, deck pptx projet).
- README enrichi avec section "Intégration LibreChat" step-by-step.
- Commit : `docs: add LibreChat integration guide and demo scenarios`.

### Phase 10 — Hardening (1h)

- Rate limiting par user (in-memory token bucket).
- Logging structuré JSON + correlation IDs.
- Healthcheck enrichi (check storage backends).
- `/metrics` endpoint Prometheus optionnel.
- Commit : `feat(ops): add rate limiting, structured logging and metrics`.

### Phase 11 — CI/CD (30 min)

- `.github/workflows/ci.yml` : lint (ruff), type check (mypy), tests, build Docker.
- Image publiée sur ghcr.io ou registry privé.
- Commit : `ci: add GitHub Actions for lint, tests and Docker build`.

**Estimation totale : ~14 heures de dev pour un POC solide, 2 jours pleins.**

---

## 9. Prompts à donner à Claude Code

### Prompt initial de bootstrap

```
Tu vas implémenter office-mcp, un serveur MCP (Model Context Protocol) en Python
qui permet de créer et modifier des documents Word, Excel et PowerPoint, intégrable
à LibreChat 0.8.4 via le transport streamable-http.

Lis d'abord le fichier PLAN_OFFICE_MCP.md à la racine du projet. Il contient le plan
complet, l'architecture, les contrats d'interface, la structure du repo, et les phases
de delivery.

Commence par la Phase 0 (Bootstrap). Crée le squelette du repo, configure Poetry,
le Dockerfile et le docker-compose.yml minimal. Fais un premier commit en
Conventional Commits. Ne commence PAS la Phase 1 tant que la Phase 0 n'est pas validée.

Règles de travail :
- Conventional Commits (feat:, fix:, chore:, docs:, test:, ci:).
- Un commit par phase, tests verts avant commit.
- Ruff + mypy en strict mode.
- Type hints complets sur tout le code.
- docstrings Google style sur les fonctions publiques.
- Pas de secret en dur, tout via pydantic-settings + .env.
- Tests pytest async (pytest-asyncio) en mode strict.

Une fois la Phase 0 terminée, arrête-toi et montre-moi le `tree` du repo
pour validation.
```

### Prompts de passage de phase

> "Phase N terminée et validée. Passe à la Phase N+1 en suivant PLAN_OFFICE_MCP.md §8. Rappel des règles : tests verts + Conventional Commits + stop à la fin pour revue."

---

## 10. Points d'attention pour Yacine

### Pendant la démo client

- **Montrer l'extensibilité** : créer un 4ᵉ backend storage en live (ex: Google Drive) pour prouver que l'abstraction tient la route. C'est ~50 lignes de code.
- **Montrer les resources MCP** : demander à l'LLM de lire `office://rules/xlsx` avant de générer un modèle financier, le résultat est visiblement meilleur.
- **Multi-user** : deux sessions client avec deux users différents, montrer qu'ils ne voient pas les fichiers de l'autre.
- **Audit log** : montrer les logs JSON structurés avec le user_id, c'est un gros plus en contexte entreprise.

### Questions probables du client

| Question | Réponse |
|---|---|
| Pourquoi FastMCP et pas SDK officiel MCP ? | Streamable-http natif, moins de boilerplate, communauté active. Le SDK officiel marche aussi, c'est interchangeable. |
| Pourquoi pas Office.js / Graph natif ? | Parce qu'on veut un MCP réutilisable hors Microsoft 365 aussi. L'abstraction storage permet de switcher. |
| Gestion des conflits simultanés ? | ETag / `if_match` sur write, 412 Precondition Failed renvoyé à l'LLM qui peut re-read et re-apply. |
| Comment on teste ? | Pyramide complète : unit sur engines, intégration sur tools, e2e avec client MCP réel. Voir `/tests`. |
| Quid de la sécurité SharePoint ? | V1 = app permissions (service principal scoped). V2 = on-behalf-of (chaque user utilise son propre token). |

### Pièges à éviter

1. **Ne pas porter trop fidèlement le workflow unpack/edit/pack** du skill Anthropic. En Python avec `python-docx` on accède directement à l'objet. L'unpack/pack est utile seulement pour les cas XML complexes (tracked changes profondes).
2. **Ne pas exposer trop d'outils**. 24 c'est déjà beaucoup. Si l'LLM s'emmêle, regrouper.
3. **Ne pas mettre LibreOffice dans le Dockerfile trop tôt** — tester d'abord sans (les formules ne seront pas recalculées mais ça marche pour le reste). Image ~800 MB avec vs ~200 MB sans.
4. **Côté client MCP : whitelister le domaine** — beaucoup de clients (dont LibreChat) bloquent par défaut en SSRF les URLs internes. Si tu pointes sur `http://office-mcp:8080`, ajoute-le à la liste autorisée du client.
5. **Headers user** : bien vérifier que le client substitue `{{user_id}}` ou équivalent avant l'envoi. Test simple : logger le header reçu et l'afficher dans les logs MCP au premier appel.

---

## 11. Livrables finaux du POC

Pour la démo au client, Yacine aura :

- ✅ Repo Git `office-mcp` public avec les phases commitées proprement.
- ✅ Image Docker `office-mcp:latest` prête à déployer.
- ✅ 3 scénarios de démo reproductibles (rapport / budget / deck).
- ✅ Doc d'architecture (`docs/architecture.md`) avec schéma.
- ✅ Tests verts (~80% couverture).
- ✅ CI/CD GitHub Actions.

Et surtout, un **article Dernière Goutte** "Building a production-grade Office MCP" à publier une fois le projet validé — parfait pour positionner Yacine sur le pivot IA.

---

*Fin du plan. À passer à Claude Code comme document de référence (lu en premier à chaque session).*