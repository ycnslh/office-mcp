# office-mcp — three demo scenarios

Three ready-to-paste prompts for the client demo. Each exercises a
distinct format and shows the end-to-end loop (tool call → file
created → clickable download in the conversation).

> **Tip.** Open two LibreChat sessions side by side — one with a
> LibreChat agent using office-mcp, one without — so the client sees
> the value the integration adds.

---

## 1. Word — Q3 revenue report

**Prompt:**

> Crée un rapport trimestriel dans `od://AI Generated/q3-report.docx`
> intitulé "Rapport Q3 2026". Consulte d'abord
> `office://rules/docx` pour les conventions de mise en forme. Ajoute :
> - une section "Résumé exécutif" avec deux paragraphes ;
> - une section "Chiffres clés" contenant un tableau Région / Revenu /
>   Croissance QoQ avec 3 lignes (EU, US, APAC) ;
> - une section "Perspectives" avec un paragraphe.
>
> Donne-moi le lien de téléchargement pour le docx et pour une
> version PDF.

**What the LLM should do:**

1. `office://rules/docx` resource read.
2. **One** `docx_create` call with `title="Rapport Q3 2026"` and a
   `blocks` list covering the three sections + the Région/Revenu/QoQ
   table — the entire document body in a single round-trip.
3. `docx_to_pdf("od://AI Generated/q3-report.docx")`.
4. Reply with two Markdown links (docx + pdf).

**Why it lands:**

- Shows formatting doctrine enforcement (table gets DXA-pinned widths,
  Arial font, proper heading styles — not free-form LLM text).
- Shows LibreOffice recalc-equivalent (PDF conversion) running inside
  the MCP.
- Shows the signed download URL flow: the user clicks, the file opens.

---

## 2. Excel — 2026 mini budget

**Prompt:**

> Construis un mini budget dans `od://AI Generated/budget-2026.xlsx`.
> Consulte `office://rules/xlsx`. Crée 3 feuilles (`Inputs`, `Calc`,
> `Output`).
>
> Dans `Inputs`, écris ces données (colonnes Région / Unités / Prix) :
> - EU  : 1200, 49,90
> - US  :  900, 52,00
> - APAC:  420, 47,50
>
> Dans `Calc`, calcule le revenu par région (Unités × Prix) et le
> total. Utilise des formules, pas des valeurs en dur.
>
> Dans `Output`, mets en valeur le total annuel avec une référence
> cross-sheet. Donne-moi le lien de téléchargement.

**What the LLM should do:**

1. `office://rules/xlsx` resource read — key takeaway: blue inputs /
   black formulas / green cross-sheet links, recalc required.
2. **One** `xlsx_create` call with `sheets=[…]` carrying the three
   sheets, the EU/US/APAC inputs (`style="input"`), the per-region
   revenue formulas, and the cross-sheet link in `Output`. A single
   LibreOffice recalc at the end bakes every formula in one pass.

**Why it lands:**

- Color-coded finance workbook out of the box — looks instantly
  like a real analyst built it.
- Formulas are actually computed by LibreOffice (shows a 1.8 s recalc
  in the logs) — not just formula strings.
- Three-sheet structure shows the MCP honours the "one topic per
  sheet" doctrine.

---

## 3. PowerPoint — roadmap 2026

**Prompt:**

> Génère un deck dans `od://AI Generated/roadmap-2026.pptx` avec le thème
> `cyberpunk`. Consulte `office://rules/pptx`. Structure :
>
> - Slide de titre : "Roadmap 2026 — office-mcp", sous-titre "Netwatch
>   — équipe IA".
> - Section "Déjà livré" — layout section.
> - Slide "Ce qui est en prod" — layout content — bullets : stockage
>   local scope par user, 16 outils MCP, 3 doctrines de mise en forme,
>   export PDF via LibreOffice headless.
> - Section "À venir" — layout section.
> - Slide "Roadmap" — layout two-content — gauche (Q1/Q2) : SharePoint
>   backend, rate limiting, logs structurés ; droite (Q3/Q4) : S3 &
>   Azure Blob, templates curés, image signée GHCR.
> - Slide "Questions ?" — layout section.
>
> Donne-moi le lien de téléchargement.

**What the LLM should do:**

1. `office://rules/pptx` resource read.
2. **One** `pptx_create` call with `theme="cyberpunk"` and a `slides=[…]`
   list covering the title, the two section dividers, the content
   slide, and the two-content slide — the entire deck in a single
   round-trip. Each slide has its `type` (`title` / `content` /
   `two-content` / `section`).
3. Reply with the download URL.

**Why it lands:**

- Themes are visible: cyberpunk = dark background, cyan title,
  magenta accents. Client sees that "theme" is a real knob, not
  cosmetic flavor.
- One tool per slide flavour (title / section / content /
  two-content). The LLM picks by tool name, never by enum — works
  even on weaker models that struggle with optional array params.

---

## Reset between scenarios

If you need to clean slate between demos, ask the user to delete the
files in their OneDrive `AI Generated/` folder, or have the LLM call
`fs_delete` on each demo URI.

## Bailout prompts

If a scenario gets stuck (LLM loops, tool returns an error it can't
recover from), use these short prompts to force a retry:

- "Oublie l'état précédent. Reprends depuis zéro sur le même
  `file_uri` : appelle `fs_delete` puis `*_create` à nouveau."
- "Montre-moi le lien de téléchargement uniquement, sous forme de
  lien Markdown cliquable."
