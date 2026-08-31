# IARD Render Service

Self-hosted document assembly for the Internal Audit Report Drafter Agent.
Replaces the fabricated `n8n-nodes-base.microsoftWord` node that the original
Document Assembly workflow design mistakenly referenced — that node type does
not exist in n8n, so this service is the actual mechanism.

## What it does

`POST /render-report` takes the current engagement state (front matter,
observations, appendix items) as JSON and returns a filled `.docx`, built by
opening Beyon's fixed report template and inserting content after the
matching section headings — never restyling or reordering the template
itself (GR-02).

Each observation's title paragraph gets a Word bookmark named `obs-<id>`, so
a later auditor-edited, re-uploaded Word file can eventually be diffed back
to the specific observation it came from (feeds the not-yet-built Round-trip
Reconciliation sub-workflow — the bookmark exists now, the diffing logic
that reads it does not yet).

Where an observation has no `cause` and isn't an Area for Improvement, the
service inserts a flagged placeholder rather than silently omitting it or
inventing one (GR-04, content rule 8).

## Setup

1. Place Beyon's actual (blank) report template at the path `TEMPLATE_PATH`
   points to (default `/etc/secrets/report_template.docx`, overridable via
   the `TEMPLATE_PATH` env var) — **not** committed to any repo, since it's
   Restricted-classified per BRD section 10. The copy this was tested
   against still has sample placeholder text baked into some sections; a
   genuinely blank version will render cleaner. For local dev, the simplest
   option is `export TEMPLATE_PATH=./report_template.docx` and drop the file
   next to `main.py`.
2. `pip install -r requirements.txt`
3. `uvicorn main:app --host 0.0.0.0 --port 8000` (or build/run the Dockerfile)
4. Confirm `GET /health` returns `{"status": "ok", "template_found": true}`.
   If `template_found` is `false`, the template hasn't landed at
   `TEMPLATE_PATH` yet — `/render-report` will fail with
   `"Package not found at '<path>'"` until it does.

## Deploying on Render (never via git)

The template is Restricted-classified and must never be committed, so it
reaches the deployed service through Render itself, not through the build:

1. Render dashboard → the `beyon-iard` service → **Environment** →
   **Secret Files** → **Add Secret File**.
2. Filename: `report_template.docx` (exact name — Render mounts it at
   `/etc/secrets/report_template.docx`, which is `main.py`'s default
   `TEMPLATE_PATH`).
3. Upload Beyon's actual template file as the contents.
4. Save — Render redeploys the service automatically. No `TEMPLATE_PATH`
   env var is needed if the filename above is used exactly as given; only
   set `TEMPLATE_PATH` explicitly if a persistent Disk is used instead of a
   Secret File (e.g. `TEMPLATE_PATH=/var/data/report_template.docx`, with
   the Disk mounted at `/var/data`).
5. Verify with `GET /health` — `template_found` should read `true`.

This keeps the Restricted file off GitHub entirely (public or private repo)
while still getting it into the running container.

## Known simplifications (be aware before treating this as final)

- Section-heading matching is done by scanning paragraph text for a known
  prefix (`"Background"`, `"Objectives:"`, etc.) — if Beyon's template
  wording changes, these anchors need updating.
- The Audit Team table fill assumes a specific 3-column table shape; if the
  real template's table differs, that block needs adjusting.
- Content is inserted *after* the matching heading rather than replacing
  existing sample text in the template — fine against a blank template, less
  clean against one that already has example content in those sections.
- No PDF export yet (BRD requires Word + PDF + PPTX at final issue) — this
  service only produces the Word document for now.
- No auth on the endpoint — needs an API key or network-level restriction
  before this touches real Restricted-classified engagement data.

## API contract

See `main.py`'s `RenderRequest` model. n8n's Document Assembly workflow
calls this via an HTTP Request node — see the corresponding fix in that
workflow for the exact call shape.
