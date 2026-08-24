# Internal Audit Report Drafter Agent — Build Package v0.1

This is the first build increment, per the sequencing agreed: state store schema, then
Intake & Parse, Front-Matter Assembly, and Document Assembly for a no-chat end-to-end
first draft, plus the Refinement chat sub-workflow designed alongside it since its
contract (state store read/write boundaries) shapes the other workflows.

## Contents

- `sql/001_schema.sql` — Postgres state store. Run this first. Includes a DB-level
  trigger enforcing GR-01 (frozen observations cannot be silently overwritten) so that
  guardrail doesn't depend solely on workflow logic behaving correctly.
- `workflows/01_orchestrator.json` — thin dispatcher: Form Trigger → new engagement
  path; Chat Trigger → resolve session to existing engagement → Refinement. No new-
  engagement path via chat, per your confirmation.
- `workflows/02_intake_and_parse.json` — classifies uploads, validates mandatory
  inputs present, records gaps for optional ones (never silent), fans out to per-file-
  type parsers (working file, kick-off deck, optional inputs incl. legacy .ppt).
- `workflows/03_front_matter_assembly.json` — verbatim / fill / opinion-lookup /
  generated-background sections. Opinion lookup currently points at the provisional
  `overall_opinion_library.json` I produced — swap for Beyon's confirmed version under
  D-2 without touching workflow logic.
- `workflows/04_document_assembly.json` — deterministic render from current state
  store. Tags each rendered observation with a stable id for round-trip mapping later.
  Includes a template-fidelity validation step (GR-02) before handoff to the reviewer.
- `workflows/05_refinement.json` — the chat-triggered sub-workflow, scoped strictly to
  Detailed Audit Issues as agreed. Classifies each instruction into one of eight
  actions, loads only the referenced observation(s) (not the whole report), blocks
  edits to frozen observations at the workflow layer (backed by the DB trigger),
  and writes a change-log entry on every mutation.

## Import order

1. Create the Postgres database and run `001_schema.sql`.
2. Import all five workflow JSON files into n8n.
3. Wire credentials: Postgres connection, LLM provider (used in Front-Matter
   Assembly's background generation and Refinement's instruction classification),
   and the Word-template rendering service/credential in Document Assembly.
4. Set `STATE_API_URL` and `REPORT_TEMPLATE_URL` environment variables referenced in
   node parameters.

## Deliberately not yet built (next increments)

- `Observation Drafting` sub-workflow (referenced by the orchestrator, not yet
  authored) — drafts each reportable working-file row into an observation.
- `Round-trip Reconciliation` — diffs an auditor-edited re-upload against the state
  store using the stable paragraph ids Document Assembly now emits.
- `Final Issue` — freeze check, Word+PDF+PPTX emission, chat closure.
- Per-file-type parser sub-workflows (`parse-working-file`, `parse-kickoff-deck`,
  `parse-optional-inputs`) are referenced but stubbed — these need the real file
  parsing logic (openpyxl-equivalent for xlsx, pptx text extraction, legacy .ppt
  handling) built out.

## Known open items carried from design

- Overall opinion library is provisional pending Beyon's D-2 confirmation.
- Process-understanding embedded images: no defined placement behavior yet — flagged,
  not silently dropped, per FR-23, but the scope decision (ignore / appendix-only /
  agent-placed) is still with you and Beyon.
- Appendix selection rule beyond the length heuristic (D-4) and the effort baseline
  for the time-saving metric (D-5) are both still open.
