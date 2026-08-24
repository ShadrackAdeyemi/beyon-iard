-- Internal Audit Report Drafter Agent — state store schema
-- Postgres. This is the authoritative record; chat memory and LLM context are derived from it, never the reverse.

CREATE TABLE engagement (
    id                      UUID PRIMARY KEY DEFAULT gen_random_uuid(),
    title                   TEXT NOT NULL,
    quarter                 TEXT NOT NULL,
    year                    INT NOT NULL,
    entity                  TEXT NOT NULL,               -- drives cover-page logo selection
    overall_opinion_label   TEXT,                         -- one of the 5 SOP ratings, set at intake
    status                  TEXT NOT NULL DEFAULT 'draft' CHECK (status IN ('draft', 'final')),
    acknowledged_department TEXT,                         -- editable at intake, pre-filled from kickoff deck
    audit_team              JSONB,                         -- {auditor: name, audit_manager: name, chief_of_internal_audit: name}
    chat_session_id         TEXT UNIQUE,                   -- resolves incoming chat messages to this engagement
    input_gaps              JSONB DEFAULT '[]'::jsonb,      -- flagged missing/optional inputs (FR-23)
    created_at              TIMESTAMPTZ NOT NULL DEFAULT now(),
    finalized_at            TIMESTAMPTZ
);

CREATE TABLE source_document (
    id              UUID PRIMARY KEY DEFAULT gen_random_uuid(),
    engagement_id   UUID NOT NULL REFERENCES engagement(id) ON DELETE CASCADE,
    doc_type        TEXT NOT NULL CHECK (doc_type IN (
                        'completed_audit_working_file', 'kickoff_deck', 'process_understanding',
                        'audit_notification', 'analysis_sheet', 'supporting_evidence')),
    file_ref        TEXT NOT NULL,                        -- pointer into file store
    parsed_at       TIMESTAMPTZ,
    parse_status    TEXT DEFAULT 'pending' CHECK (parse_status IN ('pending', 'ok', 'partial', 'failed')),
    notes           TEXT
);

CREATE TABLE working_file_row (
    id                  UUID PRIMARY KEY DEFAULT gen_random_uuid(),
    engagement_id       UUID NOT NULL REFERENCES engagement(id) ON DELETE CASCADE,
    source_document_id  UUID REFERENCES source_document(id),
    row_index           INT NOT NULL,                     -- position in the source Excel, for traceability
    risk_description     TEXT,
    control_description  TEXT,
    test_steps            TEXT,
    work_done              TEXT,
    test_status             TEXT,
    test_result              TEXT,
    test_result_summary       TEXT,
    residual_risk_level        TEXT,                       -- Very High / High / Medium / Low, drives rating colour
    evidence_refs                TEXT[],                    -- filenames / attachment refs
    is_reportable                 BOOLEAN,                   -- decided during Observation Drafting
    linked_observation_id         UUID                       -- set once drafted into an observation (nullable)
);

CREATE TABLE observation (
    id                 UUID PRIMARY KEY DEFAULT gen_random_uuid(),
    engagement_id      UUID NOT NULL REFERENCES engagement(id) ON DELETE CASCADE,
    sequence_no        INT NOT NULL,                        -- display order / issue number
    title              TEXT,
    criteria           TEXT,
    condition          TEXT,
    cause              TEXT,
    cause_flagged      BOOLEAN DEFAULT FALSE,                -- true = evidence insufficient, agent asked rather than invented (GR-04)
    risk               TEXT,
    recommendation     TEXT,
    risk_rating        TEXT CHECK (risk_rating IN (
                           'active_management_very_high', 'continuous_review_high',
                           'periodic_monitoring_medium', 'no_major_concern_low',
                           'area_for_improvement')),        -- area_for_improvement = no rating colour, excluded from dashboard/issue list (FR-26)
    source_refs        UUID[],                               -- working_file_row ids this observation traces to
    approval_status    TEXT NOT NULL DEFAULT 'draft' CHECK (approval_status IN ('draft', 'approved', 'frozen')),
    management_action  TEXT,                                 -- never written by the agent (BRD #15)
    owner              TEXT,                                 -- auditor-entered, post-draft
    target_date        DATE,                                 -- auditor-entered, post-draft
    last_edited_by     TEXT CHECK (last_edited_by IN ('agent', 'human')),
    last_edited_at     TIMESTAMPTZ NOT NULL DEFAULT now()
);

CREATE TABLE appendix_item (
    id              UUID PRIMARY KEY DEFAULT gen_random_uuid(),
    observation_id  UUID NOT NULL REFERENCES observation(id) ON DELETE CASCADE,
    appendix_number TEXT NOT NULL,                          -- e.g. "1.b", auto-numbered
    content_ref     TEXT NOT NULL,                          -- pointer into file store or inline text
    direction       TEXT NOT NULL CHECK (direction IN ('auto', 'manual')),  -- auto = length heuristic, manual = auditor instruction
    created_by      TEXT NOT NULL CHECK (created_by IN ('agent', 'human')),
    created_at      TIMESTAMPTZ NOT NULL DEFAULT now()
);

CREATE TABLE change_log (
    id              UUID PRIMARY KEY DEFAULT gen_random_uuid(),
    engagement_id   UUID NOT NULL REFERENCES engagement(id) ON DELETE CASCADE,
    observation_id  UUID REFERENCES observation(id),        -- nullable: some actions are engagement-level
    actor           TEXT NOT NULL CHECK (actor IN ('agent', 'human')),
    action          TEXT NOT NULL,                           -- e.g. 'draft', 'merge', 'split', 'reword', 're-rate', 'approve', 'move_to_appendix', 'final_issue'
    before_snapshot JSONB,
    after_snapshot  JSONB,
    timestamp       TIMESTAMPTZ NOT NULL DEFAULT now()
);

CREATE INDEX idx_observation_engagement ON observation(engagement_id);
CREATE INDEX idx_working_file_row_engagement ON working_file_row(engagement_id);
CREATE INDEX idx_change_log_engagement ON change_log(engagement_id);
CREATE UNIQUE INDEX idx_engagement_chat_session ON engagement(chat_session_id);

-- Guardrail at the data layer, not just app logic: a frozen observation cannot be overwritten
-- by anything except an explicit unfreeze action (out of scope for pilot per the "no going back" decision
-- on the second scoping call — Luigi: "I would avoid for the pilot"). Enforced here via trigger.
CREATE OR REPLACE FUNCTION prevent_frozen_overwrite() RETURNS TRIGGER AS $$
BEGIN
    IF OLD.approval_status = 'frozen' AND NEW.approval_status = 'frozen'
       AND (NEW.criteria IS DISTINCT FROM OLD.criteria
            OR NEW.condition IS DISTINCT FROM OLD.condition
            OR NEW.cause IS DISTINCT FROM OLD.cause
            OR NEW.risk IS DISTINCT FROM OLD.risk
            OR NEW.recommendation IS DISTINCT FROM OLD.recommendation
            OR NEW.risk_rating IS DISTINCT FROM OLD.risk_rating) THEN
        RAISE EXCEPTION 'Observation % is frozen and cannot be modified (GR-01)', OLD.id;
    END IF;
    RETURN NEW;
END;
$$ LANGUAGE plpgsql;

CREATE TRIGGER trg_prevent_frozen_overwrite
    BEFORE UPDATE ON observation
    FOR EACH ROW EXECUTE FUNCTION prevent_frozen_overwrite();
