-- Internal Audit Report Drafter Agent — v0.2 BRD schema migration
-- Run this against the same Supabase project as 001_schema.sql.
-- Reflects the three standing decisions from the v0.2 BRD review (2026-08-25):
--   1. Report template / summary template / entity logos / opinion library move to held configuration.
--   2. Root causes become plural, each tagged People/Process/Technology (or combination), each with its own explanation.
--   3. Approval moves from per-observation freeze to a single report-level approval (engagement.status already
--      supports this — 'draft'/'final' — so this migration removes the now-redundant per-row freeze machinery).
--
-- Also includes light, non-breaking schema prep for two BRD items that are NOT yet built (FR-27 input
-- consistency check, FR-28 regeneration-as-new-version): columns only, no workflow logic depends on them yet.

-- ============================================================
-- 1. Held configuration (report template, summary template, entity logos, opinion library)
-- ============================================================
-- These no longer arrive as per-engagement uploads (v0.2 Section 6.2). One row per config item;
-- entity_logo has one row per Beyon entity, everything else is a singleton row per config_type.
CREATE TABLE configuration (
    id              UUID PRIMARY KEY DEFAULT gen_random_uuid(),
    config_type     TEXT NOT NULL CHECK (config_type IN (
                        'report_template', 'summary_template', 'entity_logo', 'opinion_library', 'ia_sop_reference')),
    entity          TEXT,                                    -- only set (and required) when config_type = 'entity_logo'
    file_ref        TEXT,                                    -- pointer into file store, for template/logo/reference docs
    content         JSONB,                                   -- structured content, e.g. the 5-rating opinion library
    version         INT NOT NULL DEFAULT 1,
    updated_by      TEXT,
    updated_at      TIMESTAMPTZ NOT NULL DEFAULT now(),
    CONSTRAINT chk_entity_logo_has_entity CHECK (
        (config_type = 'entity_logo' AND entity IS NOT NULL) OR
        (config_type != 'entity_logo' AND entity IS NULL)
    )
);

-- This is kept as an append-only version history table rather than enforcing "one active row" at the
-- database layer (Postgres partial indexes can't reference MAX() of the same table). Workflows always
-- read the current value as: SELECT * FROM configuration WHERE config_type = ? AND entity = ?
-- ORDER BY version DESC LIMIT 1;
CREATE INDEX idx_configuration_type_entity ON configuration(config_type, entity);

-- ============================================================
-- 2. Root causes: plural, categorised, explained
-- ============================================================
-- Replaces observation.cause / observation.cause_flagged (single-cause model) with a child table.
CREATE TABLE observation_root_cause (
    id                   UUID PRIMARY KEY DEFAULT gen_random_uuid(),
    observation_id       UUID NOT NULL REFERENCES observation(id) ON DELETE CASCADE,
    sequence_no          INT NOT NULL DEFAULT 1,              -- display order within the observation's cause list
    category_people      BOOLEAN NOT NULL DEFAULT FALSE,
    category_process     BOOLEAN NOT NULL DEFAULT FALSE,
    category_technology  BOOLEAN NOT NULL DEFAULT FALSE,
    explanation          TEXT,                                -- required unless flagged (evidence insufficient)
    flagged              BOOLEAN NOT NULL DEFAULT FALSE,       -- true = evidence insufficient, agent asked rather than invented (GR-04, per-cause now)
    last_edited_by       TEXT CHECK (last_edited_by IN ('agent', 'human')),
    last_edited_at       TIMESTAMPTZ NOT NULL DEFAULT now(),
    CONSTRAINT chk_cause_has_category_or_flag CHECK (
        flagged = TRUE OR category_people OR category_process OR category_technology
    ),
    CONSTRAINT chk_cause_has_explanation_unless_flagged CHECK (
        flagged = TRUE OR explanation IS NOT NULL
    )
);

CREATE INDEX idx_observation_root_cause_observation ON observation_root_cause(observation_id);

ALTER TABLE observation DROP COLUMN IF EXISTS cause;
ALTER TABLE observation DROP COLUMN IF EXISTS cause_flagged;

-- ============================================================
-- 3. Approval: report-level only — drop the per-observation freeze
-- ============================================================
-- engagement.status already carries 'draft' / 'final' (001_schema.sql) — that IS the report-level approval
-- gate now. observation.approval_status becomes redundant and is removed, along with the trigger that
-- enforced freezing at the row level.
DROP TRIGGER IF EXISTS trg_prevent_frozen_overwrite ON observation;
DROP FUNCTION IF EXISTS prevent_frozen_overwrite();

ALTER TABLE observation DROP COLUMN IF EXISTS approval_status;

-- Replacement guardrail, enforced at the engagement level: once an engagement is 'final', no observation
-- belonging to it can be modified by anyone (agent or human) except through an explicit, separate
-- unfinalize action — out of scope for the pilot, same "no going back" stance as before, just moved up
-- a level to match where approval now actually happens.
CREATE OR REPLACE FUNCTION prevent_edit_after_final() RETURNS TRIGGER AS $$
DECLARE
    eng_status TEXT;
BEGIN
    SELECT status INTO eng_status FROM engagement WHERE id = NEW.engagement_id;
    IF eng_status = 'final' THEN
        RAISE EXCEPTION 'Engagement % is Final and its observations cannot be modified (GR-01)', NEW.engagement_id;
    END IF;
    RETURN NEW;
END;
$$ LANGUAGE plpgsql;

CREATE TRIGGER trg_prevent_edit_after_final
    BEFORE UPDATE ON observation
    FOR EACH ROW EXECUTE FUNCTION prevent_edit_after_final();

-- Same guardrail on root causes, since they're edited independently of their parent observation row.
CREATE OR REPLACE FUNCTION prevent_cause_edit_after_final() RETURNS TRIGGER AS $$
DECLARE
    eng_status TEXT;
BEGIN
    SELECT e.status INTO eng_status
    FROM engagement e
    JOIN observation o ON o.engagement_id = e.id
    WHERE o.id = COALESCE(NEW.observation_id, OLD.observation_id);
    IF eng_status = 'final' THEN
        RAISE EXCEPTION 'Parent engagement is Final; root causes cannot be modified (GR-01)';
    END IF;
    RETURN NEW;
END;
$$ LANGUAGE plpgsql;

CREATE TRIGGER trg_prevent_cause_edit_after_final
    BEFORE UPDATE OR DELETE ON observation_root_cause
    FOR EACH ROW EXECUTE FUNCTION prevent_cause_edit_after_final();

-- ============================================================
-- 4. Audit notification is now mandatory
-- ============================================================
-- No schema change needed: 'audit_notification' was already a valid source_document.doc_type value
-- in 001_schema.sql. Enforcement moves entirely to Intake & Parse's mandatory-inputs list (workflow change,
-- not a schema change) — noted here for completeness.

-- ============================================================
-- 5. Prep only, not yet wired to any workflow — FR-27 (input consistency check)
-- ============================================================
ALTER TABLE working_file_row
    ADD COLUMN IF NOT EXISTS consistency_status TEXT NOT NULL DEFAULT 'unchecked'
        CHECK (consistency_status IN ('unchecked', 'consistent', 'flagged_inconsistent', 'confirmed_by_auditor')),
    ADD COLUMN IF NOT EXISTS consistency_notes TEXT;          -- what mismatch was found, if any

-- ============================================================
-- 6. Prep only, not yet wired to any workflow — FR-28 (regeneration as a new version)
-- ============================================================
ALTER TABLE engagement
    ADD COLUMN IF NOT EXISTS source_version INT NOT NULL DEFAULT 1;   -- increments each time the agent regenerates from updated source

COMMENT ON COLUMN engagement.source_version IS
    'Increments on FR-28 full regeneration from updated source. Full version-history modeling deferred until Observation Drafting + FR-28 are actually built.';