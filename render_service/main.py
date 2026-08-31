"""
IARD Render Service — self-hosted document assembly for the Internal Audit
Report Drafter Agent.

Replaces the n8n Document Assembly workflow's fabricated "microsoftWord" node.
n8n calls POST /render-report with the current engagement state (front matter +
observations + appendix items); this service fills Beyon's actual Word template
and returns the rendered .docx as bytes.

Design constraints carried over from the BRD / HLD:
 - GR-02 template fidelity: structure, table shapes, and rating colours come
   from the fixed template, never restyled by this service.
 - Verbatim sections (objectives, scope, limitations, risk classification) are
   inserted exactly as received — this service does not paraphrase anything.
 - Each observation paragraph gets a stable bookmark named obs-<observation_id>
   so a later re-uploaded, auditor-edited Word file can be diffed back to the
   right row (feeds the not-yet-built Round-trip Reconciliation sub-workflow).
 - area_for_improvement observations get no rating colour and are excluded
   from the Audit Issues Dashboard counts and the List of Audit Issues (FR-26).
"""

import io
import os
import copy
from typing import Optional

from fastapi import FastAPI, HTTPException
from fastapi.responses import StreamingResponse
from pydantic import BaseModel
from docx import Document
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

# Beyon's fixed template is Restricted-classified (BRD Section 10) and is never committed to git
# (see Dockerfile). It's provided to the running container instead, and the location depends on
# how it's mounted:
#   - Render Secret File (recommended for the pilot): dashboard-uploaded files land at
#     /etc/secrets/<filename>, NOT in the app's working directory — hence this default.
#   - Render persistent Disk: set TEMPLATE_PATH to wherever the disk is mounted, e.g.
#     /var/data/report_template.docx.
# Either way, override via the TEMPLATE_PATH env var rather than editing this default in code.
TEMPLATE_PATH = os.environ.get("TEMPLATE_PATH", "/etc/secrets/report_template.docx")

RATING_LABELS = {
    "active_management_very_high": "Active Management (Very High)",
    "continuous_review_high": "Continuous Review (High)",
    "periodic_monitoring_medium": "Periodic Monitoring (Medium)",
    "no_major_concern_low": "No Major Concern (Low)",
    "area_for_improvement": "Area for Improvement",
}
# RGB fills matching the template's existing colour scheme for each rating.
RATING_COLOURS = {
    "active_management_very_high": "C00000",
    "continuous_review_high": "FF0000",
    "periodic_monitoring_medium": "FFC000",
    "no_major_concern_low": "92D050",
    "area_for_improvement": None,  # no colour, per FR-26
}
RISK_LETTER = {
    "active_management_very_high": "VH",
    "continuous_review_high": "H",
    "periodic_monitoring_medium": "M",
    "no_major_concern_low": "L",
    "area_for_improvement": "AFI",
}

app = FastAPI(title="IARD Render Service")


class RootCause(BaseModel):
    category_people: bool = False
    category_process: bool = False
    category_technology: bool = False
    explanation: Optional[str] = None
    flagged: bool = False  # evidence insufficient to support this specific cause (GR-04, now per-cause)


class Observation(BaseModel):
    id: str
    sequence_no: int
    title: Optional[str] = None
    criteria: Optional[str] = None
    condition: Optional[str] = None
    root_causes: list[RootCause] = []  # v0.2: plural, each categorised People/Process/Technology — replaces the old single `cause` string
    risk: Optional[str] = None
    recommendation: Optional[str] = None
    risk_rating: Optional[str] = None
    management_action: Optional[str] = None
    owner: Optional[str] = None
    target_date: Optional[str] = None


class AppendixItem(BaseModel):
    observation_id: str
    appendix_number: str
    content_ref: str


class FrontMatter(BaseModel):
    cover_page: dict
    objectives: Optional[list] = None
    scope: Optional[dict] = None
    limitations_of_internal_audit: Optional[str] = None
    risk_classification: Optional[str] = None
    acknowledgment: Optional[str] = None
    audit_team: Optional[dict] = None
    overall_opinion_paragraph: Optional[str] = None
    background: Optional[str] = None


class RenderRequest(BaseModel):
    engagement_id: str
    front_matter: FrontMatter
    observations: list[Observation]
    appendix_items: list[AppendixItem] = []


def add_bookmark(paragraph, bookmark_id: int, name: str):
    """Wrap a paragraph in a Word bookmark so a re-uploaded edit can be traced
    back to the observation it came from (round-trip reconciliation)."""
    start = OxmlElement("w:bookmarkStart")
    start.set(qn("w:id"), str(bookmark_id))
    start.set(qn("w:name"), name)
    end = OxmlElement("w:bookmarkEnd")
    end.set(qn("w:id"), str(bookmark_id))
    paragraph._p.insert(0, start)
    paragraph._p.append(end)


def find_heading(doc: Document, text_startswith: str):
    for p in doc.paragraphs:
        if p.text.strip().lower().startswith(text_startswith.lower()):
            return p
    return None


def insert_paragraph_after(paragraph, text: str = "", style=None, bold: bool = False):
    new_p = OxmlElement("w:p")
    paragraph._p.addnext(new_p)
    from docx.text.paragraph import Paragraph
    new_para = Paragraph(new_p, paragraph._parent)
    if text:
        run = new_para.add_run(text)
        if bold:
            run.bold = True
    if style:
        new_para.style = style
    return new_para


def root_cause_category_label(rc: "RootCause") -> str:
    cats = []
    if rc.category_people:
        cats.append("People")
    if rc.category_process:
        cats.append("Process")
    if rc.category_technology:
        cats.append("Technology")
    return ", ".join(cats)


def shade_cell(cell, hex_colour: str):
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:fill"), hex_colour)
    tc_pr.append(shd)


def render_report(req: RenderRequest) -> bytes:
    doc = Document(TEMPLATE_PATH)
    fm = req.front_matter

    # --- Cover page (verbatim + intake) ---
    cover = fm.cover_page
    for p in doc.paragraphs[:15]:
        t = p.text
        if "Project title" in t or "{{title}}" in t:
            p.text = cover.get("title", "")
        elif "Quarter and Year" in t or "{{quarter_year}}" in t:
            p.text = f"{cover.get('quarter', '')} {cover.get('year', '')}"
        elif "Draft Report" in t or "{{status}}" in t:
            p.text = cover.get("status", "Draft Report")

    # --- Background (generated) ---
    bg_heading = find_heading(doc, "Background")
    if bg_heading and fm.background:
        insert_paragraph_after(bg_heading, fm.background)

    # --- Objectives (verbatim) ---
    obj_heading = find_heading(doc, "Objectives:")
    if obj_heading and fm.objectives:
        anchor = obj_heading
        for item in fm.objectives:
            anchor = insert_paragraph_after(anchor, f"• {item}")

    # --- Scope (verbatim) ---
    scope_heading = find_heading(doc, "Scope:")
    if scope_heading and fm.scope:
        anchor = scope_heading
        for item in (fm.scope.get("items") or []):
            anchor = insert_paragraph_after(anchor, f"• {item}")
        if fm.scope.get("limitations"):
            anchor = insert_paragraph_after(anchor, fm.scope["limitations"])

    # --- Acknowledgment (fill) ---
    ack_heading = find_heading(doc, "Acknowledgment")
    if ack_heading and fm.acknowledgment:
        insert_paragraph_after(ack_heading, fm.acknowledgment)

    # --- Overall opinion (verbatim + selection) ---
    opinion_heading = find_heading(doc, "Overall opinion")
    if opinion_heading and fm.overall_opinion_paragraph:
        insert_paragraph_after(opinion_heading, fm.overall_opinion_paragraph)

    # --- Audit team (verbatim + fill) ---
    if fm.audit_team:
        for table in doc.tables:
            header_text = " ".join(c.text for c in table.rows[0].cells).lower()
            if "audit team" in header_text or len(table.rows) <= 3 and table.columns and len(table.columns) == 3:
                # best-effort: fill first data row with auditor/audit_manager/chief names
                try:
                    row = table.rows[1] if len(table.rows) > 1 else table.rows[0]
                    row.cells[0].text = fm.audit_team.get("auditor", "")
                    row.cells[1].text = fm.audit_team.get("audit_manager", "")
                    row.cells[2].text = fm.audit_team.get("chief_of_internal_audit", "")
                except IndexError:
                    pass
                break

    # --- Audit Issues Dashboard + List of Audit Issues (derived; exclude area_for_improvement) ---
    listable = [o for o in req.observations if o.risk_rating != "area_for_improvement"]

    dashboard_table = None
    list_table = None
    for table in doc.tables:
        header_text = " ".join(c.text for c in table.rows[0].cells).lower()
        if "business area" in header_text:
            dashboard_table = table
        if header_text.strip().startswith("n") and "audit issue" in header_text:
            list_table = table

    if list_table:
        for idx, obs in enumerate(listable, start=1):
            row_cells = list_table.add_row().cells
            row_cells[0].text = str(idx)
            row_cells[1].text = RISK_LETTER.get(obs.risk_rating, "")
            row_cells[2].text = obs.title or ""
            # Owner/target date deliberately left blank in draft (FR-16 / content rule)

    # --- Detailed Audit Issues (generated; one block per observation) ---
    section_heading = find_heading(doc, "Detailed Audit Issues")
    anchor = section_heading
    bookmark_counter = 1
    for obs in req.observations:
        label = RATING_LABELS.get(obs.risk_rating, obs.risk_rating or "")
        colour = RATING_COLOURS.get(obs.risk_rating)

        title_p = insert_paragraph_after(anchor, f"{obs.sequence_no}. {obs.title or '[untitled]'}  —  {label}")
        add_bookmark(title_p, bookmark_counter, f"obs-{obs.id}")
        bookmark_counter += 1
        anchor = title_p

        if obs.criteria:
            anchor = insert_paragraph_after(anchor, f"Criteria: {obs.criteria}")
        if obs.condition:
            anchor = insert_paragraph_after(anchor, f"Condition: {obs.condition}")
        # v0.2: root causes are plural, each categorised People/Process/Technology (or a combination),
        # each with its own explanation. Verified against the real template and issued CVM report —
        # there is no precedent there for multiple/tagged causes (that's new in v0.2), so this keeps
        # the template's existing convention (one bold "Potential Cause" label, plain text under it)
        # rather than inventing a new heading or table, and lists one cause per line under that label.
        # Confirmed with Nuella 2026-08-25.
        if obs.root_causes:
            anchor = insert_paragraph_after(anchor, "Potential Cause", bold=True)
            for rc in obs.root_causes:
                if rc.flagged or not rc.explanation:
                    anchor = insert_paragraph_after(
                        anchor, "[FLAGGED — evidence insufficient to support this cause; auditor input required]"
                    )
                else:
                    label = root_cause_category_label(rc)
                    prefix = f"({label}) — " if label else ""
                    anchor = insert_paragraph_after(anchor, f"{prefix}{rc.explanation}")
        elif obs.risk_rating != "area_for_improvement":
            anchor = insert_paragraph_after(anchor, "Potential Cause", bold=True)
            anchor = insert_paragraph_after(anchor, "[FLAGGED — evidence insufficient to support a cause; auditor input required]")
        if obs.risk:
            anchor = insert_paragraph_after(anchor, f"Risk: {obs.risk}")
        if obs.recommendation:
            anchor = insert_paragraph_after(anchor, f"Recommendation: {obs.recommendation}")

        # Management Action: never written by the agent — only rendered if the auditor already entered it.
        anchor = insert_paragraph_after(anchor, f"Management Action: {obs.management_action or ''}")

        appx = [a for a in req.appendix_items if a.observation_id == obs.id]
        if appx:
            refs = "; ".join(f"Appendix {a.appendix_number}: {a.content_ref}" for a in appx)
            anchor = insert_paragraph_after(anchor, f"See: {refs}")

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf.read()


@app.post("/render-report")
def render_report_endpoint(req: RenderRequest):
    try:
        docx_bytes = render_report(req)
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"render failed: {e}")
    return StreamingResponse(
        io.BytesIO(docx_bytes),
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        headers={"Content-Disposition": f'attachment; filename="{req.engagement_id}.docx"'},
    )


@app.get("/health")
def health():
    # Surfaces the template-availability problem directly, instead of only discovering it on the
    # first real /render-report call (which is how we found it — Aug 28/31 failures both traced
    # back to this). Still 200s so this doesn't flap the service's deploy health check; the
    # template_found field is what to check.
    return {
        "status": "ok",
        "template_path": TEMPLATE_PATH,
        "template_found": os.path.isfile(TEMPLATE_PATH),
    }
