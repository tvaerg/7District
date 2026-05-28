"""
7D Portfolio Aggregation Engine -- Schema Layer
===============================================
The CONTRACT that binds the whole pipeline together. The reasoning layer fills
it, the human-review step edits it, and the renderer reads it. Everything in the
output deck traces back to a field here.

Three kinds of field, kept deliberately separate -- this separation IS the
auditability story we promised 7D:

    STATIC   -- 7D-owned governance facts from the registry (category, lead,
                support, exit timing, baseline thesis). The LLM never writes
                these. Copied in verbatim.

    PROPOSED -- the LLM's judgement for this month: a RAG colour + a short
                reasoning string per heatmap dimension, plus the update-block
                bullets. Always carries WHY, never just a colour.

    OVERRIDE -- 7D's human edit. Optional. If present, it WINS over the proposed
                value at render time. This is the "LLM proposes, 7D adjusts"
                hybrid: the model never has the last word on a colour.

The heatmap mirrors the template columns exactly:
    Portfolio company | 7D involvement (Lead, Support) | Investment thesis |
    Performance | Financing (summary, covenant issue, amount, type, 7D amount,
    timing, 7D rescue amount) | Exit timing

The update blocks mirror the per-company slide:
    Operational update | Financial performance | Three-year value creation |
    7D focus and contribution  -- each with a RAG status + bullets.

Dependencies: pydantic v2.
"""

from __future__ import annotations

from enum import Enum
from typing import Optional
from pydantic import BaseModel, Field


# -----------------------------------------------------------------------------
# RAG colour -- the only legal status values
# -----------------------------------------------------------------------------
class RAG(str, Enum):
    """Red / Amber / Green. Matches the template's green / yellow / red cells."""
    GREEN = "green"
    AMBER = "amber"
    RED = "red"

    @property
    def label(self) -> str:
        return {"green": "On plan", "amber": "Partly on plan", "red": "Off plan"}[self.value]


# -----------------------------------------------------------------------------
# A single proposed-with-reasoning, optionally overridden, judgement
# -----------------------------------------------------------------------------
class RatedField(BaseModel):
    """One heatmap judgement: an LLM-proposed RAG + reasoning, plus optional 7D override.

    `effective_rag` / `effective_text` resolve override-over-proposal so the
    renderer never has to know which one won.
    """
    proposed_rag: RAG
    proposed_text: str = Field(..., description="Short cell text, e.g. 'Improving', 'Weak trading'")
    reasoning: str = Field(..., description="WHY this colour -- the audit trail. Not rendered on the heatmap, kept for review.")

    override_rag: Optional[RAG] = Field(default=None, description="7D's human override colour. Wins if set.")
    override_text: Optional[str] = Field(default=None, description="7D's human override cell text. Wins if set.")
    override_note: Optional[str] = Field(default=None, description="Optional note from 7D explaining the override.")

    @property
    def effective_rag(self) -> RAG:
        return self.override_rag if self.override_rag is not None else self.proposed_rag

    @property
    def effective_text(self) -> str:
        return self.override_text if self.override_text is not None else self.proposed_text

    @property
    def was_overridden(self) -> bool:
        return self.override_rag is not None or self.override_text is not None


# -----------------------------------------------------------------------------
# Financing sub-block -- mirrors the template's Financing super-column
# -----------------------------------------------------------------------------
class Financing(BaseModel):
    """The Financing columns. Most are 'n/a' unless there is an active issue.

    Rule (deterministic, not LLM-guessed): the amount/type/7D-amount/timing/
    rescue fields stay 'n/a' UNLESS `covenant_issue` indicates a real problem or
    a financing event is in play. The reasoning layer sets `has_active_issue`;
    the renderer blanks the detail columns to 'n/a' when it is False.
    """
    summary: RatedField = Field(..., description="Financing summary cell, e.g. 'OK', 'Weak trading', 'Recap discussion with bank'")
    covenant_issue: str = Field(default="No", description="e.g. 'No', 'Yes, breach Q1', 'Discussion with bank'")
    has_active_issue: bool = Field(default=False, description="Gate for whether the detail columns are populated.")

    amount: str = Field(default="n/a", description="Financing amount (MSEK), or 'n/a'")
    type: str = Field(default="n/a", description="e.g. 'Defend', or 'n/a'")
    sevend_amount: str = Field(default="n/a", description="7D Amount, or 'n/a'")
    timing: str = Field(default="n/a", description="e.g. \"Q2'26\", or 'n/a'")
    sevend_rescue_amount: str = Field(default="n/a", description="7D rescue amount, or 'n/a'")


# -----------------------------------------------------------------------------
# Update block -- one of the three per-company narrative sections
# -----------------------------------------------------------------------------
class UpdateBlock(BaseModel):
    """One narrative block on the company-update slide (Operational / Financial /
    Three-year value creation), with a RAG status ball and bullets."""
    proposed_rag: RAG
    bullets: list[str] = Field(default_factory=list, description="3-5 short bullets, owner-relevant.")
    reasoning: str = Field(default="", description="WHY this status -- audit trail.")

    override_rag: Optional[RAG] = None
    override_bullets: Optional[list[str]] = None
    override_note: Optional[str] = None

    @property
    def effective_rag(self) -> RAG:
        return self.override_rag if self.override_rag is not None else self.proposed_rag

    @property
    def effective_bullets(self) -> list[str]:
        return self.override_bullets if self.override_bullets is not None else self.bullets

    @property
    def was_overridden(self) -> bool:
        return self.override_rag is not None or self.override_bullets is not None


# -----------------------------------------------------------------------------
# The whole per-company object
# -----------------------------------------------------------------------------
class CompanyFacts(BaseModel):
    """Everything needed to render ONE company across both the heatmap row and
    its update-slide half. Filled by the reasoning layer, edited by 7D."""

    # --- STATIC (from registry; LLM never writes these) ---
    canonical_name: str
    category: str
    lead: str = "n/a"
    support: str = "n/a"
    exit_timing: str = "n/a"

    # --- PROPOSED + OVERRIDE: heatmap judgements ---
    investment_thesis: RatedField
    performance: RatedField
    financing: Financing

    # --- PROPOSED + OVERRIDE: update-slide blocks ---
    operational: UpdateBlock
    financial: UpdateBlock
    three_year: UpdateBlock
    sevend_focus: list[str] = Field(default_factory=list, description="7D focus & contribution bullets (no RAG).")

    # --- Provenance / QA ---
    source_file: str = ""
    source_kind: str = ""
    has_report_this_cycle: bool = True
    flags: list[str] = Field(default_factory=list, description="QA flags raised during reasoning/aggregation.")

    def any_override(self) -> bool:
        return (
            self.investment_thesis.was_overridden
            or self.performance.was_overridden
            or self.financing.summary.was_overridden
            or self.operational.was_overridden
            or self.financial.was_overridden
            or self.three_year.was_overridden
        )


# -----------------------------------------------------------------------------
# The full portfolio payload (what gets serialised to review.yaml and to the
# renderer's JSON)
# -----------------------------------------------------------------------------
class PortfolioReport(BaseModel):
    """The complete object for one reporting cycle. Serialises to review.yaml for
    7D to edit, then the (possibly edited) version is read back for rendering."""
    cycle_label: str = Field(..., description="e.g. 'April 2026' -- appears on the deck.")
    ceo_update: str = Field(default="", description="Free-text CEO update line on the heatmap slide.")
    companies: list[CompanyFacts] = Field(default_factory=list)

    def in_category_order(self, order: list[str]) -> list[CompanyFacts]:
        rank = {c: i for i, c in enumerate(order)}
        return sorted(self.companies, key=lambda c: rank.get(c.category, 99))

    def overridden_companies(self) -> list[str]:
        return [c.canonical_name for c in self.companies if c.any_override()]


if __name__ == "__main__":
    # Build one minimal CompanyFacts to prove the schema + override resolution.
    cf = CompanyFacts(
        canonical_name="Rapid Images", category="Core Holdings",
        lead="AJ", support="TBD", exit_timing="2028",
        investment_thesis=RatedField(
            proposed_rag=RAG.AMBER, proposed_text="Under transformation",
            reasoning="Business plan being rewritten amid IKEA shock; thesis intact but delayed.",
        ),
        performance=RatedField(
            proposed_rag=RAG.RED, proposed_text="Weak trading",
            reasoning="Q1 EBITDA hit by IKEA engineer reduction; people on the bench.",
        ),
        financing=Financing(
            summary=RatedField(proposed_rag=RAG.RED, proposed_text="Weak trading",
                               reasoning="Will not meet bank covenants for Q1."),
            covenant_issue="Yes, breach Q1", has_active_issue=True,
            amount="5", type="Defend", sevend_amount="2.6", timing="Q2'26",
        ),
        operational=UpdateBlock(proposed_rag=RAG.AMBER,
            bullets=["Healthy pipeline vs budget", "IKEA reorg is a major risk"],
            reasoning="Pipeline ok but IKEA account exposure dominates."),
        financial=UpdateBlock(proposed_rag=RAG.RED,
            bullets=["Weak Q1, IKEA-driven", "Covenant breach, bank dialogue initiated"],
            reasoning="Direct EBITDA hit from bench."),
        three_year=UpdateBlock(proposed_rag=RAG.AMBER,
            bullets=["Business plan prep delayed", "Repositioning critical"],
            reasoning="All-hands-on-deck delayed the plan."),
        sevend_focus=["Support CEO in turnaround", "Finalise new business plan"],
        source_file="input/RapidImages_mars.pptx", source_kind="pptx_text",
    )

    # Apply a 7D override and confirm it wins.
    cf.performance.override_rag = RAG.AMBER
    cf.performance.override_text = "Stabilising"
    cf.performance.override_note = "Board sees early stabilisation post cost actions."

    print("Schema OK.")
    print(f"  Performance proposed: {cf.performance.proposed_rag.value} / {cf.performance.proposed_text}")
    print(f"  Performance effective: {cf.performance.effective_rag.value} / {cf.performance.effective_text}  (overridden={cf.performance.was_overridden})")
    print(f"  Any override on company: {cf.any_override()}")
    print(f"  Financing detail gated by issue: amount={cf.financing.amount}, active={cf.financing.has_active_issue}")

    pr = PortfolioReport(cycle_label="April 2026", ceo_update="Mixed month; one covenant breach.", companies=[cf])
    print(f"  Portfolio overrides: {pr.overridden_companies()}")
    # Round-trip through JSON to prove serialisability for review.yaml / renderer.
    js = pr.model_dump_json(indent=2)
    back = PortfolioReport.model_validate_json(js)
    print(f"  JSON round-trip OK: {back.companies[0].performance.effective_text}")