"""
7D Portfolio Aggregation Engine -- Review Bridge (review.yaml)
=============================================================
The human-in-the-loop bridge. This is what makes "the LLM proposes, 7D adjusts"
concrete rather than a slogan.

Flow:
    reasoning layer -> PortfolioReport  ──write_review_yaml()──>  review.yaml
                                                                      │
                                                          7D edits override: fields
                                                                      │
    renderer  <──read_review_yaml()──  review.yaml (now with 7D's edits)

Why YAML, not the raw JSON:
    The renderer can read JSON fine. But a human cannot comfortably *edit* a deep
    JSON blob. review.yaml is shaped for the editor's eye:
      * every dimension shows the LLM's proposed RAG + the REASONING inline as a
        comment, so 7D sees WHY before deciding whether to override;
      * the override slots sit empty and obvious, ready to fill;
      * filling an override is the only edit needed -- leave it blank to accept
        the proposal.

The contract: editing review.yaml NEVER changes proposed_* values. 7D only ever
writes into override_* fields. read_review_yaml folds those overrides back onto
the PortfolioReport, where RatedField.effective_* resolves the winner.

Dependencies: pyyaml, pydantic v2, sevend_schema.
"""

from __future__ import annotations

from pathlib import Path
from typing import Any
import yaml

from sevend_schema import (
    PortfolioReport, CompanyFacts, RatedField, Financing, UpdateBlock, RAG,
)


# -----------------------------------------------------------------------------
# WRITE: PortfolioReport -> review.yaml (human-editable)
# -----------------------------------------------------------------------------
def _rated_to_yaml(rf: RatedField) -> dict:
    """One RatedField as an editable block. Proposal + reasoning visible; override blank."""
    return {
        "proposed_rag": rf.proposed_rag.value,
        "proposed_text": rf.proposed_text,
        "reasoning": rf.reasoning,
        # --- 7D edits below this line only ---
        "override_rag": rf.override_rag.value if rf.override_rag else None,
        "override_text": rf.override_text,
        "override_note": rf.override_note,
    }


def _block_to_yaml(b: UpdateBlock) -> dict:
    return {
        "proposed_rag": b.proposed_rag.value,
        "bullets": list(b.bullets),
        "reasoning": b.reasoning,
        # --- 7D edits below this line only ---
        "override_rag": b.override_rag.value if b.override_rag else None,
        "override_bullets": b.override_bullets,
        "override_note": b.override_note,
    }


def _company_to_yaml(c: CompanyFacts) -> dict:
    return {
        "canonical_name": c.canonical_name,
        "category": c.category,
        # static -- shown for context, NOT meant to be edited here (edit the registry)
        "_static": {
            "lead": c.lead, "support": c.support, "exit_timing": c.exit_timing,
        },
        "has_report_this_cycle": c.has_report_this_cycle,
        "flags": list(c.flags),
        "investment_thesis": _rated_to_yaml(c.investment_thesis),
        "performance": _rated_to_yaml(c.performance),
        "financing": {
            "summary": _rated_to_yaml(c.financing.summary),
            "covenant_issue": c.financing.covenant_issue,
            "has_active_issue": c.financing.has_active_issue,
            "amount": c.financing.amount,
            "type": c.financing.type,
            "sevend_amount": c.financing.sevend_amount,
            "timing": c.financing.timing,
            "sevend_rescue_amount": c.financing.sevend_rescue_amount,
        },
        "operational": _block_to_yaml(c.operational),
        "financial": _block_to_yaml(c.financial),
        "three_year": _block_to_yaml(c.three_year),
        "sevend_focus": list(c.sevend_focus),
    }


_HEADER = """\
# =============================================================================
# 7D PORTFOLIO REVIEW FILE
# =============================================================================
# The engine proposed the RAG colours and text below. Each dimension shows:
#   proposed_rag / proposed_text  -- the engine's suggestion
#   reasoning                     -- WHY (audit trail; not printed on the deck)
#
# TO ADJUST: fill in the override_* fields. Leave them null to accept the
# proposal. You only ever edit override_* fields -- never the proposed_* ones.
#   override_rag:  one of  green | amber | red
#   override_text: the cell text you want instead
#   override_note: optional reason for your change (kept for the record)
#
# When done, save and re-run the renderer. Overrides win over proposals.
# =============================================================================

"""


def write_review_yaml(report: PortfolioReport, path: str | Path) -> Path:
    payload = {
        "cycle_label": report.cycle_label,
        "ceo_update": report.ceo_update,
        "companies": [_company_to_yaml(c) for c in report.companies],
    }
    p = Path(path)
    p.parent.mkdir(parents=True, exist_ok=True)
    with open(p, "w", encoding="utf-8") as f:
        f.write(_HEADER)
        yaml.safe_dump(payload, f, sort_keys=False, allow_unicode=True, width=100)
    return p


# -----------------------------------------------------------------------------
# READ: review.yaml (with 7D's edits) -> PortfolioReport
# -----------------------------------------------------------------------------
def _rag_or_none(v: Any) -> RAG | None:
    if v in (None, "", "null"):
        return None
    try:
        return RAG(str(v).strip().lower())
    except ValueError:
        raise ValueError(f"Invalid RAG value in review.yaml: {v!r} (use green|amber|red)")


def _rated_from_yaml(d: dict) -> RatedField:
    return RatedField(
        proposed_rag=RAG(d["proposed_rag"]),
        proposed_text=d["proposed_text"],
        reasoning=d.get("reasoning", ""),
        override_rag=_rag_or_none(d.get("override_rag")),
        override_text=d.get("override_text") or None,
        override_note=d.get("override_note") or None,
    )


def _block_from_yaml(d: dict) -> UpdateBlock:
    ob = d.get("override_bullets")
    return UpdateBlock(
        proposed_rag=RAG(d["proposed_rag"]),
        bullets=list(d.get("bullets", [])),
        reasoning=d.get("reasoning", ""),
        override_rag=_rag_or_none(d.get("override_rag")),
        override_bullets=list(ob) if ob else None,
        override_note=d.get("override_note") or None,
    )


def _company_from_yaml(d: dict) -> CompanyFacts:
    st = d.get("_static", {})
    fin = d["financing"]
    return CompanyFacts(
        canonical_name=d["canonical_name"],
        category=d["category"],
        lead=st.get("lead", "n/a"),
        support=st.get("support", "n/a"),
        exit_timing=st.get("exit_timing", "n/a"),
        has_report_this_cycle=d.get("has_report_this_cycle", True),
        flags=list(d.get("flags", [])),
        investment_thesis=_rated_from_yaml(d["investment_thesis"]),
        performance=_rated_from_yaml(d["performance"]),
        financing=Financing(
            summary=_rated_from_yaml(fin["summary"]),
            covenant_issue=fin.get("covenant_issue", "No"),
            has_active_issue=fin.get("has_active_issue", False),
            amount=fin.get("amount", "n/a"),
            type=fin.get("type", "n/a"),
            sevend_amount=fin.get("sevend_amount", "n/a"),
            timing=fin.get("timing", "n/a"),
            sevend_rescue_amount=fin.get("sevend_rescue_amount", "n/a"),
        ),
        operational=_block_from_yaml(d["operational"]),
        financial=_block_from_yaml(d["financial"]),
        three_year=_block_from_yaml(d["three_year"]),
        sevend_focus=list(d.get("sevend_focus", [])),
    )


def read_review_yaml(path: str | Path) -> PortfolioReport:
    with open(path, "r", encoding="utf-8") as f:
        data = yaml.safe_load(f)
    return PortfolioReport(
        cycle_label=data.get("cycle_label", ""),
        ceo_update=data.get("ceo_update", ""),
        companies=[_company_from_yaml(c) for c in data.get("companies", [])],
    )


if __name__ == "__main__":
    # Build a tiny report, write review.yaml, simulate a 7D edit, read it back,
    # and confirm the override folded through to effective_*.
    from sevend_schema import CompanyFacts as CF  # noqa
    cf = CompanyFacts(
        canonical_name="Rapid Images", category="Core Holdings",
        lead="AJ", support="TBD", exit_timing="2028",
        investment_thesis=RatedField(proposed_rag=RAG.AMBER, proposed_text="Under transformation",
                                     reasoning="Plan being rewritten amid IKEA shock."),
        performance=RatedField(proposed_rag=RAG.RED, proposed_text="Weak trading",
                               reasoning="Q1 EBITDA hit by IKEA engineer reduction."),
        financing=Financing(
            summary=RatedField(proposed_rag=RAG.RED, proposed_text="Weak trading",
                               reasoning="Covenant breach Q1."),
            covenant_issue="Yes, breach Q1", has_active_issue=True,
            amount="5", type="Defend", sevend_amount="2.6", timing="Q2'26"),
        operational=UpdateBlock(proposed_rag=RAG.AMBER, bullets=["Healthy pipeline", "IKEA risk"],
                                reasoning="Pipeline ok, IKEA dominates."),
        financial=UpdateBlock(proposed_rag=RAG.RED, bullets=["Weak Q1", "Bank dialogue"],
                              reasoning="Direct EBITDA hit."),
        three_year=UpdateBlock(proposed_rag=RAG.AMBER, bullets=["Plan delayed"],
                               reasoning="All-hands delayed plan."),
        sevend_focus=["Support turnaround", "Finalise plan"],
        source_file="input/RapidImages_mars.pptx", source_kind="pptx_text",
    )
    pr = PortfolioReport(cycle_label="April 2026", ceo_update="Mixed month.", companies=[cf])

    out = write_review_yaml(pr, "output/review.yaml")
    print(f"Wrote {out}")

    # Simulate 7D editing the file: set a performance override (structured edit,
    # the way a human or a small UI would write it back).
    edited = yaml.safe_load(Path("output/review.yaml").read_text(encoding="utf-8"))
    edited["companies"][0]["performance"]["override_rag"] = "amber"
    edited["companies"][0]["performance"]["override_text"] = "Stabilising"
    edited["companies"][0]["performance"]["override_note"] = "Board sees early stabilisation."
    Path("output/review.yaml").write_text(
        yaml.safe_dump(edited, sort_keys=False, allow_unicode=True), encoding="utf-8"
    )

    back = read_review_yaml("output/review.yaml")
    c0 = back.companies[0]
    print("Read back OK.")
    print(f"  overridden companies: {back.overridden_companies()}")
    print(f"  performance proposed : {c0.performance.proposed_rag.value} / {c0.performance.proposed_text}")
    print(f"  performance effective: {c0.performance.effective_rag.value} / {c0.performance.effective_text} "
          f"(overridden={c0.performance.was_overridden})")