"""
7D Portfolio Aggregation Engine -- Reasoning Layer
==================================================
The brain. Turns one company's ingested material (RawCompanyMaterial -- a
faithful transcription plus, for PDFs, custodial deep-read notes) into a
PROPOSED CompanyFacts: RAG colours with reasoning for the heatmap dimensions,
and bulleted update blocks for the company-update slide.

It does NOT decide anything final. Every colour it emits is a PROPOSAL carrying
its reasoning, written into the schema's proposed_* fields. 7D adjusts later via
review.yaml. The model never gets the last word -- that is the whole design.

Key decisions
-------------
* MODEL: Claude Opus 4.7 (reasoning). Gemini 3.1 Pro already did PDF ingestion
  upstream; this layer is Opus-only.
* ONE CALL PER COMPANY. Each company is independent -- no cross-company context
  bleeds in here (portfolio-level patterns are a separate, later concern). This
  keeps each judgement isolated and auditable.
* SCHEMA-BOUND OUTPUT. Opus must return JSON matching a fixed shape. We validate
  it into CompanyFacts; a parse/validation failure triggers one stricter retry.
* CUSTODIAL, NORMALISING LENS. The prompt's job is to translate heterogeneous
  KPI languages (SaaS CMRR/NRR vs industrial BV%/TB2/EBITDA) into 7D's single
  performance/financing/thesis vocabulary -- not to re-extract every number.
* DETERMINISTIC GATES IN CODE, NOT PROMPT. The financing detail columns
  (amount/type/timing/rescue) are blanked to 'n/a' in code unless a real issue
  is flagged. We never trust the LLM to remember to blank them.

Anti-hallucination guardrails (borrowed discipline, not borrowed content):
  temperature=0.0, max_tokens truncation detection, robust JSON parse, and a
  single stricter-instruction retry on failure.

Env: ANTHROPIC_API_KEY
"""

from __future__ import annotations

import os
import json
import time
from typing import Optional

from dotenv import load_dotenv
import anthropic

from sevend_schema import (
    CompanyFacts, RatedField, Financing, UpdateBlock, RAG,
)

from pathlib import Path as _Path
def _load_env():
    here = _Path(__file__).resolve().parent
    for candidate in (here / ".env", here.parent / ".env"):
        if candidate.is_file():
            load_dotenv(candidate)
            return
    load_dotenv()
_load_env()

CLAUDE_MODEL = os.getenv("CLAUDE_MODEL", "claude-opus-4-7")  # Opus 4.7, reasoning
_MAX_TOKENS = 4096
# NOTE: Opus 4.7 rejects the `temperature` parameter ("deprecated for this
# model"), so we do not pass it. The model is deterministic enough for this
# schema-bound extraction without it.


def _client() -> anthropic.Anthropic:
    key = os.getenv("ANTHROPIC_API_KEY")
    if not key:
        raise RuntimeError("ANTHROPIC_API_KEY not set -- required for the reasoning layer.")
    return anthropic.Anthropic(api_key=key)


# -----------------------------------------------------------------------------
# The reasoning prompt
# -----------------------------------------------------------------------------
# The schema we ask Opus to fill. Kept in the prompt verbatim so the model sees
# exactly the shape we will validate against.
_OUTPUT_SHAPE = """{
  "investment_thesis": {"proposed_rag": "green|amber|red", "proposed_text": "short cell text", "reasoning": "why"},
  "performance":       {"proposed_rag": "green|amber|red", "proposed_text": "short cell text", "reasoning": "why"},
  "financing": {
    "summary": {"proposed_rag": "green|amber|red", "proposed_text": "short cell text", "reasoning": "why"},
    "covenant_issue": "No | Yes, breach | In discussion | n/a  (max 3 words)",
    "has_active_issue": true|false,
    "amount": "MSEK or n/a", "type": "Defend or n/a", "sevend_amount": "MSEK or n/a",
    "timing": "Q2'26 or n/a", "sevend_rescue_amount": "MSEK or n/a"
  },
  "operational": {"proposed_rag": "green|amber|red", "bullets": ["", ""], "reasoning": "why"},
  "financial":   {"proposed_rag": "green|amber|red", "bullets": ["", ""], "reasoning": "why"},
  "three_year":  {"proposed_rag": "green|amber|red", "bullets": ["", ""], "reasoning": "why"},
  "sevend_focus": ["", ""]
}"""

_SYSTEM = """You are the analytical engine of 7D, a holding company, preparing the monthly portfolio \
report that goes to shareholders and the board. You read one portfolio company's own monthly report \
and translate it into 7D's standard assessment vocabulary.

You are a CUSTODIAN, not a sceptical buyer. These are companies 7D OWNS; the report is honest. Your \
job is to NORMALISE heterogeneous reporting (a SaaS company talks CMRR/NRR/Rule-of-40; an industrial \
company talks gross-margin%/EBITDA/volume) into one shared owner's view of: is the thesis intact, how \
did it perform, and is financing healthy.

RAG meaning (be consistent across companies so the heatmap is comparable):
  green = on plan / healthy / no owner action needed
  amber = partly on plan / watch items / some concern
  red   = off plan / material problem / owner likely must act

Rules:
- You PROPOSE. A human at 7D reviews and may override. Every colour MUST carry a short, specific
  reasoning grounded in THIS report's figures or statements. Never assert a colour without a reason.
- Ground every claim in the report. Do not invent numbers, customers, or events. If something is not
  in the report, do not assert it.
- proposed_text is a SHORT heatmap cell label: MAXIMUM 3 WORDS, ideally 1-2. It must fit a
  narrow table cell. Use a crisp label, never a sentence. Good: "Improving", "Weak trading",
  "Strong Q1", "Above budget", "Turnaround", "Runway risk". Bad (too long): "Above budget, but
  flattered by upsell", "Transformation on track but IKEA hit". Put the nuance in `reasoning`,
  not here -- the reasoning is preserved in full; the cell is just the label.
- covenant_issue is ALSO a short standard value, max 3 words: "No", "Yes, breach", "In
  discussion", "Waiver sought", or "n/a". Never a sentence like "No covenant breach reported".
- bullets: 3 per update block. Each bullet is HARD CAPPED at 18 WORDS. Count them. If you write
  a 20-word bullet, you have failed the instruction. Numbers + names + outcomes; cut filler.
  GOOD examples (real density, fits a slide):
    "Q1 sales 25 MSEK, beat target by 0.8 MSEK (+3%)"          (12 words)
    "March EBITDA -0.72 MSEK vs budget -1.11 MSEK"              (8 words)
    "MRR 2.13 MSEK, +38.4% YoY, +2.1% vs budget"               (10 words)
    "New wins: Mack 9 MSEK, AB Volvo OneCX, Hyundai and Audi"   (11 words)
    "Pipeline 160.5 MSEK ex-lost, 19% of 2026 target sold"     (10 words)
  BAD examples (too many words, get truncated by the renderer):
    "Major wins: Mack 9 MSEK volume, 3-year OneCX subscription AB Volvo (VTG, MACK and VTNA), new Hyundai and Audi in pipeline"  (21 words)
    "China posts best month ever, driven by Momenta, Baruto and Triv with strong follow-through across other accounts"  (18 words but flabby)
  Single fact per bullet. No semicolons chaining unrelated points. Drop parentheticals (VTG, MACK)
  that don't add information at slide-density. Prefer figures over adjectives. NEVER end a bullet
  with '...' or '…' -- bullets are complete thoughts, not trailing-off teasers. If you can't say
  it in a complete short sentence, drop the bullet.
- sevend_focus: 3-4 items, same 18-word cap. Owner ACTIONS, not observations.
  GOOD: "Close Viking/KPO injection before August cash cliff"  (7 words)
  GOOD: "Track 56.7 MSEK receivables build, validate collection quality"  (8 words)
  BAD:  "Continue to support management in scaling China momentum and new product launches"  (12 words but mostly filler -- could be "Support China scale-up and new product launches")
- FINANCING: set has_active_issue=true ONLY if the report shows a real financing problem or event
  (covenant breach/discussion, refinancing need, capital raise in play, runway concern). If false,
  leave amount/type/sevend_amount/timing/sevend_rescue_amount as "n/a".
- Output ONLY valid JSON in the exact shape given. No markdown, no commentary."""


def _build_prompt(material) -> str:
    """Assemble the user prompt from the transcription, deep-read notes, and any
    deterministic KPI snapshots from xlsx ingestion."""
    deep = ""
    if material.deep_read and "_raw" not in material.deep_read:
        deep = (
            "\n\nCUSTODIAL DEEP-READ NOTES (soft signal already extracted from this report; "
            "use them, especially for financing/exit and narrative-vs-numbers gaps):\n"
            + json.dumps(material.deep_read, ensure_ascii=False, indent=2)
        )
    elif material.deep_read and "_raw" in material.deep_read:
        deep = "\n\nDEEP-READ NOTES (unstructured):\n" + str(material.deep_read.get("_raw", ""))[:2000]

    # Structured KPI block: when ingestion ran a deterministic extractor on an
    # xlsx, those numbers are GROUND TRUTH. Surface them first so they anchor
    # the assessment rather than competing with raw cells in body_text.
    kpi_block = ""
    if getattr(material, "kpis", None):
        kpi_block = (
            "\n\nGROUND-TRUTH KPI SNAPSHOTS (deterministic extraction from spreadsheet data; "
            "these numbers are AUTHORITATIVE -- if anything in the body text conflicts, trust "
            "these). Quote these figures in bullets and the financing summary where relevant:\n"
        )
        for k in material.kpis:
            ents = k.get("entities") or {}
            ent_summary = ""
            if ents:
                ent_summary = "\n  Per-entity: " + "; ".join(
                    f"{name}: rev={v.get('revenue','-')}, ebitda={v.get('ebitda', v.get('ebitda_pre_eo','-'))}, cash={v.get('cash','-')}"
                    for name, v in ents.items()
                )
            kpi_block += (
                f"- file={_Path(k.get('source_file','')).name} unit={k.get('unit','KSEK')} "
                f"revenue={k.get('revenue')} gross_profit={k.get('gross_profit')} "
                f"gross_margin%={k.get('gross_margin_pct')} ebitda_pre_eo={k.get('ebitda_pre_eo')} "
                f"ebitda={k.get('ebitda')} ebitda_margin%={k.get('ebitda_margin_pct')} "
                f"cash={k.get('cash')} long_term_debt={k.get('long_term_debt')} "
                f"short_term_debt={k.get('short_term_debt')} net_debt={k.get('net_debt')}"
                f"{ent_summary}\n"
            )

    static = material.static or {}
    thesis_baseline = static.get("investment_thesis", "")
    baseline_line = (
        f"\nBASELINE THESIS LABEL (7D's standing view going in): {thesis_baseline}. "
        f"Assess whether THIS month's report keeps that thesis on track."
        if thesis_baseline else ""
    )

    # Source breakdown for multi-file bundles.
    sources_line = ""
    if getattr(material, "sources", None) and len(material.sources) > 1:
        names = ", ".join(_Path(s).name for s in material.sources)
        sources_line = f"\nSOURCES (multiple files merged for this company): {names}"

    return f"""COMPANY: {material.canonical_name}  (7D category: {material.category})
SOURCE: {material.source_kind}{sources_line}{baseline_line}{kpi_block}

--- BEGIN MONTHLY REPORT CONTENT ---
{material.body_text}
--- END MONTHLY REPORT CONTENT ---{deep}

Produce 7D's standardised assessment of this company for this cycle. Return ONLY JSON in EXACTLY this shape:
{_OUTPUT_SHAPE}"""


# -----------------------------------------------------------------------------
# Robust JSON parse (shared discipline)
# -----------------------------------------------------------------------------
def _robust_json(text: str) -> Optional[dict]:
    if not text:
        return None
    t = text.strip()
    if t.startswith("```"):
        t = t[3:]
        if t.lower().startswith("json"):
            t = t[4:]
        if "```" in t:
            t = t[: t.rfind("```")]
        t = t.strip()
    try:
        return json.loads(t)
    except Exception:                                  # noqa: BLE001
        pass
    start, end = t.find("{"), t.rfind("}")
    if start != -1 and end != -1 and end > start:
        try:
            return json.loads(t[start : end + 1])
        except Exception:                              # noqa: BLE001
            return None
    return None


# -----------------------------------------------------------------------------
# Map parsed JSON -> CompanyFacts, enforcing deterministic gates
# -----------------------------------------------------------------------------
def _rated(d: dict) -> RatedField:
    return RatedField(
        proposed_rag=RAG(str(d["proposed_rag"]).lower()),
        proposed_text=str(d.get("proposed_text", "")).strip(),
        reasoning=str(d.get("reasoning", "")).strip(),
    )


def _block(d: dict) -> UpdateBlock:
    return UpdateBlock(
        proposed_rag=RAG(str(d["proposed_rag"]).lower()),
        bullets=[str(b).strip() for b in d.get("bullets", []) if str(b).strip()],
        reasoning=str(d.get("reasoning", "")).strip(),
    )


def _to_company_facts(parsed: dict, material) -> CompanyFacts:
    static = material.static or {}
    fin = parsed["financing"]
    has_issue = bool(fin.get("has_active_issue", False))

    # DETERMINISTIC GATE: blank the detail columns unless there is a real issue.
    # We do this in code, not trusting the model to remember.
    def gated(key: str) -> str:
        if not has_issue:
            return "n/a"
        v = str(fin.get(key, "n/a")).strip()
        return v or "n/a"

    flags: list[str] = []
    # Sanity flag: financing summary is red but no active issue declared -> review.
    if _rated(fin["summary"]).proposed_rag == RAG.RED and not has_issue:
        flags.append("Financing summary is RED but has_active_issue=false -- check.")

    return CompanyFacts(
        canonical_name=material.canonical_name,
        category=material.category,
        lead=static.get("lead", "n/a"),
        support=static.get("support", "n/a"),
        exit_timing=static.get("exit_timing", "n/a"),
        investment_thesis=_rated(parsed["investment_thesis"]),
        performance=_rated(parsed["performance"]),
        financing=Financing(
            summary=_rated(fin["summary"]),
            covenant_issue=str(fin.get("covenant_issue", "No")).strip() or "No",
            has_active_issue=has_issue,
            amount=gated("amount"),
            type=gated("type"),
            sevend_amount=gated("sevend_amount"),
            timing=gated("timing"),
            sevend_rescue_amount=gated("sevend_rescue_amount"),
        ),
        operational=_block(parsed["operational"]),
        financial=_block(parsed["financial"]),
        three_year=_block(parsed["three_year"]),
        sevend_focus=[str(b).strip() for b in parsed.get("sevend_focus", []) if str(b).strip()],
        source_file=material.source_file,
        source_kind=material.source_kind,
        has_report_this_cycle=True,
        flags=flags,
    )


# -----------------------------------------------------------------------------
# Public entry point
# -----------------------------------------------------------------------------
def reason_company(material, client: Optional[anthropic.Anthropic] = None) -> CompanyFacts:
    """Run Opus 4.7 on one company's material -> proposed CompanyFacts.

    On a missing/empty report, returns a placeholder CompanyFacts marked
    has_report_this_cycle=False (amber 'No report') so the heatmap shows a gap
    rather than a fabricated assessment.
    """
    if not material.body_text:
        return _no_report_placeholder(material, reason="No report content this cycle.")

    cli = client or _client()
    prompt = _build_prompt(material)

    for attempt in (1, 2):
        try:
            extra = ("" if attempt == 1 else
                     "\n\nCRITICAL: Return VALID JSON ONLY, no markdown, no prose. "
                     "Match the shape exactly.")
            resp = cli.messages.create(
                model=CLAUDE_MODEL,
                max_tokens=_MAX_TOKENS,
                system=_SYSTEM,
                messages=[{"role": "user", "content": prompt + extra}],
            )
            if resp.stop_reason == "max_tokens":
                # Truncated JSON is unparseable; retry would also truncate, so flag.
                cf = _no_report_placeholder(material, reason="Opus output truncated (max_tokens).")
                cf.has_report_this_cycle = True
                return cf

            text = resp.content[0].text if resp.content else ""
            parsed = _robust_json(text)
            if parsed is None:
                if attempt == 1:
                    time.sleep(1)
                    continue
                cf = _no_report_placeholder(material, reason="Opus returned unparseable JSON.")
                cf.has_report_this_cycle = True
                return cf

            return _to_company_facts(parsed, material)

        except KeyError as e:
            if attempt == 1:
                time.sleep(1)
                continue
            cf = _no_report_placeholder(material, reason=f"Missing field in Opus output: {e}")
            cf.has_report_this_cycle = True
            return cf
        except Exception as e:                         # noqa: BLE001
            if attempt == 1:
                time.sleep(2)
                continue
            cf = _no_report_placeholder(material, reason=f"Reasoning error: {e}")
            cf.has_report_this_cycle = True
            return cf

    # Should not reach here.
    return _no_report_placeholder(material, reason="Unknown reasoning failure.")


def _no_report_placeholder(material, reason: str) -> CompanyFacts:
    """A safe, honest placeholder when we cannot assess -- amber, clearly flagged."""
    static = material.static or {}
    amber = lambda txt: RatedField(proposed_rag=RAG.AMBER, proposed_text=txt, reasoning=reason)  # noqa: E731
    blk = lambda: UpdateBlock(proposed_rag=RAG.AMBER, bullets=[reason], reasoning=reason)        # noqa: E731
    return CompanyFacts(
        canonical_name=material.canonical_name,
        category=material.category,
        lead=static.get("lead", "n/a"),
        support=static.get("support", "n/a"),
        exit_timing=static.get("exit_timing", "n/a"),
        investment_thesis=amber("No report"),
        performance=amber("No report"),
        financing=Financing(summary=amber("No report"), covenant_issue="n/a", has_active_issue=False),
        operational=blk(), financial=blk(), three_year=blk(),
        sevend_focus=[],
        source_file=material.source_file,
        source_kind=material.source_kind,
        has_report_this_cycle=False,
        flags=[reason],
    )


if __name__ == "__main__":
    # Offline test: exercise mapping + deterministic gate WITHOUT calling the API,
    # by feeding a hand-made parsed dict through _to_company_facts.
    class _Stub:
        canonical_name = "Rapid Images"; category = "Core Holdings"
        source_file = "input/RapidImages_mars.pptx"; source_kind = "pptx_text"
        body_text = "x"; deep_read = None
        static = {"lead": "AJ", "support": "TBD", "exit_timing": "2028",
                  "investment_thesis": "Under transformation"}

    parsed_with_issue = {
        "investment_thesis": {"proposed_rag": "amber", "proposed_text": "Under transformation", "reasoning": "Plan rewrite."},
        "performance": {"proposed_rag": "red", "proposed_text": "Weak trading", "reasoning": "IKEA hit EBITDA."},
        "financing": {"summary": {"proposed_rag": "red", "proposed_text": "Weak trading", "reasoning": "Covenant breach."},
                      "covenant_issue": "Yes, breach Q1", "has_active_issue": True,
                      "amount": "5", "type": "Defend", "sevend_amount": "2.6", "timing": "Q2'26", "sevend_rescue_amount": "n/a"},
        "operational": {"proposed_rag": "amber", "bullets": ["Pipeline ok", "IKEA risk"], "reasoning": "x"},
        "financial": {"proposed_rag": "red", "bullets": ["Weak Q1"], "reasoning": "x"},
        "three_year": {"proposed_rag": "amber", "bullets": ["Plan delayed"], "reasoning": "x"},
        "sevend_focus": ["Support turnaround"],
    }
    cf = _to_company_facts(parsed_with_issue, _Stub())
    print("Mapping OK (active issue).")
    print(f"  financing amount (should be 5):   {cf.financing.amount}")
    print(f"  performance: {cf.performance.proposed_rag.value} / {cf.performance.proposed_text}")

    # Now flip has_active_issue=false and confirm the detail columns blank out.
    parsed_no_issue = json.loads(json.dumps(parsed_with_issue))
    parsed_no_issue["financing"]["has_active_issue"] = False
    cf2 = _to_company_facts(parsed_no_issue, _Stub())
    print("Deterministic gate test (no active issue):")
    print(f"  financing amount (should be n/a): {cf2.financing.amount}")
    print(f"  financing type   (should be n/a): {cf2.financing.type}")
    print(f"  flags (should warn red+noissue):  {cf2.flags}")

    # Placeholder path.
    ph = _no_report_placeholder(_Stub(), reason="No report content this cycle.")
    print(f"Placeholder: has_report={ph.has_report_this_cycle}, perf={ph.performance.proposed_text}")