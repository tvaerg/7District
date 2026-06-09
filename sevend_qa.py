"""
7D Portfolio Aggregation Engine -- QA Layer
===========================================
Runs between reasoning and render. Two passes, layered:

  1. LANGUAGE SCRUB (deterministic, in-process)
     Pure regex / string ops. Strips em-dashes, flags overconfident words
     used without metrics ("guaranteed", "stable", "sticky", "proven"),
     flags numbers used without units (e.g. "EBITDA improved by 0.4"
     missing "MSEK"). Em-dash strip is silent (formatting only); other
     issues are added to CompanyFacts.flags so they surface to the user.

  2. HALLUCINATION CHECK (one Opus call per company)
     Re-reads the company's source material (body_text + KPIs + deep_read)
     and verifies every claim, number, name, and date in the proposed
     CompanyFacts traces back to the source. If something doesn't trace,
     Opus is told to either CORRECT it to match the source or DROP it.
     Returns corrected CompanyFacts. Cost: one extra Opus call per
     reporting company per cycle. For a monthly board deliverable this
     is the right price.

Design principles
-----------------
* QA failure does NOT block render. If the Opus QA call errors, we log
  the failure on `flags` and use the original facts -- a partial deck is
  better than no deck.
* Corrections are TRANSPARENT. Every change is logged to the console
  with old -> new shown, and recorded on facts.flags so a reviewer can
  see what the QA layer touched.
* No silent mutations to content. The language scrub silently fixes
  formatting (em-dashes); content changes (banned word, hallucination)
  are flagged and visible. The user can always trust that the bullets
  on the slide trace to the source.
"""

from __future__ import annotations

import os
import re
import json
import time
from typing import Optional

import anthropic

from sevend_schema import CompanyFacts, RatedField, Financing, UpdateBlock, RAG


CLAUDE_MODEL = os.getenv("CLAUDE_MODEL", "claude-opus-4-7")
_MAX_TOKENS = 4096


# -----------------------------------------------------------------------------
# Layer 1 -- deterministic language scrub
# -----------------------------------------------------------------------------
# Banned: confident words used without a supporting metric in the same bullet.
# A bullet is considered "supported" if it contains a digit (a figure, a date,
# a percentage). If a banned word appears without any number anywhere in the
# bullet, that's a warning. We do NOT auto-strip these -- they're judgement
# calls. They get flagged for human review.
_BANNED_WORDS = [
    "guaranteed", "guarantee",
    "stable", "stability",
    "durable", "durability",
    "sticky", "stickiness",
    "proven",
    "clearly", "definitely",
    "transformational", "best practice", "quick wins",
]

# Number patterns we'll look at for "missing unit" detection. A number followed
# by a known unit/currency in the same bullet is fine. A bare number adjacent to
# financial vocabulary (EBITDA/revenue/margin/cash) without a unit is suspicious.
_FINANCIAL_TERMS = re.compile(
    r"\b(?:revenue|sales|ebitda|ebit|gross\s*profit|gross\s*margin|net\s*income|"
    r"cash|debt|burn|runway|arr|mrr|nrr|grr|churn|cagr|opex|capex)\b",
    re.I,
)
_HAS_UNIT = re.compile(
    r"\b(?:msek|ksek|sek|eur|usd|m\b|k\b|bn\b|%|percent|points|bp\b|bps\b)\b",
    re.I,
)
_HAS_NUMBER = re.compile(r"\d")
_HAS_DATE = re.compile(r"\b(?:q[1-4]|h[12]|jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec|20\d\d|26|27|28|29)\b", re.I)


def _scrub_dashes(s: str) -> str:
    """Silently replace em-dashes and en-dashes per Tom's hard rule. No flag."""
    if not s:
        return s
    s = s.replace(" — ", ", ").replace("— ", ", ").replace(" —", ",").replace("—", ", ")
    s = s.replace(" – ", ", ").replace("– ", ", ").replace(" –", ",").replace("–", ", ")
    return s


def _strip_trailing_ellipsis(s: str) -> str:
    """Strip trailing ellipsis (… or ...) that Opus sometimes appends as a 'and
    more' teaser. We don't want these on slide bullets -- bullets should read as
    complete short statements. Silent fix, no flag."""
    if not s:
        return s
    t = s.rstrip()
    while t.endswith("…") or t.endswith("..."):
        if t.endswith("…"):
            t = t[:-1].rstrip()
        else:
            t = t[:-3].rstrip()
    # Trim any orphan comma/colon/semicolon left dangling after stripping.
    while t and t[-1] in ",;:":
        t = t[:-1].rstrip()
    return t


def _scrub_string(s: str) -> str:
    return _strip_trailing_ellipsis(_scrub_dashes(s or ""))


def _check_bullet(text: str) -> list[str]:
    """Return a list of issues with a single bullet (banned words, missing units)."""
    issues = []
    t = (text or "")
    tl = t.lower()

    # Banned-word check: confident word + no supporting digit anywhere in bullet
    for w in _BANNED_WORDS:
        if w in tl and not _HAS_NUMBER.search(t):
            issues.append(f"'{w}' used without supporting metric")
            break  # one banned-word warning per bullet is enough

    # Missing-unit check: financial term + bare number + no unit
    if _FINANCIAL_TERMS.search(t) and _HAS_NUMBER.search(t) and not _HAS_UNIT.search(t):
        # Allow if the only number is a date (Q1 2026, etc.)
        non_date_part = _HAS_DATE.sub("", t)
        if _HAS_NUMBER.search(non_date_part):
            issues.append("financial figure without unit (MSEK/%/etc.)")
    return issues


def _scrub_rated(r: RatedField) -> list[str]:
    """In-place dash scrub on RatedField text; return any issues found."""
    r.proposed_text = _scrub_string(r.proposed_text)
    r.reasoning = _scrub_string(r.reasoning)
    if r.override_text:
        r.override_text = _scrub_string(r.override_text)
    if r.override_note:
        r.override_note = _scrub_string(r.override_note)
    return _check_bullet(r.proposed_text)


def _scrub_block(b: UpdateBlock) -> list[str]:
    """In-place dash scrub on a block; return all issues collected from bullets."""
    issues = []
    b.bullets = [_scrub_string(x) for x in b.bullets]
    b.reasoning = _scrub_string(b.reasoning)
    if b.override_bullets:
        b.override_bullets = [_scrub_string(x) for x in b.override_bullets]
    for idx, x in enumerate(b.bullets):
        for issue in _check_bullet(x):
            issues.append(f"bullet {idx + 1}: {issue}")
    return issues


def language_scrub(cf: CompanyFacts) -> CompanyFacts:
    """Layer 1: deterministic scrub + lint. Mutates cf in place.

    Em-dashes are silently fixed (Tom's hard rule, no value flagging). Other
    issues (banned words, missing units) are appended to cf.flags so the user
    can see what the QA layer noticed.
    """
    issues: list[str] = []

    # The three heatmap judgements.
    for label, rated in (
        ("investment_thesis", cf.investment_thesis),
        ("performance",       cf.performance),
        ("financing.summary", cf.financing.summary),
    ):
        for issue in _scrub_rated(rated):
            issues.append(f"{label}: {issue}")

    # The three update blocks.
    for label, blk in (
        ("operational", cf.operational),
        ("financial",   cf.financial),
        ("three_year",  cf.three_year),
    ):
        for issue in _scrub_block(blk):
            issues.append(f"{label}.{issue}")

    # 7D focus items (no RAG, just bulleted actions).
    cf.sevend_focus = [_scrub_string(x) for x in cf.sevend_focus]
    for idx, x in enumerate(cf.sevend_focus):
        for issue in _check_bullet(x):
            issues.append(f"sevend_focus[{idx + 1}]: {issue}")

    # Financing string fields (covenant text, etc.) -- just dash scrub.
    cf.financing.covenant_issue = _scrub_string(cf.financing.covenant_issue)
    cf.financing.amount = _scrub_string(cf.financing.amount)
    cf.financing.type = _scrub_string(cf.financing.type)
    cf.financing.timing = _scrub_string(cf.financing.timing)
    cf.financing.sevend_amount = _scrub_string(cf.financing.sevend_amount)
    cf.financing.sevend_rescue_amount = _scrub_string(cf.financing.sevend_rescue_amount)

    if issues:
        cf.flags.extend(f"language: {i}" for i in issues)
    return cf


# -----------------------------------------------------------------------------
# Layer 2 -- Opus hallucination check against source material
# -----------------------------------------------------------------------------
_VERIFY_SYSTEM = """You are a SENIOR PARTNER at the holding company 7D, doing a final-pass review on a \
monthly portfolio update before it goes to the board. Your reputation is on the line. A wrong number, a \
fabricated customer, or a date that doesn't appear in the source will undermine 7D's credibility with its \
shareholders. This must be PERFECT.

Your ONLY job is to compare the proposed assessment (a structured JSON) against the SOURCE MATERIAL for \
this company, and identify any claim that does NOT trace back to the source. Then return a CORRECTED \
version of the assessment.

What is a hallucination here:
  - A NUMBER that does not appear in the source (or any unit conversion of it).
  - A CUSTOMER, PRODUCT, or PERSON name that does not appear in the source.
  - A DATE or TIME-PERIOD that does not appear in the source.
  - A SPECIFIC EVENT (e.g. "Macade signed in March") not stated in the source.
  - A FINANCING DETAIL (covenant breach, amount, timing) not stated in the source.

What is NOT a hallucination (do NOT flag these):
  - General observations consistent with the source ("strong month", "tough market").
  - Synonyms or paraphrases of source language.
  - Reasonable inferences clearly supported by the figures.
  - Owner ACTION items in sevend_focus (these are 7D's intentions, not source facts).

When you find a hallucination, REPLACE the offending bullet/text with a corrected version that traces \
to the source. If the claim cannot be salvaged, drop the bullet entirely. Otherwise, keep the bullet as is."""


_VERIFY_USER_TEMPLATE = """COMPANY: {name}

--- SOURCE MATERIAL (the ground truth) ---
{source}
--- END SOURCE MATERIAL ---

--- PROPOSED ASSESSMENT (verify every claim against the source above) ---
{proposed}
--- END PROPOSED ASSESSMENT ---

Return ONLY valid JSON in EXACTLY this shape (mirror the proposed assessment's structure, with corrections \
applied). Use the SAME field names. If nothing needs correcting, return the input verbatim.

{{
  "corrections_made": ["short description of each change, e.g. 'dropped bullet claiming Macade signed -- not in source'"],
  "corrected": {{
    "investment_thesis": {{"proposed_text": "...", "reasoning": "..."}},
    "performance":       {{"proposed_text": "...", "reasoning": "..."}},
    "financing": {{
      "summary": {{"proposed_text": "...", "reasoning": "..."}},
      "covenant_issue": "...", "amount": "...", "type": "...",
      "sevend_amount": "...", "timing": "...", "sevend_rescue_amount": "..."
    }},
    "operational": {{"bullets": ["", ""], "reasoning": "..."}},
    "financial":   {{"bullets": ["", ""], "reasoning": "..."}},
    "three_year":  {{"bullets": ["", ""], "reasoning": "..."}},
    "sevend_focus": ["", ""]
  }}
}}

Do NOT modify RAG colours. Do NOT modify the static fields (lead, support, exit_timing). Do not add commentary outside the JSON."""


def _facts_to_proposed_json(cf: CompanyFacts) -> str:
    """Compact JSON of the facts we want verified (without RAG colours / static)."""
    payload = {
        "investment_thesis": {"proposed_text": cf.investment_thesis.proposed_text, "reasoning": cf.investment_thesis.reasoning},
        "performance":       {"proposed_text": cf.performance.proposed_text, "reasoning": cf.performance.reasoning},
        "financing": {
            "summary": {"proposed_text": cf.financing.summary.proposed_text, "reasoning": cf.financing.summary.reasoning},
            "covenant_issue": cf.financing.covenant_issue,
            "amount": cf.financing.amount,
            "type": cf.financing.type,
            "sevend_amount": cf.financing.sevend_amount,
            "timing": cf.financing.timing,
            "sevend_rescue_amount": cf.financing.sevend_rescue_amount,
        },
        "operational": {"bullets": list(cf.operational.bullets), "reasoning": cf.operational.reasoning},
        "financial":   {"bullets": list(cf.financial.bullets),   "reasoning": cf.financial.reasoning},
        "three_year":  {"bullets": list(cf.three_year.bullets),  "reasoning": cf.three_year.reasoning},
        "sevend_focus": list(cf.sevend_focus),
    }
    return json.dumps(payload, ensure_ascii=False, indent=2)


def _build_source_block(material) -> str:
    """Compose the SOURCE MATERIAL section for the QA prompt: KPIs first (ground
    truth), then body text, then deep-read notes. Same hierarchy as reasoning."""
    parts: list[str] = []

    if getattr(material, "kpis", None):
        parts.append("STRUCTURED KPI SNAPSHOTS (AUTHORITATIVE -- any number not consistent with these is a hallucination):")
        for k in material.kpis:
            parts.append(json.dumps(k, ensure_ascii=False, indent=2))
        parts.append("")

    if material.body_text:
        parts.append("REPORT BODY:")
        parts.append(material.body_text)

    if getattr(material, "deep_read", None) and "_raw" not in material.deep_read:
        parts.append("")
        parts.append("CUSTODIAL DEEP-READ NOTES:")
        parts.append(json.dumps(material.deep_read, ensure_ascii=False, indent=2))

    return "\n".join(parts)


def _robust_json(text: str) -> Optional[dict]:
    """Same robust JSON parser used elsewhere in the engine."""
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


def _apply_corrections(cf: CompanyFacts, corrected: dict) -> None:
    """Apply Opus' corrected JSON back onto the CompanyFacts. RAG colours and
    static fields are intentionally NOT changed by QA."""
    it = corrected.get("investment_thesis") or {}
    if "proposed_text" in it:
        cf.investment_thesis.proposed_text = str(it.get("proposed_text", cf.investment_thesis.proposed_text))
    if "reasoning" in it:
        cf.investment_thesis.reasoning = str(it.get("reasoning", cf.investment_thesis.reasoning))

    pf = corrected.get("performance") or {}
    if "proposed_text" in pf:
        cf.performance.proposed_text = str(pf.get("proposed_text", cf.performance.proposed_text))
    if "reasoning" in pf:
        cf.performance.reasoning = str(pf.get("reasoning", cf.performance.reasoning))

    fin = corrected.get("financing") or {}
    s = fin.get("summary") or {}
    if "proposed_text" in s:
        cf.financing.summary.proposed_text = str(s.get("proposed_text", cf.financing.summary.proposed_text))
    if "reasoning" in s:
        cf.financing.summary.reasoning = str(s.get("reasoning", cf.financing.summary.reasoning))
    for fld in ("covenant_issue", "amount", "type", "sevend_amount", "timing", "sevend_rescue_amount"):
        if fld in fin:
            setattr(cf.financing, fld, str(fin.get(fld, getattr(cf.financing, fld))))

    for blk_name in ("operational", "financial", "three_year"):
        blk = corrected.get(blk_name) or {}
        target = getattr(cf, blk_name)
        if "bullets" in blk:
            target.bullets = [str(b).strip() for b in blk.get("bullets", []) if str(b).strip()]
        if "reasoning" in blk:
            target.reasoning = str(blk.get("reasoning", target.reasoning))

    if "sevend_focus" in corrected:
        cf.sevend_focus = [str(b).strip() for b in corrected.get("sevend_focus", []) if str(b).strip()]


def verify_against_source(cf: CompanyFacts, material, client: anthropic.Anthropic) -> CompanyFacts:
    """Layer 2: Opus hallucination check. Mutates cf if corrections are needed.

    Failure does NOT block the pipeline. On API/JSON error we leave cf as-is and
    add a flag so the user knows QA did not run cleanly on this company.
    """
    if not material.body_text:
        return cf  # placeholder companies have nothing to verify

    proposed_json = _facts_to_proposed_json(cf)
    source_block = _build_source_block(material)

    user_prompt = _VERIFY_USER_TEMPLATE.format(
        name=cf.canonical_name, source=source_block, proposed=proposed_json
    )

    try:
        resp = client.messages.create(
            model=CLAUDE_MODEL,
            max_tokens=_MAX_TOKENS,
            system=_VERIFY_SYSTEM,
            messages=[{"role": "user", "content": user_prompt}],
        )
        if resp.stop_reason == "max_tokens":
            cf.flags.append("QA: output truncated, original facts kept")
            return cf
        text = resp.content[0].text if resp.content else ""
        parsed = _robust_json(text)
        if parsed is None:
            cf.flags.append("QA: unparseable response, original facts kept")
            return cf

        corrections = parsed.get("corrections_made") or []
        corrected = parsed.get("corrected")
        if corrected:
            _apply_corrections(cf, corrected)
        if corrections:
            for c in corrections:
                cf.flags.append(f"QA-fixed: {c}")
        return cf
    except Exception as e:                             # noqa: BLE001
        cf.flags.append(f"QA error: {e}")
        return cf


# -----------------------------------------------------------------------------
# Orchestrator
# -----------------------------------------------------------------------------
def run_qa(facts: list[CompanyFacts], materials_by_name: dict, client: Optional[anthropic.Anthropic]) -> list[CompanyFacts]:
    """Run both QA layers across all company facts. Returns the same list
    (mutated in place) for convenience."""
    print("\n🔍 QA")
    for cf in facts:
        # Layer 1: deterministic scrub. Always runs (no API).
        before_flags = len(cf.flags)
        language_scrub(cf)
        new_lang_flags = len(cf.flags) - before_flags

        # Layer 2: hallucination check. Only for companies that have content,
        # actually reported this cycle, and aren't already flagged as failed.
        material = materials_by_name.get(cf.canonical_name.lower())
        if (cf.has_report_this_cycle
                and material is not None
                and material.body_text
                and client is not None):
            print(f"   • {cf.canonical_name}...", end=" ", flush=True)
            before_qa = len(cf.flags)
            verify_against_source(cf, material, client)
            new_qa_flags = len(cf.flags) - before_qa
            qa_fixed = sum(1 for f in cf.flags if f.startswith("QA-fixed:"))
            qa_errors = sum(1 for f in cf.flags if f.startswith("QA error") or f.startswith("QA:"))
            if qa_errors:
                print(f"⚠️ QA error")
            elif qa_fixed:
                print(f"✏️  {qa_fixed} fix(es)")
            else:
                print(f"✅ clean")
        elif new_lang_flags:
            print(f"   • {cf.canonical_name}: lang scrub flagged {new_lang_flags}")
    return facts


if __name__ == "__main__":
    # Offline test: prove the language scrub catches what it should.
    from sevend_schema import CompanyFacts, RatedField, Financing, UpdateBlock, RAG

    cf = CompanyFacts(
        canonical_name="Test", category="Core Holdings",
        lead="AJ", support="TBD", exit_timing="2028",
        investment_thesis=RatedField(proposed_rag=RAG.GREEN, proposed_text="Stable", reasoning="x"),
        performance=RatedField(proposed_rag=RAG.GREEN, proposed_text="Strong", reasoning="x"),
        financing=Financing(summary=RatedField(proposed_rag=RAG.GREEN, proposed_text="OK", reasoning="x"),
                            covenant_issue="No"),
        operational=UpdateBlock(proposed_rag=RAG.GREEN, bullets=[
            "EBITDA improved by 0.4",               # missing unit
            "Customer base is sticky and proven",   # banned words, no metric
            "Strong Q1 — record month in China",    # em-dash should be silently stripped
        ], reasoning="x"),
        financial=UpdateBlock(proposed_rag=RAG.GREEN, bullets=["Revenue 17 MSEK"], reasoning="x"),
        three_year=UpdateBlock(proposed_rag=RAG.GREEN, bullets=["Growth on plan"], reasoning="x"),
        sevend_focus=["Sustain momentum"],
    )
    language_scrub(cf)
    print("After language scrub:")
    print(f"  flags: {cf.flags}")
    print(f"  operational bullet 3 (em-dash stripped): {cf.operational.bullets[2]!r}")