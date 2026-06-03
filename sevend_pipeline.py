"""
7D Portfolio Aggregation Engine -- Pipeline Orchestrator
========================================================
Wires the layers into one runnable chain and gives 7D a single command.

    Stage 1  INGEST    files -> RawCompanyMaterial   (Gemini 3.1 Pro / python-pptx)
    Stage 2  REASON    material -> proposed CompanyFacts (Opus 4.7)
    Stage 3  RENDER    PortfolioReport -> output .pptx   (7D's template)

One command, one deliverable: the deck. 7D reviews and edits the PPTX directly
-- text in PowerPoint, and RAG cells are real editable table cells so a colour
they disagree with is a click-cell -> fill-colour change. No intermediate file,
no override semantics to learn. The engine's reasoning for each colour is
preserved in the slide's speaker notes as an audit trail.

Coverage: every registry company that received no file this cycle still gets a
placeholder row (amber 'No report'), so the heatmap is always complete and gaps
are explicit.

Usage:
    python sevend_pipeline.py *.pdf *.pptx --cycle "April 2026"
    python sevend_pipeline.py "MR 03-2026.pdf" --cycle "March 2026"

Env: GEMINI_API_KEY / GOOGLE_API_KEY (PDF ingest), ANTHROPIC_API_KEY (reasoning).
"""

from __future__ import annotations

import sys
import re
import argparse
from datetime import datetime
from pathlib import Path

from sevend_ingestion import ingest_portfolio, RawCompanyMaterial
from sevend_reasoning import reason_company, _no_report_placeholder, _client
from sevend_schema import PortfolioReport
from sevend_render import render_report
from sevend_registry import REGISTRY, CATEGORY_ORDER


def _placeholder_material(rec) -> RawCompanyMaterial:
    """A no-content material for a registry company with no file this cycle."""
    return RawCompanyMaterial(
        canonical_name=rec.canonical_name, category=rec.category,
        source_file="(no file this cycle)", source_kind="none",
        static=rec.display_dict(),
    )


def bundle_materials(materials: list[RawCompanyMaterial]) -> list[RawCompanyMaterial]:
    """Merge multiple files for the SAME company into one bundle so reasoning
    sees everything for that company in a single Opus call. KPIs and source
    files accumulate; body_text is concatenated with file separators."""
    by_name: dict[str, RawCompanyMaterial] = {}
    for m in materials:
        key = m.canonical_name.lower()
        primary = by_name.get(key)
        if primary is None:
            # Initialise with this material; record it in `sources`.
            m.sources = [m.source_file]
            by_name[key] = m
            continue
        # Merge m into primary.
        primary.sources.append(m.source_file)
        if m.body_text:
            sep = f"\n\n========== ADDITIONAL FILE: {Path(m.source_file).name} ({m.source_kind}) ==========\n\n"
            primary.body_text = (primary.body_text + sep + m.body_text) if primary.body_text else m.body_text
        if m.kpis:
            primary.kpis.extend(m.kpis)
        if m.deep_read and not primary.deep_read:
            primary.deep_read = m.deep_read
        primary.warnings.extend(m.warnings)
        primary.errors.extend(m.errors)
        # Record the mixed source kinds so reasoning prompt can mention them.
        primary.meta.setdefault("source_kinds", [primary.source_kind])
        if m.source_kind not in primary.meta["source_kinds"]:
            primary.meta["source_kinds"].append(m.source_kind)
    return list(by_name.values())


def run_pipeline(files: list[str], cycle_label: str, ceo_update: str = "",
                 template_path: str = "assets/7d_template.pptx",
                 out_path: str | None = None) -> PortfolioReport:
    print("=" * 70)
    print(f"7D PORTFOLIO AGGREGATION  |  cycle: {cycle_label}")
    print("=" * 70)

    # Stage 1: ingest.
    materials = ingest_portfolio(files)
    materials = [m for m in materials if m.canonical_name != "UNMATCHED"]

    # Bundle: multiple files for the same company merge into ONE material so
    # reasoning sees the whole picture (narrative + KPIs) in one Opus call.
    materials = bundle_materials(materials)
    bundled = [(m.canonical_name, len(m.sources)) for m in materials if len(m.sources) > 1]
    if bundled:
        print("\n   📎 Bundled multi-file companies:")
        for nm, n in bundled:
            print(f"      {nm}: {n} files merged")

    covered = {m.canonical_name for m in materials}

    # Add placeholders for uncovered registry companies so the heatmap is complete.
    for rec in REGISTRY:
        if rec.canonical_name not in covered:
            materials.append(_placeholder_material(rec))

    # Stage 2: reason (one Opus call per company that has content; placeholder otherwise).
    print(f"\n🧠 REASONING (Opus 4.7) -- {len(materials)} companies")
    client = None
    try:
        client = _client()
    except RuntimeError as e:
        print(f"   ⚠️ {e}  -- companies with content cannot be assessed; emitting placeholders.")

    facts = []
    for m in materials:
        if not m.body_text:
            cf = _no_report_placeholder(m, reason="No report this cycle.")
            print(f"   ⬜ {m.canonical_name}: no report -> placeholder")
        elif client is None:
            cf = _no_report_placeholder(m, reason="Reasoning skipped: ANTHROPIC_API_KEY missing.")
            print(f"   ⚠️ {m.canonical_name}: skipped (no key)")
        else:
            print(f"   • {m.canonical_name}...", end=" ", flush=True)
            cf = reason_company(m, client=client)
            if cf.flags:
                # A company that HAD content but came back flagged means reasoning
                # failed (API error, bad JSON, truncation) -- surface it loudly so
                # a systemic failure is never mistaken for a genuine 'No report'.
                print(f"❌ FAILED -- {cf.flags[0]}")
            else:
                print(f"✅ {cf.performance.proposed_rag.value}/{cf.performance.proposed_text}")
        facts.append(cf)

    # Order by 7D category for a tidy review file.
    rank = {c: i for i, c in enumerate(CATEGORY_ORDER)}
    facts.sort(key=lambda c: rank.get(c.category, 99))

    report = PortfolioReport(cycle_label=cycle_label, ceo_update=ceo_update, companies=facts)

    # Stage 3: render the deck directly from the template.
    # If no explicit --out was given, auto-name with cycle label + timestamp so
    # each monthly run is preserved rather than overwriting the last.
    if out_path is None:
        safe_cycle = re.sub(r"[^\w]+", "_", cycle_label).strip("_") or "report"
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        out_path = f"output/7d_portfolio_report_{safe_cycle}_{stamp}.pptx"
    print(f"\n📄 RENDER -> {out_path}")
    out = render_report(report, template_path, out_path)
    print("\n" + "=" * 70)
    print(f"✅ Wrote {out}")
    print("   Review and edit the .pptx directly in PowerPoint if needed.")
    flagged = [c.canonical_name for c in facts if c.flags and c.has_report_this_cycle]
    if flagged:
        print(f"   ⚠️ Flagged for attention (see heatmap speaker notes): {', '.join(flagged)}")
    print("=" * 70)
    return report


def main():
    ap = argparse.ArgumentParser(description="7D portfolio aggregation pipeline (ingest -> reason -> render deck)")
    ap.add_argument("files", nargs="+", help="Input report files (.pdf / .pptx)")
    ap.add_argument("--cycle", required=True, help="Cycle label, e.g. 'April 2026'")
    ap.add_argument("--ceo-update", default="", help="Optional CEO update line for the heatmap slide")
    ap.add_argument("--template", default="assets/7d_template.pptx", help="Path to 7D template .pptx")
    ap.add_argument("--out", default=None,
                    help="Output deck path. If omitted, auto-names with cycle + timestamp "
                         "(e.g. output/7d_portfolio_report_April_2026_20260528_141503.pptx).")
    args = ap.parse_args()
    run_pipeline(args.files, cycle_label=args.cycle, ceo_update=args.ceo_update,
                 template_path=args.template, out_path=args.out)


if __name__ == "__main__":
    main()