"""
7D Portfolio Company Registry
=============================
Binds incoming report files to canonical portfolio companies and holds the
STATIC heatmap fields that do not come from the monthly report itself
(ownership category, 7D lead/support, exit timing, investment thesis baseline).

Why this exists
---------------
An input file is named e.g. "MR_03-2026.pdf" but must map to "Inretrn" and sit
under "Core Holdings" with lead=AJ, support=TBD, exit=2028. The monthly report
does NOT contain these facts -- they are 7D's own governance data. Keeping them
here (a) keeps the LLM out of guessing ownership structure, and (b) gives 7D a
single, auditable place to maintain portfolio metadata between months.

This registry is the source of truth for everything in the heatmap that is NOT
a judgement about this month's performance. The engine fills the rest.

Matching strategy (in priority order):
    1. Explicit file->company mapping (FILE_MAP) if the operator set one.
    2. Filename token match against each company's `aliases`.
    3. Fallback: fuzzy contains-match on the canonical name.
Unmatched files are surfaced loudly rather than silently dropped.
"""

from __future__ import annotations

from dataclasses import dataclass, field, asdict
from pathlib import Path
from typing import Optional
import re


# Ownership buckets, in the display order 7D uses on the heatmap.
CATEGORY_ORDER = ["Core Holdings", "Non-core holdings", "Under divestment"]


@dataclass
class CompanyRecord:
    """Static, 7D-owned metadata for one portfolio company.

    None of these fields are derived from the monthly report. They are the
    governance facts 7D maintains. The engine never writes them; it only reads
    them to populate the non-judgement columns of the heatmap.
    """
    canonical_name: str            # As it should appear on the slide, e.g. "Inretrn"
    category: str                  # One of CATEGORY_ORDER
    lead: str = "n/a"              # 7D involvement -- lead person initials
    support: str = "n/a"           # 7D involvement -- support person initials
    investment_thesis: str = ""    # Baseline thesis label, e.g. "OK", "Under transformation"
    exit_timing: str = "n/a"       # e.g. "2028", "2026/2028"
    aliases: list[str] = field(default_factory=list)  # Tokens to match filenames against

    def display_dict(self) -> dict:
        return asdict(self)


# -----------------------------------------------------------------------------
# THE REGISTRY
# -----------------------------------------------------------------------------
# Seeded from the 7D template heatmap. 7D maintains this between months.
# Values mirror the template so the MVP reproduces a recognisable output; 7D
# will edit these as the portfolio changes (divestments, new holdings, etc).
# -----------------------------------------------------------------------------
REGISTRY: list[CompanyRecord] = [
    CompanyRecord(
        canonical_name="Kvaser",
        category="Core Holdings",
        lead="NI/JA", support="AJ",
        investment_thesis="OK",
        exit_timing="Pot. Industrial Group",
        aliases=["kvaser"],
    ),
    CompanyRecord(
        canonical_name="Rapid Images",
        category="Core Holdings",
        lead="AJ", support="TBD",
        investment_thesis="Under transformation",
        exit_timing="2028",
        aliases=["rapid", "rapidimages", "rapid_images"],
    ),
    CompanyRecord(
        canonical_name="Cloud & Compute",
        category="Core Holdings",
        lead="DA", support="SGS",
        investment_thesis="Delayed",
        exit_timing="2029",
        aliases=["cloud", "compute", "cloudcompute", "cloud_compute"],
    ),
    CompanyRecord(
        canonical_name="Ocean Collective",
        category="Core Holdings",
        lead="AJ", support="TBD",
        investment_thesis="Delayed / turnaround",
        exit_timing="2028",
        aliases=["ocean", "oceancollective", "oc_bf", "ocbf", "oc", "korshags", "koster"],
    ),
    CompanyRecord(
        canonical_name="Acre",
        category="Core Holdings",
        lead="SGS", support="DA",
        investment_thesis="OK",
        exit_timing="2026/2028",
        aliases=["acre"],
    ),
    CompanyRecord(
        canonical_name="Airolit",
        category="Core Holdings",
        lead="SGS", support="DA",
        investment_thesis="OK",
        exit_timing="2026/2027",
        aliases=["airolit"],
    ),
    CompanyRecord(
        canonical_name="Inretrn",
        category="Core Holdings",
        lead="AJ", support="TBD",
        investment_thesis="OK",
        exit_timing="2028",
        aliases=["inretrn", "inretrn", "mr_", "mr-", "inretr"],
    ),
    CompanyRecord(
        canonical_name="Incipientus",
        category="Non-core holdings",
        lead="PS", support="n/a",
        investment_thesis="OK",
        exit_timing="2027",
        aliases=["incipientus"],
    ),
    CompanyRecord(
        canonical_name="Klint",
        category="Non-core holdings",
        lead="PS", support="n/a",
        investment_thesis="OK",
        exit_timing="2027",
        aliases=["klint"],
    ),
    CompanyRecord(
        canonical_name="Qamcom",
        category="Under divestment",
        lead="AJ", support="SGS",
        investment_thesis="Divest",
        exit_timing="2026",
        aliases=["qamcom"],
    ),
    CompanyRecord(
        canonical_name="Visiba",
        category="Under divestment",
        lead="AJ", support="n/a",
        investment_thesis="Divest",
        exit_timing="2027",
        aliases=["visiba"],
    ),
    CompanyRecord(
        canonical_name="Devant",
        category="Under divestment",
        lead="NI", support="AJ",
        investment_thesis="Divest",
        exit_timing="2026",
        aliases=["devant"],
    ),
]

# Optional explicit override: exact input filename (stem, lowercase) -> canonical name.
# Use this when a filename has no usable tokens. Highest matching priority.
FILE_MAP: dict[str, str] = {
    # "mr_03-2026": "Inretrn",
}


def _by_name() -> dict[str, CompanyRecord]:
    return {r.canonical_name.lower(): r for r in REGISTRY}


def match_file_to_company(file_path: str | Path) -> Optional[CompanyRecord]:
    """Resolve an input file to a CompanyRecord, or None if no confident match.

    The match is deterministic and explainable -- no LLM guessing about which
    company a file belongs to. If this returns None, the operator must add an
    alias or a FILE_MAP entry; we never guess ownership.
    """
    stem = Path(file_path).stem.lower()
    norm = re.sub(r"[^a-z0-9]+", "_", stem)
    tokens = set(t for t in norm.split("_") if t)

    # 1. Explicit operator override.
    if norm in FILE_MAP:
        return _by_name().get(FILE_MAP[norm].lower())

    # 2. Alias token match -- prefer the company with the most/longest alias hits.
    best: Optional[CompanyRecord] = None
    best_score = 0
    for rec in REGISTRY:
        score = 0
        for alias in rec.aliases:
            a = re.sub(r"[^a-z0-9]+", "_", alias.lower()).strip("_")
            if not a:
                continue
            if a in tokens:
                score += len(a) + 2          # whole-token hit, weighted by length
            elif a in norm:
                score += len(a)              # substring hit
        if score > best_score:
            best_score, best = score, rec
    if best is not None and best_score > 0:
        return best

    # 3. Canonical-name contains fallback.
    for rec in REGISTRY:
        cn = re.sub(r"[^a-z0-9]+", "_", rec.canonical_name.lower()).strip("_")
        if cn and cn in norm:
            return rec

    return None


def get_company(canonical_name: str) -> Optional[CompanyRecord]:
    return _by_name().get(canonical_name.lower())


def all_companies_in_order() -> list[CompanyRecord]:
    """Registry sorted by 7D's category display order, stable within category."""
    order = {c: i for i, c in enumerate(CATEGORY_ORDER)}
    return sorted(REGISTRY, key=lambda r: (order.get(r.category, 99),))


if __name__ == "__main__":
    # Quick self-test against the sample inputs.
    import sys
    inputs = sys.argv[1:] or [
        "input/Inretrn_MR_03-2026.pdf",
        "input/OceanCollective_2026-04.pdf",
        "input/RapidImages_mars.pptx",
        "input/Kvaser_Q1.pptx",
    ]
    print("Registry match test:")
    for f in inputs:
        rec = match_file_to_company(f)
        name = rec.canonical_name if rec else "‼️  UNMATCHED"
        cat = f"  [{rec.category}]" if rec else ""
        print(f"  {Path(f).name:40s} -> {name}{cat}")