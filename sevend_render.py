"""
7D Portfolio Aggregation Engine -- Renderer
===========================================
Injects a PortfolioReport into 7D's actual template .pptx and writes the final
deck. Template-injection (not redraw): we open 7D's template and overwrite the
content of the existing heatmap table cells and update-slide text, preserving
their exact branding, fonts, and layout.

Why this keeps editing fluid for 7D
------------------------------------
The heatmap is a NATIVE PowerPoint table. We set cell text and cell fill colour
through the table API, which leaves every cell fully hand-editable: if 7D
disagrees with a RAG colour, they click the cell -> shading -> pick a colour, or
just retype the text. No intermediate file, no override syntax. The engine's
reasoning for each proposed colour is written into the heatmap slide's speaker
notes as an audit trail, so the "why" is one keystroke away but never on the
printed slide.

What it fills
-------------
* Heatmap (slide 2): rows 2-13, one per company, matched by name in column 1.
    col 4  investment thesis   (RAG fill)
    col 5  performance         (RAG fill)
    col 6  financing summary   (RAG fill)
    col 7  covenant issue
    col 8-12  financing detail (amount/type/7D amount/timing/rescue) -- gated
    col 13 exit timing
* Update slides (3-6): two companies per slide, the Operational / Financial /
  Three-year blocks + 7D focus. (MVP: text injection into the existing layout.)

RAG -> fill hex (matches the template's own palette):
    green = 00B050 , amber = FFFF00 , red = FF0000

Dependencies: python-pptx.
"""

from __future__ import annotations

from pathlib import Path
from typing import Optional

from pptx import Presentation
from pptx.util import Pt
from pptx.dml.color import RGBColor

from sevend_schema import PortfolioReport, CompanyFacts, RAG


# Template-native RAG fills.
RAG_HEX = {
    RAG.GREEN: "00B050",
    RAG.AMBER: "FFFF00",
    RAG.RED:   "FF0000",
}
NA_HEX = "D5D5D5"  # grey, template's n/a tone

# Heatmap column indices (mapped from the template).
COL_NAME = 1
COL_THESIS = 4
COL_PERFORMANCE = 5
COL_FIN_SUMMARY = 6
COL_COVENANT = 7
COL_AMOUNT = 8
COL_TYPE = 9
COL_7D_AMOUNT = 10
COL_TIMING = 11
COL_RESCUE = 12
COL_EXIT = 13
FIRST_DATA_ROW = 2  # rows 0-1 are headers


# -----------------------------------------------------------------------------
# Low-level cell helpers -- preserve formatting, only change text + fill
# -----------------------------------------------------------------------------
def _set_cell_text(cell, text: str) -> None:
    """Overwrite a cell's text while preserving the run's existing formatting.

    We edit the first run in place (keeping its font/size/colour) and clear any
    extra runs, rather than rebuilding the paragraph -- this keeps the template's
    typography intact.
    """
    text = "" if text is None else str(text)
    tf = cell.text_frame
    if not tf.paragraphs:
        tf.text = text
        return
    p = tf.paragraphs[0]
    if p.runs:
        p.runs[0].text = text
        # Drop any trailing runs so we don't leave stale fragments.
        for extra in p.runs[1:]:
            extra._r.getparent().remove(extra._r)
    else:
        p.text = text
    # Remove any additional paragraphs beyond the first.
    for extra_p in tf.paragraphs[1:]:
        extra_p._p.getparent().remove(extra_p._p)


def _set_cell_fill(cell, hex_color: Optional[str]) -> None:
    if hex_color is None:
        return
    # python-pptx's cell.fill API is unreliable for table cells that already carry
    # a template fill -- the old <a:solidFill> can persist. Write the fill directly
    # into tcPr: remove any existing fill nodes, then insert a fresh solidFill.
    from pptx.oxml.ns import qn
    from pptx.oxml import parse_xml
    tcPr = cell._tc.get_or_add_tcPr()
    for tag in ("a:solidFill", "a:noFill", "a:gradFill", "a:blipFill", "a:pattFill", "a:grpFill"):
        for el in tcPr.findall(qn(tag)):
            tcPr.remove(el)
    nsmain = "http://schemas.openxmlformats.org/drawingml/2006/main"
    fill = parse_xml(
        f'<a:solidFill xmlns:a="{nsmain}"><a:srgbClr val="{hex_color}"/></a:solidFill>'
    )
    # tcPr children have an order; fill must precede any trailing elements like
    # a:lnL etc. Inserting right after the last line element or at end is safe here
    # because the template cells only carry margins + fill. Append works.
    tcPr.append(fill)


def _rag_fill(cell, rag: RAG) -> None:
    _set_cell_fill(cell, RAG_HEX.get(rag))


# -----------------------------------------------------------------------------
# Heatmap injection
# -----------------------------------------------------------------------------
def _find_heatmap_table(slide):
    for shape in slide.shapes:
        if shape.has_table:
            return shape.table
    return None


def _index_rows_by_name(table) -> dict[str, int]:
    """Map normalised company name -> row index, for rows that name a company."""
    idx = {}
    for r in range(FIRST_DATA_ROW, len(table.rows)):
        name = table.cell(r, COL_NAME).text.strip()
        if name:
            idx[name.lower()] = r
    return idx


def _inject_company_row(table, row: int, c: CompanyFacts) -> None:
    fin = c.financing

    # Text + RAG fills for the judgement columns.
    _set_cell_text(table.cell(row, COL_THESIS), c.investment_thesis.effective_text)
    _rag_fill(table.cell(row, COL_THESIS), c.investment_thesis.effective_rag)

    _set_cell_text(table.cell(row, COL_PERFORMANCE), c.performance.effective_text)
    _rag_fill(table.cell(row, COL_PERFORMANCE), c.performance.effective_rag)

    _set_cell_text(table.cell(row, COL_FIN_SUMMARY), fin.summary.effective_text)
    _rag_fill(table.cell(row, COL_FIN_SUMMARY), fin.summary.effective_rag)

    # Covenant + financing detail columns (already gated to n/a upstream).
    _set_cell_text(table.cell(row, COL_COVENANT), fin.covenant_issue)
    # Colour the covenant cell: red if there's an active issue, else neutral grey
    # (never leave the template's original green/red showing).
    _set_cell_fill(table.cell(row, COL_COVENANT),
                   RAG_HEX[RAG.RED] if fin.has_active_issue else NA_HEX)
    _set_cell_text(table.cell(row, COL_AMOUNT), fin.amount)
    _set_cell_text(table.cell(row, COL_TYPE), fin.type)
    _set_cell_text(table.cell(row, COL_7D_AMOUNT), fin.sevend_amount)
    _set_cell_text(table.cell(row, COL_TIMING), fin.timing)
    _set_cell_text(table.cell(row, COL_RESCUE), fin.sevend_rescue_amount)

    # Exit timing (static, from registry).
    _set_cell_text(table.cell(row, COL_EXIT), c.exit_timing)


def _write_heatmap_notes(slide, report: PortfolioReport) -> None:
    """Put the engine's reasoning for every proposed colour into speaker notes."""
    lines = [f"7D heatmap -- engine reasoning ({report.cycle_label})", ""]
    for c in report.companies:
        lines.append(f"{c.canonical_name}:")
        lines.append(f"  Thesis [{c.investment_thesis.effective_rag.value}]: {c.investment_thesis.reasoning}")
        lines.append(f"  Performance [{c.performance.effective_rag.value}]: {c.performance.reasoning}")
        lines.append(f"  Financing [{c.financing.summary.effective_rag.value}]: {c.financing.summary.reasoning}")
        if c.flags:
            lines.append(f"  ⚠️ FLAGS: {'; '.join(c.flags)}")
        lines.append("")
    slide.notes_slide.notes_text_frame.text = "\n".join(lines)


def render_heatmap(prs: Presentation, report: PortfolioReport) -> list[str]:
    """Fill the heatmap slide. Returns a list of company names not found in the table."""
    slide = prs.slides[1]  # slide 2 (0-indexed)
    table = _find_heatmap_table(slide)
    if table is None:
        raise RuntimeError("No table found on the heatmap slide (slide 2).")

    name_to_row = _index_rows_by_name(table)
    filled_rows: set[int] = set()
    missing = []
    for c in report.companies:
        row = name_to_row.get(c.canonical_name.lower())
        if row is None:
            missing.append(c.canonical_name)
            continue
        _inject_company_row(table, row, c)
        filled_rows.add(row)

    # Neutralise any data row we did NOT fill, so the template's original demo
    # data can never leak into a real deck. A row with a company name but no
    # matching CompanyFacts is blanked to a grey 'No report' state.
    for r in range(FIRST_DATA_ROW, len(table.rows)):
        if r in filled_rows:
            continue
        if not table.cell(r, COL_NAME).text.strip():
            continue  # structural/empty row, leave alone
        _neutralise_row(table, r)

    _write_heatmap_notes(slide, report)
    return missing


def _neutralise_row(table, row: int) -> None:
    """Blank a heatmap row that received no real data: grey judgement cells,
    'No report' text, n/a financing. Prevents stale demo data from showing."""
    for col, txt in (
        (COL_THESIS, "No report"),
        (COL_PERFORMANCE, "No report"),
        (COL_FIN_SUMMARY, "No report"),
    ):
        _set_cell_text(table.cell(row, col), txt)
        _set_cell_fill(table.cell(row, col), NA_HEX)
    for col in (COL_COVENANT, COL_AMOUNT, COL_TYPE, COL_7D_AMOUNT, COL_TIMING, COL_RESCUE):
        _set_cell_text(table.cell(row, col), "n/a")
        _set_cell_fill(table.cell(row, col), NA_HEX)


# -----------------------------------------------------------------------------
# Update slides (3-6): text injection
# -----------------------------------------------------------------------------
def _company_update_text(c: CompanyFacts) -> str:
    """Compose the update-block text for one company (MVP: consolidated text)."""
    def block(title, b):
        head = f"{title} [{b.effective_rag.value}]"
        bul = "\n".join(f"  • {x}" for x in b.effective_bullets) or "  • (no points)"
        return f"{head}\n{bul}"
    parts = [
        block("Operational update", c.operational),
        block("Financial performance", c.financial),
        block("Three-year value creation", c.three_year),
    ]
    if c.sevend_focus:
        parts.append("7D focus and contribution\n" + "\n".join(f"  • {x}" for x in c.sevend_focus))
    return "\n\n".join(parts)


# -----------------------------------------------------------------------------
# Public entry point
# -----------------------------------------------------------------------------
def render_report(report: PortfolioReport, template_path: str, out_path: str) -> str:
    """Render the full deck from the template. Returns the output path."""
    prs = Presentation(template_path)

    missing = render_heatmap(prs, report)

    # Title slide date label, if present (slide 1).
    # (Left as-is for MVP; 7D's template title is generic.)

    Path(out_path).parent.mkdir(parents=True, exist_ok=True)
    prs.save(out_path)

    if missing:
        print(f"   ⚠️ Companies not found as rows in the template heatmap: {', '.join(missing)}")
        print(f"      (Add a matching row to the template, or align registry names.)")
    return out_path


if __name__ == "__main__":
    # Build a small two-company report and render against the template, to prove
    # heatmap injection + RAG fills end to end (no API needed).
    from sevend_schema import RatedField, Financing, UpdateBlock

    def cf(name, cat, thesis, perf, fin_sum, cov, rag_t, rag_p, rag_f,
           amount="n/a", typ="n/a", s7d="n/a", timing="n/a", exit="2028", issue=False):
        return CompanyFacts(
            canonical_name=name, category=cat, lead="AJ", support="TBD", exit_timing=exit,
            investment_thesis=RatedField(proposed_rag=rag_t, proposed_text=thesis, reasoning=f"{name} thesis basis."),
            performance=RatedField(proposed_rag=rag_p, proposed_text=perf, reasoning=f"{name} perf basis."),
            financing=Financing(
                summary=RatedField(proposed_rag=rag_f, proposed_text=fin_sum, reasoning=f"{name} fin basis."),
                covenant_issue=cov, has_active_issue=issue,
                amount=amount, type=typ, sevend_amount=s7d, timing=timing),
            operational=UpdateBlock(proposed_rag=rag_p, bullets=["Op point A", "Op point B"], reasoning="op"),
            financial=UpdateBlock(proposed_rag=rag_f, bullets=["Fin point A"], reasoning="fin"),
            three_year=UpdateBlock(proposed_rag=rag_t, bullets=["3yr point A"], reasoning="3yr"),
            sevend_focus=["Focus A", "Focus B"],
        )

    report = PortfolioReport(
        cycle_label="April 2026", ceo_update="Mixed month.",
        companies=[
            cf("Inretrn", "Core Holdings", "OK", "Growing", "OK", "No",
               RAG.GREEN, RAG.GREEN, RAG.GREEN, exit="2028"),
            cf("Rapid Images", "Core Holdings", "Under transformation", "Weak trading",
               "Weak trading", "Yes, breach Q1", RAG.AMBER, RAG.RED, RAG.RED,
               amount="5", typ="Defend", s7d="2.6", timing="Q2'26", exit="2028", issue=True),
        ],
    )

    out = render_report(report, "template/7d_template.pptx", "output/7d_report_April_2026.pptx")
    print(f"Rendered: {out}")

    # Verify by reading back the injected cells.
    prs = Presentation(out)
    t = _find_heatmap_table(prs.slides[1])
    rows = _index_rows_by_name(t)
    for nm in ("inretrn", "rapid images"):
        r = rows[nm]
        print(f"  {nm}: thesis='{t.cell(r,COL_THESIS).text}' perf='{t.cell(r,COL_PERFORMANCE).text}' "
              f"fin='{t.cell(r,COL_FIN_SUMMARY).text}' amount='{t.cell(r,COL_AMOUNT).text}'")