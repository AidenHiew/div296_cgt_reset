"""Build the Adviser Edition user guide (.docx).

    python scripts/build_user_guide.py
    # writes docs/USER_GUIDE_Adviser_Edition.docx

Audience: financial advisers using the Division 296 Cost Base Reset Model
without CLASS Super as their administration software. The Adviser Edition
of the workbook omits the CLASS Import staging tab; everything else is
identical to the full model.
"""

from __future__ import annotations

from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt, RGBColor

from div296 import __version__

REPO_ROOT = Path(__file__).resolve().parents[1]
OUTPUT = REPO_ROOT / "docs" / "USER_GUIDE_Adviser_Edition.docx"


def _set_font(run, *, size=11, bold=False, color=None, italic=False):
    run.font.name = "Calibri"
    run.font.size = Pt(size)
    run.bold = bold
    run.italic = italic
    if color is not None:
        run.font.color.rgb = RGBColor(*color)


def add_title(doc, text):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    _set_font(p.add_run(text), size=22, bold=True, color=(0x1D, 0x3B, 0x34))


def add_subtitle(doc, text):
    p = doc.add_paragraph()
    _set_font(p.add_run(text), size=12, italic=True, color=(0x55, 0x55, 0x55))


def add_h1(doc, text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(14)
    p.paragraph_format.space_after = Pt(4)
    _set_font(p.add_run(text), size=16, bold=True, color=(0x1D, 0x3B, 0x34))


def add_h2(doc, text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(10)
    p.paragraph_format.space_after = Pt(2)
    _set_font(p.add_run(text), size=12, bold=True, color=(0x3F, 0x4F, 0x4A))


def add_para(doc, text):
    p = doc.add_paragraph()
    _set_font(p.add_run(text), size=11)


def add_bullets(doc, items):
    for item in items:
        p = doc.add_paragraph(style="List Bullet")
        _set_font(p.add_run(item), size=11)


def add_steps(doc, items):
    for item in items:
        p = doc.add_paragraph(style="List Number")
        _set_font(p.add_run(item), size=11)


def add_callout(doc, label, text, *, color=(0x8A, 0x6D, 0x00)):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(8)
    p.paragraph_format.space_after = Pt(8)
    _set_font(p.add_run(f"{label}  "), size=11, bold=True, color=color)
    _set_font(p.add_run(text), size=11, color=color)


def add_table(doc, headers, rows):
    table = doc.add_table(rows=1 + len(rows), cols=len(headers))
    table.style = "Light Grid Accent 1"
    for j, h in enumerate(headers):
        cell = table.rows[0].cells[j]
        cell.text = ""
        _set_font(cell.paragraphs[0].add_run(h), size=10, bold=True)
    for i, row in enumerate(rows, start=1):
        for j, val in enumerate(row):
            cell = table.rows[i].cells[j]
            cell.text = ""
            _set_font(cell.paragraphs[0].add_run(str(val)), size=10)


# ---------------------------------------------------------------------------

def build_guide() -> Document:
    doc = Document()

    # ---- Cover ------------------------------------------------------------
    add_title(doc, "Division 296 Cost Base Reset Model")
    add_subtitle(doc, f"User Guide — Adviser Edition  ·  v{__version__}")
    add_para(doc,
        "This guide explains how to use the Adviser Edition of the Division 296 "
        "Cost Base Reset Model: a self-contained Excel workbook that models the "
        "impact of the Division 296 cost-base reset election on an SMSF with one "
        "or more members above the $3 million Total Superannuation Balance (TSB) "
        "threshold. The Adviser Edition is identical to the full model except "
        "the CLASS Super import staging tab has been removed — data goes "
        "straight into the Inputs tab.")

    add_callout(doc, "Illustrative only.",
        "This model is a planning aid, not advice. Confirm every figure against "
        "the final ATO method, the fund's tax position, and a registered tax "
        "agent or licensed adviser before acting.")

    # ---- What the tool does ----------------------------------------------
    add_h1(doc, "1. What the tool does")
    add_para(doc,
        "The workbook compares two scenarios side by side for a single SMSF at "
        "the 30 June 2026 reset election date:")
    add_bullets(doc, [
        "If no reset (default) — every asset keeps its original cost base for "
        "both ordinary CGT and Division 296 calculations.",
        "If elected to reset — every CGT asset's Division 296 cost base is "
        "stepped up (or down) to its market value at 30 June 2026. Ordinary "
        "CGT continues to use the original cost base.",
    ])
    add_para(doc,
        "The reset election is one-off, all-or-nothing, and irrevocable. The "
        "model never recommends an outcome — it computes both scenarios and "
        "presents the signed difference so the adviser can make the call.")
    add_para(doc, "The model also surfaces:")
    add_bullets(doc, [
        "Fund-level Division 296 earnings, with capital-loss netting within "
        "the income year (s102-5 method, floored at zero).",
        "Per-member Division 296 tax, decomposed into the $3m–$10m band (15%) "
        "and the above-$10m band (25%).",
        "Per-asset Division 296 gain, tax, and reset impact — including the "
        "‘reset trap’ on assets currently sitting at an unrealised loss.",
        "A print-ready single-page Comparison tearsheet suitable for a client "
        "meeting.",
    ])

    # ---- Tab layout ------------------------------------------------------
    add_h1(doc, "2. The four tabs")
    add_table(doc, ["Tab", "Purpose", "User edits here?"], [
        ("Inputs",     "Member TSBs, asset register, advanced assumptions",      "Yes"),
        ("Analyser",   "Fund summary, per-member tax, per-asset detail, recon",  "No (formula-driven)"),
        ("Comparison", "One-page tearsheet — headline, per-scenario subtotals, "
                       "per-member breakdown, top-10 assets by impact",          "No"),
        ("Notes",      "Terminology, caveats, valuation log, provenance",        "Royal Assent date only"),
    ])
    add_para(doc,
        "Read across left to right: Inputs feeds Analyser; Analyser drives the "
        "Comparison tearsheet; Notes carries the disclosures. All cross-tab "
        "references are live formulas — change a figure on Inputs and every "
        "downstream cell updates immediately.")

    # ---- Quick start -----------------------------------------------------
    add_h1(doc, "3. Quick start — five steps")
    add_steps(doc, [
        "Open the workbook. Note the yellow ‘Sample data preloaded’ badge at "
        "the top of the Inputs tab. The model ships with three sample assets "
        "(P1, S1, L1) and two sample member TSBs so you can see what a "
        "completed analysis looks like before you start.",
        "On the Inputs tab, Section 1 (Members), overwrite the sample TSBs "
        "with the fund's actual member balances. The Split %, $3m–$10m band, "
        "and above-$10m band columns are auto-derived — leave them alone.",
        "In Section 2 (Asset register), overwrite the sample rows with the "
        "fund's CGT assets. The yellow sample-data badge clears automatically "
        "the moment every sample input has been replaced.",
        "Section 3 (Advanced assumptions) only needs attention if the ATO "
        "publishes a change to a threshold, rate, or indexation step. The "
        "defaults match the enacted Division 296 law as at the build date.",
        "Switch to the Comparison tab to read the headline result, then drill "
        "into the Analyser tab for per-asset detail.",
    ])

    # ---- Inputs ----------------------------------------------------------
    add_h1(doc, "4. Inputs tab — what goes where")

    add_h2(doc, "4.1 Section 1 — Members (rows 6–12)")
    add_para(doc,
        "Up to four members. For each, enter Member label (column A — "
        "free text) and TSB ($) (column B — current Total Superannuation Balance "
        "in dollars). The three derived columns to the right are read-only:")
    add_bullets(doc, [
        "Split % of fund earnings (auto) — each member's TSB ÷ total fund TSB.",
        "Proportion in $3m–$10m band (auto) — band1, taxed at 15%.",
        "Proportion above $10m (auto) — band2, taxed at 25%.",
    ])
    add_para(doc,
        "A member with zero TSB renders as a blank data row. The TSB diagnostic "
        "traffic-light banner just above the table tells you immediately "
        "whether Division 296 applies to this fund (green / amber / deep amber).")

    add_h2(doc, "4.2 Section 2 — Asset register (rows 15–65, 50 rows)")
    add_table(doc, ["Column", "Header", "What to enter"], [
        ("A", "Asset code",
         "Short identifier (e.g. P1, S1). Free text."),
        ("B", "Asset name",
         "Plain-language description (e.g. ‘Commercial property’)."),
        ("C", "Original cost base",
         "The asset's pre-reset cost base — used for ordinary CGT."),
        ("D", "Current market value (as at today)",
         "Optional context; not used in tax calcs. Useful for reviewer sanity-check."),
        ("E", "Market value at 30 Jun 2026",
         "Load-bearing input. Used as the Division 296 cost base if reset is elected."),
        ("F", "Valuation source / date",
         "Free text — e.g. ‘Independent val, 30/06/26’. Mirrored to the Notes valuation log."),
        ("G", "Projected sale proceeds",
         "Expected cash on disposal. If you're not modelling a sale, leave blank — "
         "the row will be excluded from totals and an amber banner will tell you so."),
        ("H", "Projected gain/loss",
         "Formula — locked. = proceeds − original cost base. Don't touch."),
        ("I", "Held > 12 months?",
         "Yes or No (dropdown). Drives the 1/3 CGT discount eligibility."),
    ])
    add_callout(doc, "Important.",
        "If you enter Projected sale proceeds (col G) for a row but leave "
        "Market value at 30 Jun 2026 (col E) blank, that whole row is "
        "excluded from every figure. A red banner on Inputs flags this so it "
        "can't go unnoticed.",
        color=(0xA6, 0x1B, 0x1B))

    add_h2(doc, "4.3 Section 3 — Advanced assumptions (rows 67–75)")
    add_para(doc,
        "Eight named-range cells that feed every downstream formula. The "
        "defaults match the enacted law:")
    add_bullets(doc, [
        "Tier 1 rate (15%) — applied to the $3m–$10m TSB slice.",
        "Tier 2 rate (25%) — applied to the slice above $10m.",
        "Threshold 1 ($3,000,000) and Threshold 2 ($10,000,000).",
        "CGT discount rate (1/3 = 33.333%) — s115-100 ITAA 1997.",
        "Fund CGT rate (15%) — accumulation phase.",
        "Indexation increments for both thresholds (used for forward-looking modelling).",
    ])
    add_para(doc,
        "Change these only if you're modelling a regulator update or a scenario "
        "the defaults don't capture. They are intentionally pulled out so any "
        "reviewer can see exactly what assumption produced any given number.")

    # ---- Analyser --------------------------------------------------------
    add_h1(doc, "5. Analyser tab — reading the answer")
    add_para(doc,
        "The Analyser is laid out so the headline result is at the top of the "
        "tab and detail is below.")

    add_h2(doc, "5.1 Fund summary (rows 6–13)")
    add_para(doc,
        "Side-by-side, three columns: If no reset · If elected to reset · "
        "Difference. Read the bottom row (‘Total Div 296 tax’) for the "
        "headline impact of the election. The Difference column is signed:")
    add_bullets(doc, [
        "Negative (green) — electing the reset reduces Division 296 tax.",
        "Positive (red) — electing the reset costs more Division 296 tax "
        "(the ‘reset trap’, typically when assets are sitting at an "
        "unrealised loss at 30 June 2026).",
    ])

    add_h2(doc, "5.2 Per-asset analysis (rows 15–67)")
    add_para(doc,
        "Always shows the elected-reset scenario. Twelve visible columns; the "
        "ones that matter to most readers are:")
    add_bullets(doc, [
        "K — Div 296 tax (the only authoritative per-asset tax number).",
        "L — Reset impact (signed: −ve = reset reduces this asset's Div 296 "
        "burden; +ve = the asset is a reset-trap candidate).",
        "E — Ord gross gain and G — Per-asset Ord CGT are diagnostic, "
        "greyed, and explicitly do NOT sum to the fund Ord CGT figure on the "
        "Reconciliation panel (capital losses net at the fund level, not per "
        "asset). Read the headline number on Reconciliation, not the per-asset sum.",
    ])

    add_h2(doc, "5.3 Reconciliation (rows 70–74)")
    add_para(doc,
        "Authoritative fund-level numbers: Ordinary CGT (after intra-year "
        "capital-loss netting), Division 296 tax payable (matches the headline), "
        "and any current-year net unused capital loss carried forward.")

    # ---- Comparison ------------------------------------------------------
    add_h1(doc, "6. Comparison tab — the one-page tearsheet")
    add_para(doc,
        "Print-ready landscape A4. Stable layout suitable for a client meeting "
        "or file note:")
    add_bullets(doc, [
        "Header — fund / prepared by / date / version.",
        "Members & TSB strip — each member's balance and the fund total.",
        "Headline cards — total Div 296 tax under each scenario and the net "
        "effect (signed).",
        "Per-scenario subtotals — Earnings, Ord CGT (unchanged), Div 296 tax, "
        "Total burden.",
        "Per-member breakdown — TSB and Div 296 tax for both scenarios.",
        "Per-asset detail — top 10 assets by absolute change in Div 296 tax, "
        "with an overflow note pointing at the Analyser for the full register.",
    ])

    # ---- Notes -----------------------------------------------------------
    add_h1(doc, "7. Notes tab — what's there")
    add_para(doc,
        "Read-only except the Royal Assent date cell. Carries:")
    add_bullets(doc, [
        "Terminology — definitions for every domain term used in the model.",
        "Caveats — the assumptions and exclusions you must hold in mind when "
        "reading any number (prior-year losses, pension phase, wash sale risk, "
        "transaction costs, alternative levers, etc.).",
        "Valuation log — mirrors each asset's valuation source/date and 30 "
        "Jun 2026 market value, so the file carries its own audit trail.",
        "Hidden provenance block — build version, build date, git short SHA.",
    ])

    # ---- Caveats ---------------------------------------------------------
    add_h1(doc, "8. Caveats you must hold in mind")
    add_bullets(doc, [
        "Prior-year capital losses are not modelled. The workbook nets gains "
        "and losses within the current income year but takes no brought-forward "
        "balance as an input — apply any carry-forward losses outside the model.",
        "Pension phase is not modelled. The model assumes 100% accumulation "
        "phase (fund earnings tax = 15%). For funds with pension members it "
        "overstates ordinary CGT and the reset's relative attractiveness.",
        "Reset-OFF is realised-only. In reality, Division 296 with no reset "
        "is assessed on the year-on-year movement in TSB (realised + "
        "unrealised). The Comparison tab compares realised vs realised, so it "
        "understates the no-reset Division 296 burden.",
        "Multi-member splits are TSB-proportional. A real fund determines each "
        "member's share via an actuarial certificate on time-weighted average "
        "balances. The model approximates with current-TSB share.",
        "Wash sale / Part IVA risk. Pre-30 June 2026 disposals purely for the "
        "tax outcome sit in anti-avoidance territory (TR 2008/1, Part IVA).",
        "Transaction costs (brokerage, stamp duty, illiquidity, market-timing) "
        "are not modelled.",
        "Sheet protection is tamper-evident, not tamper-proof. It exists to "
        "prevent accidental overwrites, not enforce immutability.",
    ])

    # ---- Troubleshooting -------------------------------------------------
    add_h1(doc, "9. Troubleshooting")
    add_table(doc, ["Symptom", "What it means", "What to do"], [
        ("Yellow ‘Sample data preloaded’ badge still showing",
         "At least one sample input (P1/S1/L1 codes or the seeded TSBs) hasn't "
         "been overwritten.",
         "Overwrite every sample cell on the Inputs tab with the fund's actual "
         "figures; the badge clears automatically."),
        ("Red banner: ‘Some rows have Projected sale proceeds but no Market "
         "value at 30 Jun 2026’",
         "Rows with col G filled and col E blank are excluded from every figure.",
         "Either enter the 30 Jun 2026 market value (col E), or clear the "
         "Projected sale proceeds cell (col G) for that row."),
        ("Red banner: ‘Formulas detected in the asset register’",
         "Something was pasted with a normal Ctrl+V instead of Paste-Special > "
         "Values, leaving live formulas in the register cells.",
         "Press Ctrl+Z to undo the paste, then re-paste using Paste-Special > "
         "Values."),
        ("TSB diagnostic banner is deep amber (above-$10m message)",
         "At least one member's TSB exceeds $10m, so the 25% above-$10m tier "
         "applies in addition to the 15% $3m–$10m tier.",
         "Read the Comparison tearsheet for the modelled impact — no input "
         "change required."),
        ("Per-asset Ord CGT column doesn't sum to the Reconciliation Ord CGT",
         "Correct behaviour. The per-asset column is diagnostic only — "
         "capital-loss netting happens at the fund level (s102-5), not per "
         "asset.",
         "Rely on the Reconciliation panel number for the authoritative "
         "fund Ord CGT figure."),
    ])

    # ---- Final disclosure ------------------------------------------------
    add_h1(doc, "10. About this guide")
    add_para(doc,
        f"Document built for Division 296 Cost Base Reset Model v{__version__} "
        "(Adviser Edition). The model itself carries a hidden provenance block "
        "on the Notes tab recording its build version, build date, and source "
        "commit. If you receive a workbook and need to verify it against this "
        "guide, unhide the rows immediately below the valuation log on Notes.")
    add_callout(doc, "Reminder.",
        "Illustrative only — not financial, tax or legal advice. Confirm "
        "against the final ATO method, regulations, and a registered tax "
        "agent or licensed adviser before relying on any figure.",
        color=(0x8A, 0x6D, 0x00))

    return doc


def main() -> None:
    OUTPUT.parent.mkdir(parents=True, exist_ok=True)
    doc = build_guide()
    doc.save(OUTPUT)
    print(f"Wrote {OUTPUT}")


if __name__ == "__main__":
    main()
