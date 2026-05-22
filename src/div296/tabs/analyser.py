"""Analyser tab — full 9-column audit trail per spec §5.

Layout (every formula references either an Inputs cell or a named range —
never a magic number):

    Row 1     Title
    Row 3     Section band: "Current scenario (mirrored from Inputs)"
    Rows 4–8  Lever mirrors (read-only =Inputs!B*)

    Row 10    Section band: "Fund summary"
    Row 11    Fund Div 296 earnings (auto / manual switch)
    Rows 12–15 Member 1–4 attributed Div 296 tax
    Row 16    Total Div 296 tax (the headline)

    Row 18    Section band: "Per-asset analysis"
    Row 19    Column headers (A..I visible, J..K hidden helpers)
    Rows 20–69 50 data rows; column mapping:
                  A Asset           B Proceeds              C Original CB
                  D Div 296 CB      E Ordinary taxable gn   F Ordinary CGT
                  G Div 296 adj gn  H Div 296 tax           I Reset impact
                  J helper (with reset col 7)  K helper (without reset col 7)
    Row 70    Totals

    Row 72    Section band: "Reconciliation"
    Row 73    Ordinary CGT payable
    Row 74    Div 296 tax payable (= headline)
    Row 75    Capital losses carried forward

Per-asset Div 296 tax (col 8) is the pro-rata of the headline (locked decision):
    H{r} = IF(SUMIF(G$20:G$69,">0")=0, 0,
              MAX(0, G{r}) / SUMIF(G$20:G$69,">0") * $B$16)
"""

from __future__ import annotations

from openpyxl.formatting.rule import FormulaRule
from openpyxl.utils import get_column_letter
from openpyxl.workbook import Workbook
from openpyxl.worksheet.worksheet import Worksheet

from div296.assumptions import ASSUMPTIONS
from div296.styles import (
    BODY_FONT, CENTER, FMT_CURRENCY, FMT_PERCENT,
    SECTION_BAND_FILL, SECTION_BAND_FONT, TITLE_FONT, TRAP_FILL,
)
from div296.tabs.inputs import (
    CONTROL_ROWS, MEMBERS_FIRST_DATA_ROW,
    REGISTER_FIRST_DATA_ROW,
)


SHEET = "Analyser"
INPUTS_SHEET = "'Inputs'"

# --- Layout constants ---
TITLE_ROW = 1
LEVER_BAND_ROW = 3
LEVER_FIRST_ROW = 4
LEVER_LAST_ROW = LEVER_FIRST_ROW + len(CONTROL_ROWS) - 1

FUND_BAND_ROW = 10
FUND_EARNINGS_ROW = 11
MEMBER_TAX_FIRST_ROW = 12
MEMBER_TAX_LAST_ROW = MEMBER_TAX_FIRST_ROW + ASSUMPTIONS.member_count - 1
HEADLINE_ROW = MEMBER_TAX_LAST_ROW + 1                # row 16

PERASSET_BAND_ROW = 18
PERASSET_HEADER_ROW = 19
PERASSET_FIRST_ROW = 20
PERASSET_LAST_ROW = PERASSET_FIRST_ROW + ASSUMPTIONS.asset_register_rows - 1
TOTALS_ROW = PERASSET_LAST_ROW + 1

RECON_BAND_ROW = TOTALS_ROW + 2
RECON_ORD_CGT_ROW = RECON_BAND_ROW + 1
RECON_DIV296_ROW = RECON_BAND_ROW + 2
RECON_LOSSES_ROW = RECON_BAND_ROW + 3

# Inputs↔Analyser offset (Inputs row = Analyser row - OFFSET)
ROW_OFFSET = PERASSET_FIRST_ROW - REGISTER_FIRST_DATA_ROW   # 20 - 13 = 7


def _band(ws: Worksheet, row: int, text: str, last_col_letter: str = "K") -> None:
    ws.cell(row=row, column=1, value=text).font = SECTION_BAND_FONT
    ws.merge_cells(f"A{row}:{last_col_letter}{row}")
    for col_idx in range(1, ord(last_col_letter) - ord("A") + 2):
        ws.cell(row=row, column=col_idx).fill = SECTION_BAND_FILL


def _ord_taxable_formula(proceeds: str, orig: str, held: str) -> str:
    """Discount applied to gains only; losses kept full."""
    raw = f"({proceeds}-{orig})"
    return (
        f'=IF({proceeds}="","",'
        f'IF({raw}<=0,{raw},'
        f'IF(AND({held}="Yes",discount_on="ON"),{raw}*(1-discount_rate),{raw})))'
    )


def _div296_adj_formula(proceeds: str, cost_base_expr: str, held: str) -> str:
    raw = f"({proceeds}-{cost_base_expr})"
    return (
        f'=IF({proceeds}="","",'
        f'IF({raw}<=0,{raw},'
        f'IF(AND({held}="Yes",discount_on="ON"),{raw}*(1-discount_rate),{raw})))'
    )


def _member_tax_formula(member_inputs_row: int) -> str:
    """Per-member Div 296 tax. Reads TSB/split/auto-proportion/override from Inputs."""
    tsb = f"{INPUTS_SHEET}!B{member_inputs_row}"
    split = f"{INPUTS_SHEET}!C{member_inputs_row}"
    auto_p = f"{INPUTS_SHEET}!D{member_inputs_row}"
    override = f"{INPUTS_SHEET}!E{member_inputs_row}"
    earnings_m = f"$B${FUND_EARNINGS_ROW}*{split}"
    effective_p = f'IF({override}="",{auto_p},{override})'
    # Two-tier formula (used when tier10_on = "ON")
    band1 = f'MAX(0,MIN({tsb},threshold_2)-threshold_1)/{tsb}'
    band2 = f'MAX(0,{tsb}-threshold_2)/{tsb}'
    tier_on_expr = (
        f'{earnings_m}*{band1}*rate_tier1 + {earnings_m}*{band2}*rate_tier2'
    )
    tier_off_expr = f'{earnings_m}*{effective_p}*rate_tier1'
    return (
        f'=IF(OR({tsb}="",{split}="",{tsb}<=0,{split}<=0,$B${FUND_EARNINGS_ROW}<=0),0,'
        f'IF(tier10_on="ON",{tier_on_expr},{tier_off_expr}))'
    )


def build(wb: Workbook) -> Worksheet:
    ws = wb.create_sheet(SHEET)

    # --- Title ---
    ws.cell(row=TITLE_ROW, column=1,
            value="Division 296 Cost Base Reset Model — Analyser").font = TITLE_FONT
    ws.merge_cells(f"A{TITLE_ROW}:K{TITLE_ROW}")

    # --- Mirror-lever strip ---
    _band(ws, LEVER_BAND_ROW, "Current scenario (mirrored from Inputs — read-only)")
    for offset, (_name, (label, _row, coord, _options, _default)) in enumerate(
        CONTROL_ROWS.items()
    ):
        row = LEVER_FIRST_ROW + offset
        ws.cell(row=row, column=1, value=label).font = BODY_FONT
        mirror = ws.cell(row=row, column=2, value=f"={INPUTS_SHEET}!{coord}")
        mirror.font = BODY_FONT
        if "earnings" in label.lower() and "manual" in label.lower():
            mirror.number_format = FMT_CURRENCY

    # --- Fund summary block ---
    _band(ws, FUND_BAND_ROW, "Fund summary")

    # Fund earnings — auto (sum of positive col 7) or manual override.
    g_first, g_last = f"G{PERASSET_FIRST_ROW}", f"G{PERASSET_LAST_ROW}"
    ws.cell(row=FUND_EARNINGS_ROW, column=1,
            value="Fund Div 296 earnings (auto / manual)").font = BODY_FONT
    ws.cell(
        row=FUND_EARNINGS_ROW, column=2,
        value=f'=IF(earnings_source="Manual",manual_earnings,SUMIF({g_first}:{g_last},">0"))',
    ).number_format = FMT_CURRENCY

    # Per-member attributed Div 296 tax.
    for i in range(ASSUMPTIONS.member_count):
        row = MEMBER_TAX_FIRST_ROW + i
        inputs_row = MEMBERS_FIRST_DATA_ROW + i
        ws.cell(row=row, column=1, value=f"Member {i+1} Div 296 tax").font = BODY_FONT
        cell = ws.cell(row=row, column=2, value=_member_tax_formula(inputs_row))
        cell.number_format = FMT_CURRENCY

    # Headline.
    ws.cell(row=HEADLINE_ROW, column=1,
            value="Total Div 296 tax (headline)").font = BODY_FONT
    ws.cell(
        row=HEADLINE_ROW, column=2,
        value=f"=SUM(B{MEMBER_TAX_FIRST_ROW}:B{MEMBER_TAX_LAST_ROW})",
    ).number_format = FMT_CURRENCY

    # --- Per-asset analysis ---
    _band(ws, PERASSET_BAND_ROW, "Per-asset analysis (50 rows)")
    headers = [
        "Asset", "Proceeds", "Original cost base", "Div 296 cost base",
        "Ordinary taxable capital gain", "Ordinary CGT",
        "Div 296 adjusted taxable capital gain", "Div 296 tax",
        "Reset impact", "Helper: col 7 with reset", "Helper: col 7 without reset",
    ]
    for col_idx, header in enumerate(headers, start=1):
        c = ws.cell(row=PERASSET_HEADER_ROW, column=col_idx, value=header)
        c.font = SECTION_BAND_FONT
        c.fill = SECTION_BAND_FILL
        c.alignment = CENTER

    # Hide helper columns.
    ws.column_dimensions["J"].hidden = True
    ws.column_dimensions["K"].hidden = True

    # Currency formats per column.
    currency_cols = [2, 3, 4, 5, 6, 7, 8, 9, 10, 11]

    for offset in range(ASSUMPTIONS.asset_register_rows):
        a_row = PERASSET_FIRST_ROW + offset
        i_row = a_row - ROW_OFFSET

        code = f"{INPUTS_SHEET}!A{i_row}"
        name = f"{INPUTS_SHEET}!B{i_row}"
        orig = f"{INPUTS_SHEET}!D{i_row}"
        mv = f"{INPUTS_SHEET}!F{i_row}"
        proceeds = f"{INPUTS_SHEET}!H{i_row}"
        held = f"{INPUTS_SHEET}!I{i_row}"

        # Col 1 — Asset
        ws.cell(row=a_row, column=1,
                value=f'=IF({code}="","",{name}&" ("&{code}&")")')
        # Col 2 — Proceeds
        ws.cell(row=a_row, column=2, value=f'=IF({proceeds}="","",{proceeds})')
        # Col 3 — Original cost base
        ws.cell(row=a_row, column=3, value=f'=IF({orig}="","",{orig})')
        # Col 4 — Div 296 cost base
        ws.cell(
            row=a_row, column=4,
            value=f'=IF({proceeds}="","",IF(reset_on="ON",{mv},{orig}))',
        )
        # Col 5 — Ordinary taxable capital gain
        ws.cell(row=a_row, column=5,
                value=_ord_taxable_formula(proceeds, orig, held))
        # Col 6 — Ordinary CGT (per-asset silo)
        ws.cell(
            row=a_row, column=6,
            value=f'=IF({proceeds}="","",MAX(0,E{a_row})*fund_cgt_rate)',
        )
        # Col 7 — Div 296 adjusted gain (current scenario)
        cb_current = f'IF(reset_on="ON",{mv},{orig})'
        ws.cell(row=a_row, column=7,
                value=_div296_adj_formula(proceeds, cb_current, held))
        # Col 8 — Div 296 tax (pro-rata of headline)
        ws.cell(
            row=a_row, column=8,
            value=(
                f'=IF({proceeds}="","",'
                f'IF(SUMIF({g_first}:{g_last},">0")=0,0,'
                f'MAX(0,G{a_row})/SUMIF({g_first}:{g_last},">0")*$B${HEADLINE_ROW}))'
            ),
        )
        # Helper J — col 7 WITH reset (cost base = MV)
        ws.cell(row=a_row, column=10,
                value=_div296_adj_formula(proceeds, mv, held))
        # Helper K — col 7 WITHOUT reset (cost base = original)
        ws.cell(row=a_row, column=11,
                value=_div296_adj_formula(proceeds, orig, held))
        # Col 9 — Reset impact = J − K
        ws.cell(
            row=a_row, column=9,
            value=f'=IF({proceeds}="","",J{a_row}-K{a_row})',
        )

        for col_idx in currency_cols:
            ws.cell(row=a_row, column=col_idx).number_format = FMT_CURRENCY

    # --- Totals row ---
    ws.cell(row=TOTALS_ROW, column=1, value="TOTALS").font = SECTION_BAND_FONT
    ws.cell(row=TOTALS_ROW, column=1).fill = SECTION_BAND_FILL
    for col_idx in (2, 6, 7, 8):
        col_letter = get_column_letter(col_idx)
        rng = f"{col_letter}{PERASSET_FIRST_ROW}:{col_letter}{PERASSET_LAST_ROW}"
        cell = ws.cell(row=TOTALS_ROW, column=col_idx, value=f"=SUM({rng})")
        cell.number_format = FMT_CURRENCY
        cell.font = SECTION_BAND_FONT
        cell.fill = SECTION_BAND_FILL

    # --- Reconciliation panel ---
    _band(ws, RECON_BAND_ROW, "Reconciliation")
    ws.cell(row=RECON_ORD_CGT_ROW, column=1,
            value="Ordinary CGT payable").font = BODY_FONT
    ws.cell(
        row=RECON_ORD_CGT_ROW, column=2,
        value=f"=SUM(F{PERASSET_FIRST_ROW}:F{PERASSET_LAST_ROW})",
    ).number_format = FMT_CURRENCY

    ws.cell(row=RECON_DIV296_ROW, column=1,
            value="Div 296 tax payable (headline)").font = BODY_FONT
    ws.cell(row=RECON_DIV296_ROW, column=2,
            value=f"=B{HEADLINE_ROW}").number_format = FMT_CURRENCY

    ws.cell(row=RECON_LOSSES_ROW, column=1,
            value="Capital losses carried forward").font = BODY_FONT
    # CF losses = Σ MAX(0, -raw_ord_gain) = Σ MAX(0, original − proceeds)
    cf_terms = []
    for offset in range(ASSUMPTIONS.asset_register_rows):
        i_row = REGISTER_FIRST_DATA_ROW + offset
        orig_i = f"{INPUTS_SHEET}!D{i_row}"
        proc_i = f"{INPUTS_SHEET}!H{i_row}"
        cf_terms.append(f'IF({proc_i}="",0,MAX(0,{orig_i}-{proc_i}))')
    ws.cell(
        row=RECON_LOSSES_ROW, column=2,
        value="=" + "+".join(cf_terms),
    ).number_format = FMT_CURRENCY

    # --- Trap shading: row red when ord raw gain < 0 AND col 7 > 0 ---
    rng = (
        f"A{PERASSET_FIRST_ROW}:I{PERASSET_LAST_ROW}"
    )
    trap_rule = FormulaRule(
        formula=[
            f"AND($A{PERASSET_FIRST_ROW}<>\"\","
            f"($B{PERASSET_FIRST_ROW}-$C{PERASSET_FIRST_ROW})<0,"
            f"$G{PERASSET_FIRST_ROW}>0)"
        ],
        fill=TRAP_FILL,
    )
    ws.conditional_formatting.add(rng, trap_rule)

    # --- Column widths + freeze ---
    widths = [32, 14, 16, 16, 22, 14, 24, 14, 14, 14, 14]
    for col_idx, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(col_idx)].width = w
    ws.freeze_panes = f"A{PERASSET_HEADER_ROW + 1}"

    return ws
