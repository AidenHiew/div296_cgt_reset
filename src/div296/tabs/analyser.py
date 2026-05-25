"""Analyser tab — full audit trail (v2.0.0 layout).

Layout (every formula references either an Inputs cell or a named range —
never a magic number):

    Row 1     Title
    Row 2     State strip (current scenario at a glance — new in v2)
    Row 4     Section band: "Current scenario (mirrored from Inputs)"
    Rows 5–7  Lever mirrors (read-only =Inputs!B*)

    Row 11    Section band: "Fund summary"
    Row 12    Fund Div 296 earnings (sum of positive realised gains)
    Rows 13–16 Member 1–4 attributed Div 296 tax
    Row 17    Total Div 296 tax (the headline)

    Row 19    Section band: "Per-asset analysis"
    Row 20    Column headers (A..J visible, K..L hidden helpers)
    Rows 21–70 50 data rows; column mapping (v2 row-num col inserted; per-asset
               cols shifted right by 1):
                  A #row-num             B Asset                C Proceeds
                  D Original CB          E Ordinary taxable gn  F Ordinary CGT
                  G Div 296 CB           H Div 296 adj gn       I Div 296 tax
                  J Reset impact         K helper (with reset)  L helper (without reset)
    Row 71    Totals

    Row 73    Section band: "Reconciliation"
    Row 74    Ordinary CGT payable
    Row 75    Div 296 tax payable (= headline)
    Row 76    Capital losses carried forward

Per-asset Div 296 tax (col I) is the pro-rata of the headline (locked decision):
    I{r} = IF(SUMIF(H$21:H$70,">0")=0, 0,
              MAX(0, H{r}) / SUMIF(H$21:H$70,">0") * $B$17)
"""

from __future__ import annotations

from openpyxl.comments import Comment
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.workbook import Workbook
from openpyxl.worksheet.worksheet import Worksheet

from div296.assumptions import ASSUMPTIONS
from div296.styles import (
    BODY_FONT, CENTER, FMT_CURRENCY, ROWNUM_FILL, ROWNUM_FONT,
    SECTION_BAND_FILL, SECTION_BAND_FONT, STATE_STRIP_FILL, TITLE_FONT, TRAP_FILL,
    PROC_HEADER_FILL, PROC_DATA_FILL,
    ORD_HEADER_FILL, ORD_DATA_FILL,
    DIV_HEADER_FILL, DIV_DATA_FILL,
    RESET_HEADER_FILL, RESET_DATA_FILL,
    RESET_HEADER_FONT, GROUP_HEADER_FONT,
    TOTALS_FILL, TOTALS_FONT,
)
from div296.tabs.inputs import (
    CONTROL_ROWS, MEMBERS_FIRST_DATA_ROW,
    REGISTER_FIRST_DATA_ROW,
)


SHEET = "Analyser"
INPUTS_SHEET = "'Inputs'"

# --- Layout constants (v2.0.0 — state strip + row-num col) ---
TITLE_ROW = 1
STATE_STRIP_ROW = 2                                    # NEW v2
LEVER_BAND_ROW = 4                                     # was 3
LEVER_FIRST_ROW = 5                                    # was 4
LEVER_LAST_ROW = LEVER_FIRST_ROW + len(CONTROL_ROWS) - 1

ESTIMATE_BANNER_ROW = 9                                # v2.3 — single-source estimate disclaimer
FUND_BAND_ROW = 11                                     # was 10
FUND_EARNINGS_ROW = 12                                 # was 11
MEMBER_TAX_FIRST_ROW = 13                              # was 12
MEMBER_TAX_LAST_ROW = MEMBER_TAX_FIRST_ROW + ASSUMPTIONS.member_count - 1
HEADLINE_ROW = MEMBER_TAX_LAST_ROW + 1                 # row 17 (was 16)

PERASSET_BAND_ROW = 19                                 # was 18
PERASSET_HEADER_ROW = 20                               # was 19
PERASSET_FIRST_ROW = 21                                # was 20
PERASSET_LAST_ROW = PERASSET_FIRST_ROW + ASSUMPTIONS.asset_register_rows - 1   # row 70
TOTALS_ROW = PERASSET_LAST_ROW + 1                     # row 71

TRAP_LEGEND_ROW = TOTALS_ROW + 1                       # row 72 (v2.2.0)
RECON_BAND_ROW = TOTALS_ROW + 2                        # row 73 (was 72)
RECON_ORD_CGT_ROW = RECON_BAND_ROW + 1
RECON_DIV296_ROW = RECON_BAND_ROW + 2
RECON_LOSSES_ROW = RECON_BAND_ROW + 3
RECON_LOSSES_CAPTION_ROW = RECON_LOSSES_ROW + 1        # row 77 (v2.2.0)

# Inputs↔Analyser offset (Inputs row = Analyser row - OFFSET)
ROW_OFFSET = PERASSET_FIRST_ROW - REGISTER_FIRST_DATA_ROW   # 21 - 20 = 1

# --- Column constants (v2.0.0 — row-number col inserted as col A) ---
ROWNUM_COL = 1
ASSET_COL = 2          # was 1
PROCEEDS_COL = 3       # was 2
ORIG_CB_COL = 4        # was 3
ORD_GAIN_COL = 5       # was 4
ORD_CGT_COL = 6        # was 5
DIV_CB_COL = 7         # was 6
DIV_GAIN_COL = 8       # was 7
DIV_TAX_COL = 9        # was 8
RESET_IMPACT_COL = 10  # was 9
HELPER_J_COL = 11      # was 10 (now col K in letters)
HELPER_K_COL = 12      # was 11 (now col L in letters)
LAST_VISIBLE_COL_LETTER = "J"     # was "I"; widens by 1 for row-num


def _band(ws: Worksheet, row: int, text: str, last_col_letter: str = "L") -> None:
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
    ws.merge_cells(f"A{TITLE_ROW}:L{TITLE_ROW}")

    # --- State strip (NEW v2.0.0) ---
    # One-line summary of the active scenario, anchored above the lever mirror.
    # Pulls live from named ranges + the Analyser headline cell.
    # v2.2.0: humanised — reads as a sentence rather than a log line.
    state_cell = ws.cell(row=STATE_STRIP_ROW, column=1)
    state_cell.value = (
        '="Right now you are viewing: reset "&LOWER(reset_on)&'
        '", CGT discount "&LOWER(discount_on)&'
        '", $10m tier "&LOWER(tier10_on)&'
        '".  Headline Div 296 tax in this view: "'
        '&TEXT(B' + str(HEADLINE_ROW) + ',"$#,##0")&"."'
    )
    state_cell.font = BODY_FONT
    state_cell.fill = STATE_STRIP_FILL
    state_cell.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.merge_cells(f"A{STATE_STRIP_ROW}:L{STATE_STRIP_ROW}")
    ws.row_dimensions[STATE_STRIP_ROW].height = 22

    # --- v2.2.0: Sample-data warning (row 3, between state strip and lever band) ---
    sample_detect = (
        f'AND({INPUTS_SHEET}!A{REGISTER_FIRST_DATA_ROW}="P1",'
        f'{INPUTS_SHEET}!A{REGISTER_FIRST_DATA_ROW + 1}="S1",'
        f'{INPUTS_SHEET}!A{REGISTER_FIRST_DATA_ROW + 2}="L1")'
    )
    badge = ws.cell(
        row=3, column=1,
        value=(
            f'=IF({sample_detect},'
            '"⚠  Sample data detected — figures below are illustrative only '
            'until the asset register on Inputs is replaced with the actual '
            'fund\'s holdings.","")'
        ),
    )
    badge.font = Font(name="Arial", size=10, bold=True, italic=True, color="8A6D00")
    badge.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.merge_cells("A3:L3")
    ws.row_dimensions[3].height = 20
    sample_amber_rule = FormulaRule(
        formula=[sample_detect],
        fill=PatternFill("solid", fgColor="FFF4CE"),
    )
    ws.conditional_formatting.add("A3:L3", sample_amber_rule)

    # --- Mirror-lever strip ---
    _band(ws, LEVER_BAND_ROW, "Current scenario (mirrored from Inputs — read-only)")
    for offset, (_name, (label, _row, coord, _options, _default)) in enumerate(
        CONTROL_ROWS.items()
    ):
        row = LEVER_FIRST_ROW + offset
        ws.cell(row=row, column=1, value=label).font = BODY_FONT
        mirror = ws.cell(row=row, column=2, value=f"={INPUTS_SHEET}!{coord}")
        mirror.font = BODY_FONT

    # --- v2.3: Estimate disclaimer banner (single source of truth) ---
    # Sits between the lever mirror strip and the fund summary so every figure
    # below it is read with the caveat already in mind. Avoids having to sprinkle
    # "estimated" through every row label downstream.
    est = ws.cell(
        row=ESTIMATE_BANNER_ROW, column=1,
        value=("Important assumption: the figures shown in this analysis are "
               "estimates only and are based on the information entered on the "
               "Inputs tab. Actual Division 296 outcomes may differ depending "
               "on final taxable income, asset values, member balances and "
               "other relevant adjustments."),
    )
    est.font = Font(name="Arial", size=10, italic=True, color="8A6D00")
    est.fill = PatternFill("solid", fgColor="FFF8E1")  # very pale amber
    est.alignment = Alignment(horizontal="left", vertical="center",
                              wrap_text=True, indent=1)
    est.border = Border(
        top=Side(style="thin", color="E0C77A"),
        bottom=Side(style="thin", color="E0C77A"),
    )
    ws.merge_cells(f"A{ESTIMATE_BANNER_ROW}:L{ESTIMATE_BANNER_ROW}")
    ws.row_dimensions[ESTIMATE_BANNER_ROW].height = 42

    # --- Fund summary block ---
    _band(ws, FUND_BAND_ROW, "Fund summary")

    # Fund earnings — sum of positive col-H realised Div 296 gains (was col G
    # in v1; per-asset block shifted right by 1 due to row-number col A).
    g_first, g_last = f"H{PERASSET_FIRST_ROW}", f"H{PERASSET_LAST_ROW}"
    ws.cell(row=FUND_EARNINGS_ROW, column=1,
            value="Fund Div 296 earnings").font = BODY_FONT
    ws.cell(
        row=FUND_EARNINGS_ROW, column=2,
        value=f'=SUMIF({g_first}:{g_last},">0")',
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
        "#",                                                  # NEW v2.0.0 — row number
        "Asset", "Projected sale proceeds", "Original cost base",
        "Ordinary taxable capital gain", "Ordinary CGT",
        "Div 296 cost base",
        "Div 296 adjusted taxable capital gain", "Div 296 tax",
        "Reset impact", "Helper: col G with reset", "Helper: col G without reset",
    ]
    # v2.0.0: column-group tinting on header row.
    # Col 1 (#) and Col 2 (Asset) stay dark teal (standard band).
    # Col 3 (Proceeds) = sand. Cols 4-6 = slate. Cols 7-9 = sage-teal.
    # Col 10 = gold (Reset impact). Cols 11-12 = helpers (will be hidden).
    HEADER_STYLE = {
        ROWNUM_COL:       (SECTION_BAND_FILL,  SECTION_BAND_FONT),
        ASSET_COL:        (SECTION_BAND_FILL,  SECTION_BAND_FONT),
        PROCEEDS_COL:     (PROC_HEADER_FILL,   GROUP_HEADER_FONT),
        ORIG_CB_COL:      (ORD_HEADER_FILL,    GROUP_HEADER_FONT),
        ORD_GAIN_COL:     (ORD_HEADER_FILL,    GROUP_HEADER_FONT),
        ORD_CGT_COL:      (ORD_HEADER_FILL,    GROUP_HEADER_FONT),
        DIV_CB_COL:       (DIV_HEADER_FILL,    GROUP_HEADER_FONT),
        DIV_GAIN_COL:     (DIV_HEADER_FILL,    GROUP_HEADER_FONT),
        DIV_TAX_COL:      (DIV_HEADER_FILL,    GROUP_HEADER_FONT),
        RESET_IMPACT_COL: (RESET_HEADER_FILL,  RESET_HEADER_FONT),
        HELPER_J_COL:     (SECTION_BAND_FILL,  SECTION_BAND_FONT),
        HELPER_K_COL:     (SECTION_BAND_FILL,  SECTION_BAND_FONT),
    }
    for col_idx, header in enumerate(headers, start=1):
        c = ws.cell(row=PERASSET_HEADER_ROW, column=col_idx, value=header)
        fill, font = HEADER_STYLE[col_idx]
        c.font = font
        c.fill = fill
        c.alignment = CENTER

    # Hide helper columns (shifted to K/L after row-number col A insertion).
    ws.column_dimensions["K"].hidden = True
    ws.column_dimensions["L"].hidden = True

    # Currency formats per column (v2 cols 3..12 — proceeds through helpers).
    currency_cols = [PROCEEDS_COL, ORIG_CB_COL, ORD_GAIN_COL, ORD_CGT_COL,
                     DIV_CB_COL, DIV_GAIN_COL, DIV_TAX_COL, RESET_IMPACT_COL,
                     HELPER_J_COL, HELPER_K_COL]

    # v2.0.0: per-column data fills (visible cols only — helpers stay unfilled).
    DATA_FILL = {
        PROCEEDS_COL:     PROC_DATA_FILL,
        ORIG_CB_COL:      ORD_DATA_FILL,
        ORD_GAIN_COL:     ORD_DATA_FILL,
        ORD_CGT_COL:      ORD_DATA_FILL,
        DIV_CB_COL:       DIV_DATA_FILL,
        DIV_GAIN_COL:     DIV_DATA_FILL,
        DIV_TAX_COL:      DIV_DATA_FILL,
        RESET_IMPACT_COL: RESET_DATA_FILL,
    }

    for offset in range(ASSUMPTIONS.asset_register_rows):
        a_row = PERASSET_FIRST_ROW + offset
        i_row = a_row - ROW_OFFSET

        # v2.3 Inputs column layout (Quantity dropped):
        # A code / B name / C orig CB / D MV today / E MV 30 Jun
        # F val source / G proceeds / H projected G/L / I held>12m
        code = f"{INPUTS_SHEET}!A{i_row}"
        name = f"{INPUTS_SHEET}!B{i_row}"
        orig = f"{INPUTS_SHEET}!C{i_row}"        # was D
        mv = f"{INPUTS_SHEET}!E{i_row}"          # was F
        proceeds = f"{INPUTS_SHEET}!G{i_row}"    # was H
        held = f"{INPUTS_SHEET}!I{i_row}"        # unchanged (still col I)

        # Col A — Row number (v2.0.0)
        rn = ws.cell(row=a_row, column=ROWNUM_COL, value=offset + 1)
        rn.fill = ROWNUM_FILL
        rn.font = ROWNUM_FONT
        rn.alignment = Alignment(horizontal="center", vertical="center")
        # Col B — Asset
        ws.cell(row=a_row, column=ASSET_COL,
                value=f'=IF({code}="","",{name}&" ("&{code}&")")')
        # Col C — Proceeds
        ws.cell(row=a_row, column=PROCEEDS_COL, value=f'=IF({proceeds}="","",{proceeds})')
        # Col D — Original cost base
        ws.cell(row=a_row, column=ORIG_CB_COL, value=f'=IF({orig}="","",{orig})')
        # Col E — Ordinary taxable capital gain
        ws.cell(row=a_row, column=ORD_GAIN_COL,
                value=_ord_taxable_formula(proceeds, orig, held))
        # Col F — Ordinary CGT (per-asset silo)
        ws.cell(
            row=a_row, column=ORD_CGT_COL,
            value=f'=IF({proceeds}="","",MAX(0,E{a_row})*fund_cgt_rate)',
        )
        # Col G — Div 296 cost base (moved here from col 4 for v1.5)
        ws.cell(
            row=a_row, column=DIV_CB_COL,
            value=f'=IF({proceeds}="","",IF(reset_on="ON",{mv},{orig}))',
        )
        # Col H — Div 296 adjusted gain (current scenario)
        cb_current = f'IF(reset_on="ON",{mv},{orig})'
        ws.cell(row=a_row, column=DIV_GAIN_COL,
                value=_div296_adj_formula(proceeds, cb_current, held))
        # Col I — Div 296 tax (pro-rata of headline). Headline lives at col B
        # in the fund block (which doesn't shift cols), only row → 17.
        ws.cell(
            row=a_row, column=DIV_TAX_COL,
            value=(
                f'=IF({proceeds}="","",'
                f'IF(SUMIF({g_first}:{g_last},">0")=0,0,'
                f'MAX(0,H{a_row})/SUMIF({g_first}:{g_last},">0")*$B${HEADLINE_ROW}))'
            ),
        )
        # Helper K — col H WITH reset (cost base = MV)
        ws.cell(row=a_row, column=HELPER_J_COL,
                value=_div296_adj_formula(proceeds, mv, held))
        # Helper L — col H WITHOUT reset (cost base = original)
        ws.cell(row=a_row, column=HELPER_K_COL,
                value=_div296_adj_formula(proceeds, orig, held))
        # Col J — Reset impact = K − L (helpers shifted J/K → K/L)
        ws.cell(
            row=a_row, column=RESET_IMPACT_COL,
            value=f'=IF({proceeds}="","",K{a_row}-L{a_row})',
        )

        for col_idx in currency_cols:
            cell = ws.cell(row=a_row, column=col_idx)
            cell.number_format = FMT_CURRENCY
            if col_idx in DATA_FILL:
                cell.fill = DATA_FILL[col_idx]

    # --- Totals row ---
    # v2.0.0: distinct totals styling — medium teal fill, dark-teal bold text,
    # double top border to visually detach from the per-asset rows.
    DOUBLE_TOP = Border(top=Side(style="double", color="1D3B34"))

    # v2.2.0: "Total" word in col A (row-num col) instead of the Σ glyph;
    # "TOTAL" label in col B.
    sigma_cell = ws.cell(row=TOTALS_ROW, column=ROWNUM_COL, value="Total")
    sigma_cell.fill = TOTALS_FILL
    sigma_cell.font = TOTALS_FONT
    sigma_cell.alignment = Alignment(horizontal="center", vertical="center")
    sigma_cell.border = DOUBLE_TOP

    totals_label = ws.cell(row=TOTALS_ROW, column=ASSET_COL, value="TOTAL")
    totals_label.font = TOTALS_FONT
    totals_label.fill = TOTALS_FILL
    totals_label.border = DOUBLE_TOP

    # Totals on Proceeds (C), Ord CGT (F), Div 296 adj gain (H), Div 296 tax (I).
    for col_letter, col_const in (
        ("C", PROCEEDS_COL),         # Proceeds (was B)
        ("F", ORD_CGT_COL),          # Ord CGT (was E)
        ("H", DIV_GAIN_COL),         # Div 296 adj gain (was G)
        ("I", DIV_TAX_COL),          # Div 296 tax (was H)
    ):
        rng = f"{col_letter}{PERASSET_FIRST_ROW}:{col_letter}{PERASSET_LAST_ROW}"
        cell = ws.cell(row=TOTALS_ROW, column=col_const, value=f"=SUM({rng})")
        cell.number_format = FMT_CURRENCY
        cell.font = TOTALS_FONT
        cell.fill = TOTALS_FILL
        cell.border = DOUBLE_TOP

    # Empty totals-row cells (cols not summed) still get the band fill + border
    # so the row reads as one continuous stripe.
    for col_const in (ORIG_CB_COL, ORD_GAIN_COL, DIV_CB_COL, RESET_IMPACT_COL):
        c = ws.cell(row=TOTALS_ROW, column=col_const)
        c.fill = TOTALS_FILL
        c.border = DOUBLE_TOP

    # --- v2.2.0: Trap-row legend (one-line key for the red row CF below) ---
    legend = ws.cell(
        row=TRAP_LEGEND_ROW, column=1,
        value=("Red row = reset would create Div 296 tax on an asset currently "
               "in an unrealised-loss position (the 'reset trap'). "
               "Gold column = how much the reset election shifts each asset's "
               "Div 296 gain."),
    )
    legend.font = Font(name="Arial", size=9, italic=True, color="666666")
    legend.alignment = Alignment(horizontal="left", vertical="center",
                                 wrap_text=True, indent=1)
    # Small red swatch in the row-number cell as visual cue.
    swatch = ws.cell(row=TRAP_LEGEND_ROW, column=1)
    swatch.fill = PatternFill("solid", fgColor="FBE9E9")
    ws.merge_cells(f"A{TRAP_LEGEND_ROW}:L{TRAP_LEGEND_ROW}")
    ws.row_dimensions[TRAP_LEGEND_ROW].height = 20

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
        orig_i = f"{INPUTS_SHEET}!C{i_row}"   # v2.3: was D
        proc_i = f"{INPUTS_SHEET}!G{i_row}"   # v2.3: was H
        cf_terms.append(f'IF({proc_i}="",0,MAX(0,{orig_i}-{proc_i}))')
    ws.cell(
        row=RECON_LOSSES_ROW, column=2,
        value="=" + "+".join(cf_terms),
    ).number_format = FMT_CURRENCY

    # v2.2.0: Plain-English caption under the CF losses row.
    caption = ws.cell(
        row=RECON_LOSSES_CAPTION_ROW, column=1,
        value=("These losses can be applied against future realised capital "
               "gains within the fund — they have no effect on Div 296."),
    )
    caption.font = Font(name="Arial", size=9, italic=True, color="666666")
    caption.alignment = Alignment(horizontal="left", vertical="center",
                                  wrap_text=True, indent=1)
    ws.merge_cells(f"A{RECON_LOSSES_CAPTION_ROW}:L{RECON_LOSSES_CAPTION_ROW}")
    ws.row_dimensions[RECON_LOSSES_CAPTION_ROW].height = 20

    # --- Trap shading: row red when ord raw gain < 0 AND Div 296 adj gain > 0 ---
    rng = (
        f"A{PERASSET_FIRST_ROW}:{LAST_VISIBLE_COL_LETTER}{PERASSET_LAST_ROW}"
    )
    trap_rule = FormulaRule(
        formula=[
            f"AND($B{PERASSET_FIRST_ROW}<>\"\","
            f"($C{PERASSET_FIRST_ROW}-$D{PERASSET_FIRST_ROW})<0,"
            f"$H{PERASSET_FIRST_ROW}>0)"
        ],
        fill=TRAP_FILL,
    )
    ws.conditional_formatting.add(rng, trap_rule)

    # --- Column widths (v2: row-num col A=5, then existing widths) ---
    widths = [5, 32, 14, 16, 16, 22, 14, 24, 14, 14, 14, 14]
    for col_idx, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(col_idx)].width = w
    # v2.0.0: free scroll (no freeze panes). Header row repeats on every
    # printed page via print_title_rows below.
    ws.print_title_rows = f"{PERASSET_HEADER_ROW}:{PERASSET_HEADER_ROW}"

    # --- Print header watermark (compliance signal on every printed page) ---
    ws.oddHeader.center.text = "ILLUSTRATIVE — NOT ADVICE"
    ws.oddHeader.center.size = 28
    ws.oddHeader.center.color = "CCCCCC"

    # --- Sheet protection (read-only; recalc + column resize allowed) ---
    ws.protection.sheet = True
    ws.protection.formatColumns = False
    ws.protection.formatRows = False

    return ws
