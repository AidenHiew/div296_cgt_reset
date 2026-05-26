"""Analyser tab — v3.0 layout.

Restructured in v3.0 around two design goals:

1. **See the answer immediately.** Fund summary at the top shows BOTH scenarios
   side-by-side (`If no reset (default)` / `If elected to reset` / `Difference`)
   with per-member tax breakdown. No tab navigation required.
2. **Single source of truth.** Per-member tax formulas read `band1`/`band2`
   straight from Inputs col D/E — what the user sees on Inputs IS what the
   calc uses.

The control panel removal (decision 1) means there are no toggles to mirror
or narrate. Row 2 carries a read-only "Parameters in effect" strip pulled
live from named ranges.

Layout (every formula references either an Inputs cell or a named range —
never a magic number):

    Row 1     Title
    Row 2     Parameters in effect (read-only, sourced from named ranges)
    Row 3     Sample-data warning (conditional)
    Row 4     Estimate disclaimer banner (always visible)

    Row 6     Section band: "Fund summary"
    Row 7     Header row: blank | If no reset (default) | If elected | Diff
    Row 8     Fund Div 296 earnings    (3 scenario columns + diff)
    Rows 9-12 Member 1-4 Div 296 tax   (3 scenario columns + diff)
    Row 13    Total Div 296 tax (headline)  (3 scenario columns + diff)

    Row 15    Section band: "Per-asset analysis (elected-reset scenario)"
    Row 16    Column headers (A..J visible, K..L hidden helpers)
    Rows 17-66 50 data rows. Always shows the elected-reset cost base.
                  A #row-num             B Asset                C Proceeds
                  D Original CB          E Ordinary taxable gn  F Ordinary CGT
                  G Div 296 CB           H Div 296 adj gn       I Div 296 tax
                  J Reset impact         K helper (with reset)  L helper (without reset)
    Row 67    Totals

    Row 69    Section band: "Reconciliation"
    Row 70    Ordinary CGT payable
    Row 71    Div 296 tax payable (= elected-reset headline)
    Row 72    Capital losses carried forward
    Row 73    Plain-English caption

Per-asset Div 296 tax (col I) is the pro-rata of the elected-reset headline
(D{HEADLINE_ROW}):
    I{r} = IF(SUMIF(H17:H66,">0")=0, 0,
              MAX(0, H{r}) / SUMIF(H17:H66,">0") * $D$13)

Hidden helper columns K, L compute both reset scenarios unconditionally
(they do not reference any toggle), so the Reset Impact column J = K - L
remains scenario-agnostic.
"""

from __future__ import annotations

from openpyxl.comments import Comment
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.workbook import Workbook
from openpyxl.worksheet.worksheet import Worksheet

from div296._formulas import per_member_div296_tax_formula
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
    MEMBERS_FIRST_DATA_ROW,
    REGISTER_FIRST_DATA_ROW,
)


SHEET = "Analyser"
INPUTS_SHEET = "'Inputs'"

# --- Layout constants (v3.0 — control panel removed; side-by-side fund summary) ---
TITLE_ROW = 1
STATE_STRIP_ROW = 2                            # "Parameters in effect" line
SAMPLE_WARN_ROW = 3                            # conditional sample-data warning
ESTIMATE_BANNER_ROW = 4                        # was 9; moved up since lever rows gone

FUND_BAND_ROW = 6                              # was 11
FUND_HEADER_ROW = 7                            # NEW — scenario column headers
FUND_EARNINGS_ROW = 8                          # was 12
MEMBER_TAX_FIRST_ROW = 9                       # was 13
MEMBER_TAX_LAST_ROW = MEMBER_TAX_FIRST_ROW + ASSUMPTIONS.member_count - 1   # row 12
HEADLINE_ROW = MEMBER_TAX_LAST_ROW + 1         # row 13 (was 17)

PERASSET_BAND_ROW = 15                         # was 19
PERASSET_HEADER_ROW = 16                       # was 20
PERASSET_FIRST_ROW = 17                        # was 21
PERASSET_LAST_ROW = PERASSET_FIRST_ROW + ASSUMPTIONS.asset_register_rows - 1   # row 66
TOTALS_ROW = PERASSET_LAST_ROW + 1             # row 67

TRAP_LEGEND_ROW = TOTALS_ROW + 1               # row 68
RECON_BAND_ROW = TOTALS_ROW + 2                # row 69
RECON_ORD_CGT_ROW = RECON_BAND_ROW + 1         # row 70
RECON_DIV296_ROW = RECON_BAND_ROW + 2          # row 71
RECON_LOSSES_ROW = RECON_BAND_ROW + 3          # row 72
RECON_LOSSES_CAPTION_ROW = RECON_LOSSES_ROW + 1   # row 73

# Inputs↔Analyser offset (Inputs row = Analyser row - OFFSET)
ROW_OFFSET = PERASSET_FIRST_ROW - REGISTER_FIRST_DATA_ROW   # 17 - 16 = 1

# --- Fund summary column layout (v3.0 — side-by-side scenarios) ---
# A & B merged for label; C/D/E for scenario data.
FUND_LABEL_COL_RANGE = ("A", "B")              # merged for labels
FUND_NORESET_COL = 3                           # col C
FUND_ELECTED_COL = 4                           # col D
FUND_DIFF_COL = 5                              # col E

# --- Per-asset column constants (unchanged from v2.6) ---
ROWNUM_COL = 1
ASSET_COL = 2
PROCEEDS_COL = 3
ORIG_CB_COL = 4
ORD_GAIN_COL = 5
ORD_CGT_COL = 6
DIV_CB_COL = 7
DIV_GAIN_COL = 8
DIV_TAX_COL = 9
RESET_IMPACT_COL = 10
HELPER_J_COL = 11
HELPER_K_COL = 12
LAST_VISIBLE_COL_LETTER = "J"


def _band(ws: Worksheet, row: int, text: str, last_col_letter: str = "L") -> None:
    ws.cell(row=row, column=1, value=text).font = SECTION_BAND_FONT
    ws.merge_cells(f"A{row}:{last_col_letter}{row}")
    for col_idx in range(1, ord(last_col_letter) - ord("A") + 2):
        ws.cell(row=row, column=col_idx).fill = SECTION_BAND_FILL


def _ord_taxable_formula(proceeds: str, orig: str, held: str) -> str:
    """v3.0: discount applies iff held > 12 months (no `discount_on` toggle)."""
    raw = f"({proceeds}-{orig})"
    return (
        f'=IF({proceeds}="","",'
        f'IF({raw}<=0,{raw},'
        f'IF({held}="Yes",{raw}*(1-discount_rate),{raw})))'
    )


def _div296_adj_formula(proceeds: str, cost_base_expr: str, held: str) -> str:
    """v3.0: discount applies iff held > 12 months (no `discount_on` toggle)."""
    raw = f"({proceeds}-{cost_base_expr})"
    return (
        f'=IF({proceeds}="","",'
        f'IF({raw}<=0,{raw},'
        f'IF({held}="Yes",{raw}*(1-discount_rate),{raw})))'
    )


def build(wb: Workbook) -> Worksheet:
    ws = wb.create_sheet(SHEET)

    # --- Title ---
    ws.cell(row=TITLE_ROW, column=1,
            value="Division 296 Cost Base Reset Model — Analyser").font = TITLE_FONT
    ws.merge_cells(f"A{TITLE_ROW}:L{TITLE_ROW}")

    # --- Row 2: Parameters in effect (v3.0 — replaces v2.x toggle-narration state strip) ---
    state_formula = (
        '="Parameters in effect: Rates "&TEXT(rate_tier1,"0%")&" tier 1 / "'
        '&TEXT(rate_tier2,"0%")&" tier 2 · Thresholds "'
        '&TEXT(threshold_1,"$#,##0")&" / "&TEXT(threshold_2,"$#,##0")'
        '&" · CGT discount "&TEXT(discount_rate,"0.00%")'
        '&" · Fund CGT "&TEXT(fund_cgt_rate,"0%")'
    )
    state_cell = ws.cell(row=STATE_STRIP_ROW, column=1, value=state_formula)
    state_cell.font = Font(name="Arial", size=9, italic=True, color="555555")
    state_cell.fill = STATE_STRIP_FILL
    state_cell.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.merge_cells(f"A{STATE_STRIP_ROW}:L{STATE_STRIP_ROW}")
    ws.row_dimensions[STATE_STRIP_ROW].height = 20

    # --- Row 3: Sample-data warning (conditional) ---
    sample_detect = (
        f'AND({INPUTS_SHEET}!A{REGISTER_FIRST_DATA_ROW}="P1",'
        f'{INPUTS_SHEET}!A{REGISTER_FIRST_DATA_ROW + 1}="S1",'
        f'{INPUTS_SHEET}!A{REGISTER_FIRST_DATA_ROW + 2}="L1")'
    )
    badge = ws.cell(
        row=SAMPLE_WARN_ROW, column=1,
        value=(
            f'=IF({sample_detect},'
            '"⚠  Sample data detected — figures below are illustrative only '
            'until the asset register on Inputs is replaced with the actual '
            'fund\'s holdings.","")'
        ),
    )
    badge.font = Font(name="Arial", size=10, bold=True, italic=True, color="8A6D00")
    badge.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.merge_cells(f"A{SAMPLE_WARN_ROW}:L{SAMPLE_WARN_ROW}")
    ws.row_dimensions[SAMPLE_WARN_ROW].height = 20
    sample_amber_rule = FormulaRule(
        formula=[sample_detect],
        fill=PatternFill("solid", fgColor="FFF4CE"),
    )
    ws.conditional_formatting.add(f"A{SAMPLE_WARN_ROW}:L{SAMPLE_WARN_ROW}", sample_amber_rule)

    # --- Row 4: Estimate disclaimer banner (always visible) ---
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

    # --- Fund summary block (v3.0 — side-by-side scenarios + signed Difference) ---
    _band(ws, FUND_BAND_ROW, "Fund summary")

    # Helper col ranges for the per-asset gain columns we SUMIF over.
    h_first, h_last = f"H{PERASSET_FIRST_ROW}", f"H{PERASSET_LAST_ROW}"   # elected
    l_first, l_last = f"L{PERASSET_FIRST_ROW}", f"L{PERASSET_LAST_ROW}"   # no-reset (helper)

    # --- Row 7: Header row ---
    fund_headers = [
        (1, "", None),                                  # cols A:B label zone
        (FUND_NORESET_COL, "If no reset (default)", "0F6E56"),    # green-ish
        (FUND_ELECTED_COL, "If elected to reset", "C7A752"),      # gold (reset theme)
        (FUND_DIFF_COL, "Difference (signed)", "1D3B34"),         # dark teal
    ]
    # Merge label cols A:B on header row too (keeps the band visually clean).
    ws.merge_cells(
        f"{FUND_LABEL_COL_RANGE[0]}{FUND_HEADER_ROW}:"
        f"{FUND_LABEL_COL_RANGE[1]}{FUND_HEADER_ROW}"
    )
    for col_idx, label, fg in fund_headers[1:]:
        c = ws.cell(row=FUND_HEADER_ROW, column=col_idx, value=label)
        c.font = Font(name="Arial", size=10, bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor=fg)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[FUND_HEADER_ROW].height = 30

    # --- Row 8: Fund Div 296 earnings ---
    earnings_label_range = (
        f"{FUND_LABEL_COL_RANGE[0]}{FUND_EARNINGS_ROW}:"
        f"{FUND_LABEL_COL_RANGE[1]}{FUND_EARNINGS_ROW}"
    )
    ws.merge_cells(earnings_label_range)
    elabel = ws.cell(row=FUND_EARNINGS_ROW, column=1, value="Fund Div 296 earnings")
    elabel.font = BODY_FONT
    elabel.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    # No-reset earnings = SUMIF over helper col L (per-asset no-reset gain).
    e_noreset = ws.cell(
        row=FUND_EARNINGS_ROW, column=FUND_NORESET_COL,
        value=f'=SUMIF({l_first}:{l_last},">0")',
    )
    e_noreset.number_format = FMT_CURRENCY
    e_noreset.comment = Comment(
        ("Sum of positive Div 296 adjusted gains across all 50 register rows, "
         "using the ORIGINAL cost base (i.e. assuming no reset election). "
         "Mirrors helper column L of the per-asset table."),
        "v3.0",
    )
    # Elected earnings = SUMIF over col H (per-asset adjusted gain, which always
    # uses the elected-reset cost base in v3.0).
    e_elected = ws.cell(
        row=FUND_EARNINGS_ROW, column=FUND_ELECTED_COL,
        value=f'=SUMIF({h_first}:{h_last},">0")',
    )
    e_elected.number_format = FMT_CURRENCY
    e_elected.comment = Comment(
        ("Sum of positive Div 296 adjusted gains across all 50 register rows, "
         "using the MARKET VALUE at 30 Jun 2026 as cost base (election made). "
         "Mirrors column H of the per-asset table."),
        "v3.0",
    )
    # Difference (signed) = elected - no_reset.
    e_diff = ws.cell(
        row=FUND_EARNINGS_ROW, column=FUND_DIFF_COL,
        value=f"=D{FUND_EARNINGS_ROW}-C{FUND_EARNINGS_ROW}",
    )
    e_diff.number_format = '$#,##0;[Red]($#,##0);"-"'

    # --- Rows 9-12: Per-member Div 296 tax (3 scenario columns + signed diff) ---
    # Cell references for the earnings driver of each scenario.
    earnings_cell_noreset = f"$C${FUND_EARNINGS_ROW}"
    earnings_cell_elected = f"$D${FUND_EARNINGS_ROW}"
    for i in range(ASSUMPTIONS.member_count):
        row = MEMBER_TAX_FIRST_ROW + i
        inputs_row = MEMBERS_FIRST_DATA_ROW + i
        # Label (merged A:B)
        ws.merge_cells(f"A{row}:B{row}")
        lbl = ws.cell(row=row, column=1, value=f"Member {i+1} Div 296 tax")
        lbl.font = BODY_FONT
        lbl.alignment = Alignment(horizontal="left", vertical="center", indent=1)
        # No-reset scenario
        t_no = ws.cell(
            row=row, column=FUND_NORESET_COL,
            value=per_member_div296_tax_formula(inputs_row, earnings_cell_noreset),
        )
        t_no.number_format = FMT_CURRENCY
        # Elected scenario
        t_el = ws.cell(
            row=row, column=FUND_ELECTED_COL,
            value=per_member_div296_tax_formula(inputs_row, earnings_cell_elected),
        )
        t_el.number_format = FMT_CURRENCY
        # Difference (signed)
        t_df = ws.cell(
            row=row, column=FUND_DIFF_COL,
            value=f"=D{row}-C{row}",
        )
        t_df.number_format = '$#,##0;[Red]($#,##0);"-"'

    # --- Row 13: Headline total ---
    ws.merge_cells(f"A{HEADLINE_ROW}:B{HEADLINE_ROW}")
    hlbl = ws.cell(row=HEADLINE_ROW, column=1, value="Total Div 296 tax (headline)")
    hlbl.font = Font(name="Arial", size=11, bold=True, color="1D3B34")
    hlbl.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    hlbl.fill = PatternFill("solid", fgColor="EFF5F3")
    for col_idx in (FUND_NORESET_COL, FUND_ELECTED_COL, FUND_DIFF_COL):
        col_letter = get_column_letter(col_idx)
        h_total = ws.cell(
            row=HEADLINE_ROW, column=col_idx,
            value=f"=SUM({col_letter}{MEMBER_TAX_FIRST_ROW}:{col_letter}{MEMBER_TAX_LAST_ROW})",
        )
        h_total.font = Font(name="Arial", size=11, bold=True, color="1D3B34")
        h_total.fill = PatternFill("solid", fgColor="EFF5F3")
        if col_idx == FUND_DIFF_COL:
            h_total.number_format = '$#,##0;[Red]($#,##0);"-"'
        else:
            h_total.number_format = FMT_CURRENCY
        h_total.border = Border(top=Side(style="thin", color="1D3B34"))

    # --- Per-asset analysis (elected-reset scenario only) ---
    _band(ws, PERASSET_BAND_ROW, "Per-asset analysis (elected-reset scenario)")
    headers = [
        "#",
        "Asset", "Projected sale proceeds", "Original cost base",
        "Ordinary taxable capital gain", "Ordinary CGT",
        "Div 296 cost base",
        "Div 296 adjusted taxable capital gain", "Div 296 tax",
        "Reset impact", "Helper: col H with reset", "Helper: col H without reset",
    ]
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

    # Hide helper columns K, L.
    ws.column_dimensions["K"].hidden = True
    ws.column_dimensions["L"].hidden = True

    currency_cols = [PROCEEDS_COL, ORIG_CB_COL, ORD_GAIN_COL, ORD_CGT_COL,
                     DIV_CB_COL, DIV_GAIN_COL, DIV_TAX_COL, RESET_IMPACT_COL,
                     HELPER_J_COL, HELPER_K_COL]

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

        # Inputs column layout (v2.3): A code / B name / C orig CB / D MV today /
        # E MV 30 Jun / F val source / G proceeds / H projected G/L / I held>12m
        code = f"{INPUTS_SHEET}!A{i_row}"
        name = f"{INPUTS_SHEET}!B{i_row}"
        orig = f"{INPUTS_SHEET}!C{i_row}"
        mv = f"{INPUTS_SHEET}!E{i_row}"
        proceeds = f"{INPUTS_SHEET}!G{i_row}"
        held = f"{INPUTS_SHEET}!I{i_row}"

        # Col A — Row number
        rn = ws.cell(row=a_row, column=ROWNUM_COL, value=offset + 1)
        rn.fill = ROWNUM_FILL
        rn.font = ROWNUM_FONT
        rn.alignment = Alignment(horizontal="center", vertical="center")
        # Col B — Asset display
        ws.cell(row=a_row, column=ASSET_COL,
                value=f'=IF({code}="","",{code}&" - "&{name})')
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
        # Col G — Div 296 cost base. v3.0: always the elected reset cost base.
        ws.cell(
            row=a_row, column=DIV_CB_COL,
            value=f'=IF({proceeds}="","",{mv})',
        )
        # Col H — Div 296 adjusted gain (elected-reset scenario).
        ws.cell(row=a_row, column=DIV_GAIN_COL,
                value=_div296_adj_formula(proceeds, mv, held))
        # Col I — Div 296 tax (pro-rata of the elected-reset headline at D{HEADLINE_ROW}).
        ws.cell(
            row=a_row, column=DIV_TAX_COL,
            value=(
                f'=IF({proceeds}="","",'
                f'IF(SUMIF({h_first}:{h_last},">0")=0,0,'
                f'MAX(0,H{a_row})/SUMIF({h_first}:{h_last},">0")*$D${HEADLINE_ROW}))'
            ),
        )
        # Helper K — col H WITH reset (cost base = MV). Unconditional.
        ws.cell(row=a_row, column=HELPER_J_COL,
                value=_div296_adj_formula(proceeds, mv, held))
        # Helper L — col H WITHOUT reset (cost base = original). Unconditional.
        ws.cell(row=a_row, column=HELPER_K_COL,
                value=_div296_adj_formula(proceeds, orig, held))
        # Col J — Reset impact = K − L.
        ws.cell(
            row=a_row, column=RESET_IMPACT_COL,
            value=f'=IF({proceeds}="","",K{a_row}-L{a_row})',
        )

        for col_idx in currency_cols:
            cell = ws.cell(row=a_row, column=col_idx)
            cell.number_format = FMT_CURRENCY
            if col_idx in DATA_FILL:
                cell.fill = DATA_FILL[col_idx]

        # Thin borders on every visible per-asset cell; medium-weight left border
        # on Reset Impact (col J) to highlight it.
        thin = Side(style="thin", color="D5DEDA")
        heavy = Side(style="medium", color="C29A3B")
        for col_idx in range(1, RESET_IMPACT_COL + 1):
            cell = ws.cell(row=a_row, column=col_idx)
            cell.border = Border(
                left=heavy if col_idx == RESET_IMPACT_COL else thin,
                right=thin, top=thin, bottom=thin,
            )
        ws.cell(row=a_row, column=RESET_IMPACT_COL).font = Font(
            name="Arial", size=10, bold=True, color="1D3B34",
        )

    # --- Totals row ---
    DOUBLE_TOP = Border(top=Side(style="double", color="1D3B34"))

    sigma_cell = ws.cell(row=TOTALS_ROW, column=ROWNUM_COL, value="Total")
    sigma_cell.fill = TOTALS_FILL
    sigma_cell.font = TOTALS_FONT
    sigma_cell.alignment = Alignment(horizontal="center", vertical="center")
    sigma_cell.border = DOUBLE_TOP

    totals_label = ws.cell(row=TOTALS_ROW, column=ASSET_COL, value="TOTAL")
    totals_label.font = TOTALS_FONT
    totals_label.fill = TOTALS_FILL
    totals_label.border = DOUBLE_TOP

    for col_letter, col_const in (
        ("C", PROCEEDS_COL),
        ("F", ORD_CGT_COL),
        ("H", DIV_GAIN_COL),
        ("I", DIV_TAX_COL),
    ):
        rng = f"{col_letter}{PERASSET_FIRST_ROW}:{col_letter}{PERASSET_LAST_ROW}"
        cell = ws.cell(row=TOTALS_ROW, column=col_const, value=f"=SUM({rng})")
        cell.number_format = FMT_CURRENCY
        cell.font = TOTALS_FONT
        cell.fill = TOTALS_FILL
        cell.border = DOUBLE_TOP

    for col_const in (ORIG_CB_COL, ORD_GAIN_COL, DIV_CB_COL, RESET_IMPACT_COL):
        c = ws.cell(row=TOTALS_ROW, column=col_const)
        c.fill = TOTALS_FILL
        c.border = DOUBLE_TOP

    # --- Trap-row + reset-impact legend (plain-English) ---
    legend = ws.cell(
        row=TRAP_LEGEND_ROW, column=1,
        value=("Reset impact (gold column) shows the estimated change in "
               "Division 296 gain if the reset election is applied. A "
               "negative amount (shown in green) generally indicates a "
               "favourable impact, while a positive amount (shown in red) "
               "may indicate additional Division 296 exposure. Particular "
               "care should be taken with red-shaded rows, where a reset "
               "creates taxable Division 296 gain on an asset currently in "
               "an unrealised loss position (the 'reset trap')."),
    )
    legend.font = Font(name="Arial", size=9, italic=True, color="666666")
    legend.alignment = Alignment(horizontal="left", vertical="center",
                                 wrap_text=True, indent=1)
    swatch = ws.cell(row=TRAP_LEGEND_ROW, column=1)
    swatch.fill = PatternFill("solid", fgColor="FBE9E9")
    ws.merge_cells(f"A{TRAP_LEGEND_ROW}:L{TRAP_LEGEND_ROW}")
    ws.row_dimensions[TRAP_LEGEND_ROW].height = 20

    # --- Reconciliation panel (stays at bottom per v3.0 decision 8) ---
    _band(ws, RECON_BAND_ROW, "Reconciliation")
    ws.cell(row=RECON_ORD_CGT_ROW, column=1,
            value="Ordinary CGT payable").font = BODY_FONT
    ws.cell(
        row=RECON_ORD_CGT_ROW, column=2,
        value=f"=SUM(F{PERASSET_FIRST_ROW}:F{PERASSET_LAST_ROW})",
    ).number_format = FMT_CURRENCY

    ws.cell(row=RECON_DIV296_ROW, column=1,
            value="Div 296 tax payable (elected-reset headline)").font = BODY_FONT
    # v3.0: headline lives at col D row 13 (elected-reset scenario).
    ws.cell(row=RECON_DIV296_ROW, column=2,
            value=f"=D{HEADLINE_ROW}").number_format = FMT_CURRENCY

    ws.cell(row=RECON_LOSSES_ROW, column=1,
            value="Capital losses carried forward").font = BODY_FONT
    cf_terms = []
    for offset in range(ASSUMPTIONS.asset_register_rows):
        i_row = REGISTER_FIRST_DATA_ROW + offset
        orig_i = f"{INPUTS_SHEET}!C{i_row}"
        proc_i = f"{INPUTS_SHEET}!G{i_row}"
        cf_terms.append(f'IF({proc_i}="",0,MAX(0,{orig_i}-{proc_i}))')
    ws.cell(
        row=RECON_LOSSES_ROW, column=2,
        value="=" + "+".join(cf_terms),
    ).number_format = FMT_CURRENCY

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

    # --- Trap shading on per-asset rows: red when ord raw gain < 0 AND Div 296 adj gain > 0 ---
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

    # --- Reset Impact (col J) — green for favourable (<0), red for unfavourable (>0) ---
    j_range = f"J{PERASSET_FIRST_ROW}:J{PERASSET_LAST_ROW}"
    j_fav_rule = FormulaRule(
        formula=[f"AND(ISNUMBER(J{PERASSET_FIRST_ROW}),J{PERASSET_FIRST_ROW}<0)"],
        font=Font(name="Arial", size=10, bold=True, color="0B6E4F"),
    )
    j_unfav_rule = FormulaRule(
        formula=[f"AND(ISNUMBER(J{PERASSET_FIRST_ROW}),J{PERASSET_FIRST_ROW}>0)"],
        font=Font(name="Arial", size=10, bold=True, color="A61B1B"),
    )
    ws.conditional_formatting.add(j_range, j_fav_rule)
    ws.conditional_formatting.add(j_range, j_unfav_rule)

    # --- Column widths ---
    widths = [5, 32, 14, 16, 16, 22, 14, 24, 14, 14, 14, 14]
    for col_idx, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(col_idx)].width = w
    # Free scroll (no freeze panes). Header row repeats on every printed page.
    ws.print_title_rows = f"{PERASSET_HEADER_ROW}:{PERASSET_HEADER_ROW}"

    # --- Print header watermark ---
    ws.oddHeader.center.text = "ILLUSTRATIVE — NOT ADVICE"
    ws.oddHeader.center.size = 28
    ws.oddHeader.center.color = "CCCCCC"

    # --- Sheet protection ---
    ws.protection.sheet = True
    ws.protection.formatColumns = False
    ws.protection.formatRows = False

    return ws
