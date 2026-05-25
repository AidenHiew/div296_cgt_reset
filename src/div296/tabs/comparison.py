"""Comparison tab — print-ready landscape A4 tearsheet (v1.5 redesign).

Layout (top to bottom):
    Rows 1-10    Header block (banner, title, firm/prepared/date, disclaimer, logo)
    Row 12       Section band: "Side-by-side comparison"
    Rows 13-14   Fund-context strip (TSB / proportion / discount / tier)
    Row 16       Band: "Headline — total Div 296 tax"
    Rows 17-18   Three metric cards: Scenario A | Scenario B | Net effect
    Row 20       Band: "Per-scenario subtotals"
    Rows 21-25   Subtotals table (Earnings / Ord CGT (unchanged) / Div 296 tax / Total)
    Row 27       Band: "Per-asset detail (first 15 assets)"
    Rows 28-29   Panel headers + column sub-headers
    Rows 30-44   15 data rows
    Row 45       Note: "Showing first 15 assets — see Analyser for the full register"
    Rows 47-48   Reminder + sort-order notes
    Rows 50+     Chart

Columns:
    A-E:  Panel A (no reset)         — Asset / Proceeds / Cost base / Adj gain / Tax
    F:    Δ column (panel B − panel A adj gain)
    G-K:  Panel B (reset elected)    — same five columns
    L-M:  Hidden helper block (fund earnings, member tax, headlines, chart data)

Both panels compute independently of the master reset toggle, so the
comparison always shows the *real* before/after picture regardless of
what the user has selected on Inputs.

Deviation from spec §8: panels carry more columns than the "lean" spec
(Proceeds + Div 296 cost base added) per user request; subtotals at top
rather than bottom; total tax burden row added. Net effect remains
neutral ("Net effect = A − B") — no recommendation language.
"""

from __future__ import annotations

from openpyxl.comments import Comment
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Protection, Side
from openpyxl.workbook import Workbook
from openpyxl.worksheet.worksheet import Worksheet

from div296.assumptions import ASSUMPTIONS
from div296.styles import (
    BODY_FONT, CENTER, FMT_CURRENCY, FMT_CURRENCY_DELTA, FMT_PERCENT,
    INPUT_FILL, INPUT_FONT, SECTION_BAND_FILL, SECTION_BAND_FONT,
    THIN_BOX, TITLE_FONT, TRAP_FILL,
    PROC_DATA_FILL, ORD_DATA_FILL, DIV_DATA_FILL,
)
from div296.tabs.inputs import (
    MEMBERS_FIRST_DATA_ROW, REGISTER_FIRST_DATA_ROW,
    REGISTER_LAST_DATA_ROW,
)


SHEET = "Comparison"
INPUTS_SHEET = "'Inputs'"
ANALYSER_SHEET = "'Analyser'"

# How many assets to display on the printable comparison page.
# v1.7: 10 rows, sorted by absolute Δ (Scenario B − Scenario A) descending so
# the assets most affected by the reset election appear at the top. Empty
# register rows sort to the bottom and render blank.
DISPLAY_ROWS = 10

# --- Layout rows ---
WATERMARK_ROW = 1
TITLE_ROW = 2
HEADER_FIRM_ROW = 4
HEADER_PREPARED_FOR_ROW = 5
HEADER_PREPARED_BY_ROW = 6
HEADER_DATE_ROW = 7
HEADER_DISCLAIMER_ROW = 8
HEADER_LOGO_ROW = 10

BAND_BODY_ROW = 12
CONTEXT_LABEL_ROW = 13
# v2.5 FB-8: strip now shows per-member TSB (rows 14..14+member_count-1)
# plus a Total row. Other context tiles (proportion, discount, tier) get
# vertically merged across those rows on the right.
CONTEXT_VALUE_ROW = 14   # first member row (kept name for back-compat)
CONTEXT_MEMBER_LAST_ROW = CONTEXT_VALUE_ROW + ASSUMPTIONS.member_count - 1   # row 17
CONTEXT_TOTAL_ROW = CONTEXT_MEMBER_LAST_ROW + 1                              # row 18

BAND_HEADLINE_ROW = CONTEXT_TOTAL_ROW + 1    # row 19 (was 16)
CARD_LABEL_ROW = BAND_HEADLINE_ROW + 1       # row 20 (was 17)
CARD_VALUE_ROW = CARD_LABEL_ROW + 1          # row 21 (was 18)
CARD_FOOTNOTE_ROW = CARD_VALUE_ROW + 1       # row 22 (was 19)

BAND_SUBTOTALS_ROW = CARD_FOOTNOTE_ROW + 1   # row 23 (was 20)
SUBTOTAL_HEADER_ROW = BAND_SUBTOTALS_ROW + 1 # row 24 (was 21)
SUBTOTAL_EARNINGS_ROW = SUBTOTAL_HEADER_ROW + 1   # row 25
SUBTOTAL_ORD_CGT_ROW = SUBTOTAL_EARNINGS_ROW + 1  # row 26
SUBTOTAL_DIV296_ROW = SUBTOTAL_ORD_CGT_ROW + 1    # row 27
SUBTOTAL_BURDEN_ROW = SUBTOTAL_DIV296_ROW + 1     # row 28

# v2.3: Per-member breakdown block between subtotals and per-asset detail.
PER_MEMBER_BAND_ROW = SUBTOTAL_BURDEN_ROW + 2          # row 27
PER_MEMBER_HEADER_ROW = PER_MEMBER_BAND_ROW + 1        # row 28
PER_MEMBER_FIRST_ROW = PER_MEMBER_HEADER_ROW + 1       # row 29
PER_MEMBER_LAST_ROW = PER_MEMBER_FIRST_ROW + ASSUMPTIONS.member_count - 1   # row 32

BAND_DETAIL_ROW = PER_MEMBER_LAST_ROW + 2              # row 34 (was 27)
PANEL_TITLE_ROW = BAND_DETAIL_ROW + 1                  # row 35 (was 28)
PANEL_HEADER_ROW = BAND_DETAIL_ROW + 2                 # row 36 (was 29)
DATA_FIRST_ROW = BAND_DETAIL_ROW + 3                   # row 37 (was 30)
DATA_LAST_ROW = DATA_FIRST_ROW + DISPLAY_ROWS - 1      # row 46 (was 44)
DATA_OVERFLOW_NOTE_ROW = DATA_LAST_ROW + 1             # row 47 (was 45)

REMINDER_ROW = DATA_OVERFLOW_NOTE_ROW + 1              # row 48 (was 46)
SORT_NOTE_ROW = REMINDER_ROW + 1                       # row 49 (was 47)

# v2.5 FB-3: chart removed. Print area now ends with the sort note (SORT_NOTE_ROW).

# --- Layout columns ---
PANEL_A_COLS = ("A", "B", "C", "D", "E")   # Asset, Proceeds, Cost base, Adj gain, Tax
PANEL_B_COLS = ("F", "G", "H", "I", "J")   # v2.5 FB-7: moved left (was G-K)
DELTA_COL = "K"                             # v2.5 FB-7: Change at the END now (was F)
LAST_VISIBLE_COL = DELTA_COL                # "K"

# Hidden helper columns (right of the visible panels)
HELPER_COL_A = "L"   # Scenario A working column
HELPER_COL_B = "M"   # Scenario B working column

# v1.7: per-register-row helper grid for sort-by-impact lookup.
# Cols N/O/P/Q span register rows (REGISTER_FIRST_DATA_ROW..REGISTER_LAST_DATA_ROW).
PER_REG_GAIN_A_COL = "N"   # Scenario A adj gain per register row
PER_REG_GAIN_B_COL = "O"   # Scenario B adj gain per register row
PER_REG_DELTA_COL = "P"    # |Δ| + row tiebreaker; empty rows = -1
# Col R: matched register row for each visible display row (rows 30..39).
MATCHED_ROW_COL = "R"

# Hidden helper rows in cols L/M
HELPER_FUND_EARNINGS_ROW = 1
HELPER_MEMBER_TAX_FIRST_ROW = 2
HELPER_MEMBER_TAX_LAST_ROW = HELPER_MEMBER_TAX_FIRST_ROW + ASSUMPTIONS.member_count - 1
HELPER_HEADLINE_ROW = HELPER_MEMBER_TAX_LAST_ROW + 1   # row 6

# --- Styling palette specific to v1.5 ---
CARD_FILL_A = PatternFill("solid", fgColor="EFF5F3")    # very light teal
CARD_FILL_B = PatternFill("solid", fgColor="EAF3FB")    # very light blue
CARD_FILL_DELTA = PatternFill("solid", fgColor="FFF5E1")  # very light gold
CARD_BORDER = Border(
    left=Side(style="medium", color="1D3B34"),
    right=Side(style="medium", color="1D3B34"),
    top=Side(style="medium", color="1D3B34"),
    bottom=Side(style="medium", color="1D3B34"),
)
CARD_LABEL_FONT = Font(name="Arial", size=10, italic=True, color="1D3B34")
CARD_VALUE_FONT = Font(name="Arial", size=22, bold=True, color="1D3B34")
CONTEXT_LABEL_FONT = Font(name="Arial", size=9, color="666666")
CONTEXT_VALUE_FONT = Font(name="Arial", size=11, bold=True, color="1D3B34")
TOTAL_BURDEN_FONT = Font(name="Arial", size=11, bold=True, color="1D3B34")
TOTAL_BURDEN_FILL = PatternFill("solid", fgColor="EFF5F3")
COST_BASE_ACCENT_FILL = PatternFill("solid", fgColor="FFF8E6")   # subtle gold tint
DELTA_HEADER_FILL = PatternFill("solid", fgColor="FFF5E1")
DELTA_FONT = Font(name="Arial", size=10, italic=True, color="8A6D00")


# --- Small builders ---

def _band(ws: Worksheet, row: int, text: str, last_col_letter: str = LAST_VISIBLE_COL) -> None:
    ws.cell(row=row, column=1, value=text).font = SECTION_BAND_FONT
    ws.merge_cells(f"A{row}:{last_col_letter}{row}")
    for col_idx in range(1, ord(last_col_letter) - ord("A") + 2):
        ws.cell(row=row, column=col_idx).fill = SECTION_BAND_FILL


def _input_cell(ws: Worksheet, coord: str, value=None) -> None:
    cell = ws[coord]
    if value is not None:
        cell.value = value
    cell.font = INPUT_FONT
    cell.fill = INPUT_FILL
    cell.border = THIN_BOX


def _div296_adj_formula(proceeds: str, cost_base_expr: str, held: str) -> str:
    raw = f"({proceeds}-{cost_base_expr})"
    return (
        f'=IF({proceeds}="","",'
        f'IF({raw}<=0,{raw},'
        f'IF(AND({held}="Yes",discount_on="ON"),{raw}*(1-discount_rate),{raw})))'
    )


def _member_tax_formula(member_inputs_row: int, earnings_cell: str) -> str:
    """Same shape as Analyser's per-member tax, with earnings cell parameterised."""
    tsb = f"{INPUTS_SHEET}!B{member_inputs_row}"
    split = f"{INPUTS_SHEET}!C{member_inputs_row}"
    auto_p = f"{INPUTS_SHEET}!D{member_inputs_row}"
    override = f"{INPUTS_SHEET}!E{member_inputs_row}"
    earnings_m = f"{earnings_cell}*{split}"
    effective_p = f'IF({override}="",{auto_p},{override})'
    band1 = f"MAX(0,MIN({tsb},threshold_2)-threshold_1)/{tsb}"
    band2 = f"MAX(0,{tsb}-threshold_2)/{tsb}"
    tier_on = f"{earnings_m}*{band1}*rate_tier1 + {earnings_m}*{band2}*rate_tier2"
    tier_off = f"{earnings_m}*{effective_p}*rate_tier1"
    return (
        f'=IF(OR({tsb}="",{split}="",{tsb}<=0,{split}<=0,{earnings_cell}<=0),0,'
        f'IF(tier10_on="ON",{tier_on},{tier_off}))'
    )


# --- Big builders for each visual block ---

def _build_header_block(ws: Worksheet) -> None:
    # Row 1: on-screen ILLUSTRATIVE banner
    watermark_font = Font(name="Arial", size=14, bold=True, italic=True, color="999999")
    cell = ws.cell(row=WATERMARK_ROW, column=1, value="ILLUSTRATIVE — NOT ADVICE")
    cell.font = watermark_font
    cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.merge_cells(f"A{WATERMARK_ROW}:{LAST_VISIBLE_COL}{WATERMARK_ROW}")

    # Row 2: Title
    title_cell = ws.cell(row=TITLE_ROW, column=1,
                         value="Division 296 Cost Base Reset Model — Comparison")
    title_cell.font = TITLE_FONT
    title_cell.alignment = Alignment(horizontal="center")
    ws.merge_cells(f"A{TITLE_ROW}:{LAST_VISIBLE_COL}{TITLE_ROW}")

    # Rows 4-7: editable cells for firm / prepared for / by / date
    header_rows = [
        (HEADER_FIRM_ROW, "Firm name:"),
        (HEADER_PREPARED_FOR_ROW, "Prepared for:"),
        (HEADER_PREPARED_BY_ROW, "Prepared by:"),
        (HEADER_DATE_ROW, "Date:"),
    ]
    for row, label in header_rows:
        ws.cell(row=row, column=1, value=label).font = BODY_FONT
        _input_cell(ws, f"B{row}", value="")
        ws.merge_cells(f"B{row}:E{row}")

    # Row 8: disclaimer line (read-only)
    disc = ws.cell(row=HEADER_DISCLAIMER_ROW, column=1,
                   value="Illustrative only — not financial, tax or legal advice.")
    disc.font = Font(name="Arial", size=10, italic=True, color="666666")
    ws.merge_cells(f"A{HEADER_DISCLAIMER_ROW}:{LAST_VISIBLE_COL}{HEADER_DISCLAIMER_ROW}")

    # Row 10: logo placeholder (top-right area)
    logo = ws.cell(row=HEADER_LOGO_ROW, column=10, value="[ logo ]")
    logo.font = Font(name="Arial", size=10, italic=True, color="999999")
    logo.alignment = Alignment(horizontal="center", vertical="center")
    logo.fill = PatternFill("solid", fgColor="F5F5F5")
    ws.merge_cells(f"J{HEADER_LOGO_ROW}:{LAST_VISIBLE_COL}{HEADER_LOGO_ROW + 1}")

    # v2.2.0: Sample-data warning (row 9, full width left of the logo block).
    # Formula returns the warning string ONLY when the first three asset codes
    # still match the sample register; CF paints amber only in that case.
    sample_detect = (
        f'AND({INPUTS_SHEET}!A{REGISTER_FIRST_DATA_ROW}="P1",'
        f'{INPUTS_SHEET}!A{REGISTER_FIRST_DATA_ROW + 1}="S1",'
        f'{INPUTS_SHEET}!A{REGISTER_FIRST_DATA_ROW + 2}="L1")'
    )
    sample_row = HEADER_DISCLAIMER_ROW + 1   # row 9
    badge = ws.cell(
        row=sample_row, column=1,
        value=(
            f'=IF({sample_detect},'
            '"⚠  Sample data detected — the figures shown are illustrative '
            'only until the asset register on Inputs is replaced with the '
            "actual fund's holdings.\",\"\")"
        ),
    )
    badge.font = Font(name="Arial", size=10, bold=True, italic=True, color="8A6D00")
    badge.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.merge_cells(f"A{sample_row}:I{sample_row}")
    ws.row_dimensions[sample_row].height = 20
    amber_rule = FormulaRule(
        formula=[sample_detect],
        fill=PatternFill("solid", fgColor="FFF4CE"),
    )
    ws.conditional_formatting.add(f"A{sample_row}:I{sample_row}", amber_rule)


def _build_context_strip(ws: Worksheet) -> None:
    """Top context block. v2.5 FB-8: the left tile now shows a per-member
    TSB mini-table (Member 1..4 + Total) instead of a single fund aggregate.
    The other three tiles (proportion, discount, tier) sit on the right and
    vertically merge across the same row range so the strip reads as one panel.
    """
    last_input_member_row = MEMBERS_FIRST_DATA_ROW + ASSUMPTIONS.member_count - 1

    # --- Row 13 labels (four horizontal merges across cols A:K) ---
    labels = [
        ("A:C", "Members & TSB"),
        ("D:F", "Highest member proportion above $3m"),
        ("G:H", "CGT discount"),
        ("I:K", "$10m / +25% tier"),
    ]
    for merge_range, label in labels:
        start, end = merge_range.split(":")
        cell = ws[f"{start}{CONTEXT_LABEL_ROW}"]
        cell.value = label
        cell.font = CONTEXT_LABEL_FONT
        cell.alignment = Alignment(horizontal="left", indent=1)
        ws.merge_cells(f"{start}{CONTEXT_LABEL_ROW}:{end}{CONTEXT_LABEL_ROW}")

    # --- Left tile (cols A:C) — per-member TSB mini-table ---
    # Col A always shows the placeholder label; col B-C merged for TSB value
    # (blank when the corresponding Inputs row has no TSB).
    for i in range(ASSUMPTIONS.member_count):
        c_row = CONTEXT_VALUE_ROW + i
        inputs_row = MEMBERS_FIRST_DATA_ROW + i
        tsb_ref = f"{INPUTS_SHEET}!B{inputs_row}"

        label_cell = ws.cell(row=c_row, column=1, value=f"Member {i+1}")
        label_cell.font = CONTEXT_LABEL_FONT
        label_cell.alignment = Alignment(horizontal="left", indent=1)

        value_cell = ws.cell(row=c_row, column=2,
                             value=f'=IF({tsb_ref}>0,{tsb_ref},"")')
        value_cell.font = CONTEXT_VALUE_FONT
        value_cell.number_format = FMT_CURRENCY
        value_cell.alignment = Alignment(horizontal="right", indent=1)
        ws.merge_cells(f"B{c_row}:C{c_row}")

        ws.row_dimensions[c_row].height = 16

    # Total row underneath the members.
    total_label = ws.cell(row=CONTEXT_TOTAL_ROW, column=1, value="Total")
    total_label.font = Font(name="Arial", size=10, bold=True, color="1D3B34")
    total_label.alignment = Alignment(horizontal="left", indent=1)

    total_value = ws.cell(
        row=CONTEXT_TOTAL_ROW, column=2,
        value=f"=SUM({INPUTS_SHEET}!B{MEMBERS_FIRST_DATA_ROW}:B{last_input_member_row})",
    )
    total_value.font = Font(name="Arial", size=11, bold=True, color="1D3B34")
    total_value.number_format = FMT_CURRENCY
    total_value.alignment = Alignment(horizontal="right", indent=1)
    ws.merge_cells(f"B{CONTEXT_TOTAL_ROW}:C{CONTEXT_TOTAL_ROW}")
    ws.row_dimensions[CONTEXT_TOTAL_ROW].height = 18

    # --- Right tiles (D:F, G:H, I:K) — vertically merged across the
    # member-rows + total-row range so each value reads as one panel.
    strip_first_row = CONTEXT_VALUE_ROW
    strip_last_row = CONTEXT_TOTAL_ROW
    tile_values = [
        ("D", "F",
         f"=MAX({INPUTS_SHEET}!D{MEMBERS_FIRST_DATA_ROW}:D{last_input_member_row})",
         FMT_PERCENT),
        ("G", "H", "=discount_on", None),
        ("I", "K", "=tier10_on",   None),
    ]
    for start, end, formula, fmt in tile_values:
        cell = ws[f"{start}{strip_first_row}"]
        cell.value = formula
        cell.font = CONTEXT_VALUE_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center")
        if fmt:
            cell.number_format = fmt
        ws.merge_cells(f"{start}{strip_first_row}:{end}{strip_last_row}")


def _build_per_register_helpers(ws: Worksheet) -> tuple[str, str, str]:
    """Per-register-row helper grid (cols N/O/P) used to sort the per-asset
    detail panel by |Δ| descending.

    N = Scenario A Div 296 adj gain (cost base = original)
    O = Scenario B Div 296 adj gain (cost base = MV)
    P = |tax_b − tax_a| + row tiebreaker, where tax_x = MAX(0,gain_x)/SUMIF(>0)
        × headline_x (cells L6/M6). v2.5 FB-7: switched from gain-delta to
        tax-delta so the "top 10 most affected" matches the visible Change col
        (which now displays tax delta, not gain delta).
        Empty register rows return -1 so LARGE pushes them to the bottom.

    Returns the absolute ranges (e.g. '$N$20:$N$69').
    """
    n_first, n_last = REGISTER_FIRST_DATA_ROW, REGISTER_LAST_DATA_ROW
    # Range refs reused inside the per-row tax formula.
    a_range = f"${PER_REG_GAIN_A_COL}${n_first}:${PER_REG_GAIN_A_COL}${n_last}"
    b_range = f"${PER_REG_GAIN_B_COL}${n_first}:${PER_REG_GAIN_B_COL}${n_last}"
    headline_a_abs = f"${HELPER_COL_A}${HELPER_HEADLINE_ROW}"   # $L$6
    headline_b_abs = f"${HELPER_COL_B}${HELPER_HEADLINE_ROW}"   # $M$6
    for n in range(n_first, n_last + 1):
        # v2.3 Inputs column layout: A code / B name / C orig CB / D MV today
        # / E MV 30 Jun / F val source / G proceeds / H projected G/L / I held
        proceeds = f"{INPUTS_SHEET}!G{n}"   # was H
        orig = f"{INPUTS_SHEET}!C{n}"        # was D
        mv = f"{INPUTS_SHEET}!E{n}"          # was F
        held = f"{INPUTS_SHEET}!I{n}"        # unchanged
        ws[f"{PER_REG_GAIN_A_COL}{n}"] = _div296_adj_formula(proceeds, orig, held)
        ws[f"{PER_REG_GAIN_B_COL}{n}"] = _div296_adj_formula(proceeds, mv, held)
        # v2.5: per-row tax inlined into the sort metric so the helper grid
        # itself measures tax impact, not raw gain impact.
        tax_a = (
            f'IF(SUMIF({a_range},">0")=0,0,'
            f'MAX(0,{PER_REG_GAIN_A_COL}{n})/SUMIF({a_range},">0")*{headline_a_abs})'
        )
        tax_b = (
            f'IF(SUMIF({b_range},">0")=0,0,'
            f'MAX(0,{PER_REG_GAIN_B_COL}{n})/SUMIF({b_range},">0")*{headline_b_abs})'
        )
        # Tiebreaker: small positive amount decreasing with register row, so
        # earlier-listed assets win ties without ever turning P negative.
        tiebreak = f"({n_last}-ROW())*0.001"
        ws[f"{PER_REG_DELTA_COL}{n}"] = (
            f'=IF({proceeds}="",-1,ABS(({tax_b})-({tax_a}))+{tiebreak})'
        )
    return (
        a_range, b_range,
        f"${PER_REG_DELTA_COL}${n_first}:${PER_REG_DELTA_COL}${n_last}",
    )


def _build_helpers(
    ws: Worksheet, gain_a_range: str, gain_b_range: str,
) -> tuple[str, str]:
    """Hidden L/M block: fund earnings (over the full register), per-member tax,
    and headline totals. Returns (headline_a_abs, headline_b_abs)."""
    ws[f"{HELPER_COL_A}{HELPER_FUND_EARNINGS_ROW}"] = f'=SUMIF({gain_a_range},">0")'
    ws[f"{HELPER_COL_B}{HELPER_FUND_EARNINGS_ROW}"] = f'=SUMIF({gain_b_range},">0")'

    a_earnings = f"${HELPER_COL_A}${HELPER_FUND_EARNINGS_ROW}"
    b_earnings = f"${HELPER_COL_B}${HELPER_FUND_EARNINGS_ROW}"

    for i in range(ASSUMPTIONS.member_count):
        helper_row = HELPER_MEMBER_TAX_FIRST_ROW + i
        inputs_row = MEMBERS_FIRST_DATA_ROW + i
        ws[f"{HELPER_COL_A}{helper_row}"] = _member_tax_formula(inputs_row, a_earnings)
        ws[f"{HELPER_COL_B}{helper_row}"] = _member_tax_formula(inputs_row, b_earnings)

    headline_a_cell = f"{HELPER_COL_A}{HELPER_HEADLINE_ROW}"
    headline_b_cell = f"{HELPER_COL_B}{HELPER_HEADLINE_ROW}"
    ws[headline_a_cell] = (
        f"=SUM({HELPER_COL_A}{HELPER_MEMBER_TAX_FIRST_ROW}:"
        f"{HELPER_COL_A}{HELPER_MEMBER_TAX_LAST_ROW})"
    )
    ws[headline_b_cell] = (
        f"=SUM({HELPER_COL_B}{HELPER_MEMBER_TAX_FIRST_ROW}:"
        f"{HELPER_COL_B}{HELPER_MEMBER_TAX_LAST_ROW})"
    )

    # Hide helper columns (L/M plus the per-register grid + matched-row col R).
    for col_letter in (HELPER_COL_A, HELPER_COL_B,
                       PER_REG_GAIN_A_COL, PER_REG_GAIN_B_COL,
                       PER_REG_DELTA_COL, "Q", MATCHED_ROW_COL):
        ws.column_dimensions[col_letter].hidden = True

    abs_a = f"${HELPER_COL_A}${HELPER_HEADLINE_ROW}"
    abs_b = f"${HELPER_COL_B}${HELPER_HEADLINE_ROW}"
    return abs_a, abs_b


def _build_metric_cards(ws: Worksheet, headline_a: str, headline_b: str) -> None:
    """Three boxed cards — Scenario A / Scenario B / Net effect."""
    _band(ws, BAND_HEADLINE_ROW, "Headline — total Div 296 tax")

    # v2.5 FB-4: labels standardised across all sections — "If no reset (default)",
    # "If elected to reset", "Change". Change card formula is SIGNED (reset −
    # default), so a reset that reduces tax shows a red-bracket negative.
    cards = [
        ("A:D", "If no reset (default)",  f"={headline_a[1:]}",                       FMT_CURRENCY,       CARD_FILL_A),
        ("E:H", "If elected to reset",    f"={headline_b[1:]}",                       FMT_CURRENCY,       CARD_FILL_B),
        ("I:K", "Change",                 f"={headline_b[1:]}-{headline_a[1:]}",      FMT_CURRENCY_DELTA, CARD_FILL_DELTA),
    ]
    for merge_range, label, value_formula, value_fmt, fill in cards:
        start = merge_range.split(":")[0]
        end = merge_range.split(":")[1]

        # Label row
        label_cell = ws[f"{start}{CARD_LABEL_ROW}"]
        label_cell.value = label
        label_cell.font = CARD_LABEL_FONT
        label_cell.alignment = Alignment(horizontal="center", vertical="center")
        ws.merge_cells(f"{start}{CARD_LABEL_ROW}:{end}{CARD_LABEL_ROW}")

        # Value row (big)
        value_cell = ws[f"{start}{CARD_VALUE_ROW}"]
        value_cell.value = value_formula
        value_cell.font = CARD_VALUE_FONT
        value_cell.number_format = value_fmt
        value_cell.alignment = Alignment(horizontal="center", vertical="center")
        ws.merge_cells(f"{start}{CARD_VALUE_ROW}:{end}{CARD_VALUE_ROW}")

        # Border + fill across both rows of the card
        for row in (CARD_LABEL_ROW, CARD_VALUE_ROW):
            for col_letter in _expand_cols(start, end):
                cell = ws[f"{col_letter}{row}"]
                cell.fill = fill
                cell.border = CARD_BORDER

    ws.row_dimensions[CARD_LABEL_ROW].height = 18
    ws.row_dimensions[CARD_VALUE_ROW].height = 36

    # v2.3: Combined headline explanation + Year-1 caveat.
    # Replaces the v2.2.0 standalone Year-1 caveat — explanation comes first so
    # the reader frames the headline figures before reading any subtotals.
    footnote = ws.cell(
        row=CARD_FOOTNOTE_ROW, column=1,
        value=("The headline figures summarise the estimated Division 296 "
               "impact under each scenario. They help compare the relative "
               "outcomes of each option rather than provide a final tax "
               "calculation.\nHeadline figures are Year 1 only — Div 296 is "
               "assessed annually, and the same earnings hit recurs each year "
               "a member's TSB remains above $3m (thresholds index over time)."),
    )
    footnote.font = Font(name="Arial", size=9, italic=True, color="666666")
    footnote.alignment = Alignment(horizontal="left", vertical="center",
                                   wrap_text=True, indent=1)
    ws.merge_cells(f"A{CARD_FOOTNOTE_ROW}:{LAST_VISIBLE_COL}{CARD_FOOTNOTE_ROW}")
    ws.row_dimensions[CARD_FOOTNOTE_ROW].height = 48


def _expand_cols(start: str, end: str) -> list[str]:
    return [chr(c) for c in range(ord(start), ord(end) + 1)]


def _build_subtotals(ws: Worksheet, headline_a: str, headline_b: str,
                     gain_a_range: str, gain_b_range: str) -> None:
    """Four-row subtotal table: earnings, ord CGT, Div 296 tax, total burden."""
    _band(ws, BAND_SUBTOTALS_ROW, "Per-scenario subtotals")

    # Header row — v2.5 FB-5: labels standardised across all sections.
    header_cells = [
        ("A", "Subtotal"),
        ("B", "If no reset (default)"),
        ("C", "If elected to reset"),
        ("D", "Change"),
    ]
    for col, text in header_cells:
        c = ws[f"{col}{SUBTOTAL_HEADER_ROW}"]
        c.value = text
        c.font = SECTION_BAND_FONT
        c.fill = SECTION_BAND_FILL
        c.alignment = CENTER

    # Reference to the Analyser's Ordinary CGT total — this is the same in both
    # scenarios (ordinary CGT doesn't depend on reset election).
    # Analyser B74 = Ordinary CGT payable (reconciliation panel). The state strip
    # added in v2.0.0 shifted this from B73 → B74 but the Comparison reference
    # wasn't updated; fixed in v2.2.0.
    ord_cgt_ref = f"={ANALYSER_SHEET}!B74"

    # v2.3 C-3: per-row definitions surfaced as cell Comments on the col-A label
    # so the reader can hover for a plain-English explanation of each subtotal.
    rows = [
        (SUBTOTAL_EARNINGS_ROW, "Div 296 earnings",
         f"=SUMIF({gain_a_range},\">0\")", f"=SUMIF({gain_b_range},\">0\")",
         "The total estimated Division 296 adjusted taxable capital gain "
         "across all assets included in the analysis. Excludes losses."),
        (SUBTOTAL_ORD_CGT_ROW,  "Ordinary CGT (unchanged by reset)",
         ord_cgt_ref, ord_cgt_ref,
         "Ordinary capital gains tax (fund CGT rate applied to the realised "
         "gain after any 1/3 CGT discount). Same in both scenarios — the "
         "reset election affects Div 296, not ordinary CGT."),
        (SUBTOTAL_DIV296_ROW,   "Div 296 tax (headline)",
         f"={headline_a[1:]}", f"={headline_b[1:]}",
         "Division 296 additional tax payable. Each member's share is "
         "computed from their TSB proportion above the $3m threshold, with "
         "the +25% tier band applied where the toggle is enabled."),
        (SUBTOTAL_BURDEN_ROW,   "TOTAL TAX BURDEN",
         f"=B{SUBTOTAL_ORD_CGT_ROW}+B{SUBTOTAL_DIV296_ROW}",
         f"=C{SUBTOTAL_ORD_CGT_ROW}+C{SUBTOTAL_DIV296_ROW}",
         "Ordinary CGT + Div 296 tax. Year 1 only — see the headline "
         "footnote about projecting these figures forward."),
    ]
    subtotal_border = Border(
        left=Side(style="thin", color="C5CECA"),
        right=Side(style="thin", color="C5CECA"),
        top=Side(style="thin", color="C5CECA"),
        bottom=Side(style="thin", color="C5CECA"),
    )
    for row, label, a_val, b_val, definition in rows:
        ws[f"A{row}"] = label
        ws[f"B{row}"] = a_val
        ws[f"C{row}"] = b_val
        # v2.5 FB-5: Change is SIGNED reset − default (negative = saving).
        ws[f"D{row}"] = f"=C{row}-B{row}"
        ws[f"A{row}"].alignment = Alignment(wrap_text=True, vertical="center")
        ws[f"A{row}"].comment = Comment(definition, "v2.3")
        for col in ("B", "C"):
            ws[f"{col}{row}"].number_format = FMT_CURRENCY
            ws[f"{col}{row}"].alignment = Alignment(horizontal="right", vertical="center")
        # Change col uses signed-delta format (red brackets for negatives).
        ws[f"D{row}"].number_format = FMT_CURRENCY_DELTA
        ws[f"D{row}"].alignment = Alignment(horizontal="right", vertical="center")

        # v2.3: light borders on all subtotal cells.
        for col in ("A", "B", "C", "D"):
            ws[f"{col}{row}"].border = subtotal_border

        # v2.4 FB-4: taller rows (was 20) so wrapped labels don't clip
        # and currency values breathe.
        ws.row_dimensions[row].height = 22

        # Emphasise the total-burden row
        if row == SUBTOTAL_BURDEN_ROW:
            top_border = Border(
                left=subtotal_border.left,
                right=subtotal_border.right,
                top=Side(style="medium", color="1D3B34"),
                bottom=subtotal_border.bottom,
            )
            for col in ("A", "B", "C", "D"):
                cell = ws[f"{col}{row}"]
                cell.font = TOTAL_BURDEN_FONT
                cell.fill = TOTAL_BURDEN_FILL
                cell.border = top_border
        else:
            ws[f"A{row}"].font = BODY_FONT

    # v2.5 FB-5: dropped the green-savings / red-cost conditional formatting.
    # Sign convention is now consistent across the tab (Change = reset − default,
    # negative = saving), and the [Red] brackets in FMT_CURRENCY_DELTA give the
    # eye the same signal without per-cell CF.

    # v2.3 C-3: short italic explanation immediately under the subtotals block.
    note_row = SUBTOTAL_BURDEN_ROW + 1
    note = ws.cell(
        row=note_row, column=1,
        value=("Change = If elected to reset − If no reset. A negative value "
               "(shown in red brackets) means the reset reduces that line of "
               "tax. Hover over each subtotal label for a plain-English "
               "definition."),
    )
    note.font = Font(name="Arial", size=9, italic=True, color="666666")
    note.alignment = Alignment(horizontal="left", vertical="center",
                               wrap_text=True, indent=1)
    ws.merge_cells(f"A{note_row}:{LAST_VISIBLE_COL}{note_row}")
    ws.row_dimensions[note_row].height = 16


def _build_per_member_breakdown(
    ws: Worksheet, headline_a: str, headline_b: str,
) -> None:
    """v2.3: 4-row per-member breakdown showing each member's TSB and their
    share of the Div 296 tax under both scenarios. Empty members render blank
    so the block degrades gracefully for single-member or 2-member funds.

    Data already exists in hidden cols L/M (per-member tax helpers); this block
    just surfaces it visibly.
    """
    _band(ws, PER_MEMBER_BAND_ROW, "Per-member breakdown")

    # Header row — v2.5 FB-6: labels standardised across all sections.
    headers = [
        ("A", "Member"),
        ("B", "TSB"),
        ("C", "If no reset (default)"),
        ("D", "If elected to reset"),
        ("E", "Change"),
    ]
    for col, text in headers:
        c = ws[f"{col}{PER_MEMBER_HEADER_ROW}"]
        c.value = text
        c.font = SECTION_BAND_FONT
        c.fill = SECTION_BAND_FILL
        c.alignment = CENTER

    # Data rows — one per member. v2.5 FB-6: member labels always show as
    # placeholders ("Member 1".."Member 4") regardless of TSB, so an empty
    # Inputs page still presents the 4-member structure on Comparison.
    # Numeric cells stay blank when TSB is empty/zero.
    for i in range(ASSUMPTIONS.member_count):
        row = PER_MEMBER_FIRST_ROW + i
        inputs_row = MEMBERS_FIRST_DATA_ROW + i
        helper_row = HELPER_MEMBER_TAX_FIRST_ROW + i
        tsb_ref = f"{INPUTS_SHEET}!B{inputs_row}"
        tax_a_ref = f"{HELPER_COL_A}{helper_row}"
        tax_b_ref = f"{HELPER_COL_B}{helper_row}"

        # Col A — Member label, ALWAYS shown (placeholder for empty members).
        label = ws.cell(row=row, column=1, value=f"Member {i+1}")
        label.font = BODY_FONT
        label.alignment = Alignment(horizontal="left", indent=1)

        # Col B — TSB
        tsb = ws.cell(row=row, column=2,
                      value=f'=IF({tsb_ref}>0,{tsb_ref},"")')
        tsb.number_format = FMT_CURRENCY
        tsb.alignment = Alignment(horizontal="right")

        # Col C — If no reset (default) Div 296 tax
        ca = ws.cell(row=row, column=3,
                     value=f'=IF({tsb_ref}>0,{tax_a_ref},"")')
        ca.number_format = FMT_CURRENCY
        ca.alignment = Alignment(horizontal="right")

        # Col D — If elected to reset Div 296 tax
        cb = ws.cell(row=row, column=4,
                     value=f'=IF({tsb_ref}>0,{tax_b_ref},"")')
        cb.number_format = FMT_CURRENCY
        cb.alignment = Alignment(horizontal="right")

        # Col E — Change SIGNED (reset − default; negative = saving).
        chg = ws.cell(row=row, column=5,
                      value=f'=IF({tsb_ref}>0,{tax_b_ref}-{tax_a_ref},"")')
        chg.number_format = FMT_CURRENCY_DELTA
        chg.alignment = Alignment(horizontal="right")

        ws.row_dimensions[row].height = 18


def _build_per_asset_detail(
    ws: Worksheet, gain_a_range: str, gain_b_range: str, delta_range: str,
    headline_a: str, headline_b: str,
) -> None:
    # v2.5 FB-7: band header reworded — emphasises tax impact (the new sort
    # metric, displayed in the Change column).
    _band(ws, BAND_DETAIL_ROW,
          f"Top {DISPLAY_ROWS} assets — Div 296 tax impact if you elect to reset")

    # Panel titles row — v2.5 FB-7: layout is now Panel A | Panel B | Change.
    panel_a_first, panel_a_last = PANEL_A_COLS[0], PANEL_A_COLS[-1]
    panel_b_first, panel_b_last = PANEL_B_COLS[0], PANEL_B_COLS[-1]
    panel_a_title = ws[f"{panel_a_first}{PANEL_TITLE_ROW}"]
    panel_a_title.value = "If no reset (default)"
    panel_a_title.font = SECTION_BAND_FONT
    panel_a_title.fill = SECTION_BAND_FILL
    panel_a_title.alignment = CENTER
    ws.merge_cells(f"{panel_a_first}{PANEL_TITLE_ROW}:{panel_a_last}{PANEL_TITLE_ROW}")

    panel_b_title = ws[f"{panel_b_first}{PANEL_TITLE_ROW}"]
    panel_b_title.value = "If elected to reset"
    panel_b_title.font = SECTION_BAND_FONT
    panel_b_title.fill = SECTION_BAND_FILL
    panel_b_title.alignment = CENTER
    ws.merge_cells(f"{panel_b_first}{PANEL_TITLE_ROW}:{panel_b_last}{PANEL_TITLE_ROW}")

    delta_title = ws[f"{DELTA_COL}{PANEL_TITLE_ROW}"]
    delta_title.value = "Change"
    delta_title.font = DELTA_FONT
    delta_title.fill = DELTA_HEADER_FILL
    delta_title.alignment = CENTER

    # Column sub-headers (29)
    sub_headers = ["Asset", "Proceeds", "Div 296 cost base",
                   "Div 296 adj gain", "Div 296 tax"]
    for col, header in zip(PANEL_A_COLS, sub_headers):
        c = ws[f"{col}{PANEL_HEADER_ROW}"]
        c.value = header
        c.font = SECTION_BAND_FONT
        c.fill = SECTION_BAND_FILL
        c.alignment = CENTER
    for col, header in zip(PANEL_B_COLS, sub_headers):
        c = ws[f"{col}{PANEL_HEADER_ROW}"]
        c.value = header
        c.font = SECTION_BAND_FONT
        c.fill = SECTION_BAND_FILL
        c.alignment = CENTER
    # v2.5 FB-7: sub-header reflects the new tax-delta metric (was "gain change").
    delta_sub = ws[f"{DELTA_COL}{PANEL_HEADER_ROW}"]
    delta_sub.value = "Div 296 tax"
    delta_sub.font = DELTA_FONT
    delta_sub.fill = DELTA_HEADER_FILL
    delta_sub.alignment = CENTER

    # Per-row data formulas — sort by |Δ| descending via LARGE / MATCH / INDEX.
    # Empty register rows score -1 on the delta range; LARGE pushes them past
    # rank DISPLAY_ROWS so matched=0 → blank cells in the visible panel.
    asset_a, proc_a, cb_a, gain_a, tax_a = PANEL_A_COLS
    asset_b, proc_b, cb_b, gain_b, tax_b = PANEL_B_COLS

    for offset in range(DISPLAY_ROWS):
        c_row = DATA_FIRST_ROW + offset
        k = offset + 1   # rank 1..DISPLAY_ROWS
        matched = f"${MATCHED_ROW_COL}{c_row}"   # absolute col, this row

        # Hidden: the register row that fills this visible slot. 0 = no asset
        # (fewer assets than DISPLAY_ROWS, or LARGE landed on an empty marker).
        ws[f"{MATCHED_ROW_COL}{c_row}"] = (
            f'=IFERROR('
            f'IF(LARGE({delta_range},{k})<=0,0,'
            f'MATCH(LARGE({delta_range},{k}),{delta_range},0)+{REGISTER_FIRST_DATA_ROW - 1}),'
            f'0)'
        )

        def _input(col: str) -> str:
            return f"INDEX({INPUTS_SHEET}!${col}:${col},{matched})"

        # v2.3: Inputs col letters shifted — D→C (orig), F→E (MV), H→G (proceeds)
        a_code  = _input("A")
        a_name  = _input("B")
        a_orig  = _input("C")   # was D
        a_mv    = _input("E")   # was F
        a_proc  = _input("G")   # was H

        # Pre-computed gains from the per-register grid (cols N/O).
        gain_a_lookup = f"INDEX(${PER_REG_GAIN_A_COL}:${PER_REG_GAIN_A_COL},{matched})"
        gain_b_lookup = f"INDEX(${PER_REG_GAIN_B_COL}:${PER_REG_GAIN_B_COL},{matched})"

        # Panel A (no reset → cost base = original)
        # v2.4 FB-2: "{code} - {name}" (was "{name} ({code})")
        ws[f"{asset_a}{c_row}"] = f'=IF({matched}=0,"",{a_code}&" - "&{a_name})'
        ws[f"{proc_a}{c_row}"]  = f'=IF({matched}=0,"",{a_proc})'
        ws[f"{cb_a}{c_row}"]    = f'=IF({matched}=0,"",{a_orig})'
        ws[f"{gain_a}{c_row}"]  = f'=IF({matched}=0,"",{gain_a_lookup})'
        ws[f"{tax_a}{c_row}"]   = (
            f'=IF({matched}=0,"",'
            f'IF(SUMIF({gain_a_range},">0")=0,0,'
            f'MAX(0,{gain_a}{c_row})/SUMIF({gain_a_range},">0")*{headline_a}))'
        )

        # Panel B (reset elected → cost base = MV)
        # v2.4 FB-2: "{code} - {name}" (was "{name} ({code})")
        ws[f"{asset_b}{c_row}"] = f'=IF({matched}=0,"",{a_code}&" - "&{a_name})'
        ws[f"{proc_b}{c_row}"]  = f'=IF({matched}=0,"",{a_proc})'
        ws[f"{cb_b}{c_row}"]    = f'=IF({matched}=0,"",{a_mv})'
        ws[f"{gain_b}{c_row}"]  = f'=IF({matched}=0,"",{gain_b_lookup})'
        ws[f"{tax_b}{c_row}"]   = (
            f'=IF({matched}=0,"",'
            f'IF(SUMIF({gain_b_range},">0")=0,0,'
            f'MAX(0,{gain_b}{c_row})/SUMIF({gain_b_range},">0")*{headline_b}))'
        )

        # v2.5 FB-7: Change column at the FAR RIGHT, signed TAX delta
        # (reset tax − default tax). Negative = reset reduces tax (red brackets).
        ws[f"{DELTA_COL}{c_row}"] = (
            f'=IF({matched}=0,"",{tax_b}{c_row}-{tax_a}{c_row})'
        )

        # Currency formatting on all numeric cells.
        for col in (proc_a, cb_a, gain_a, tax_a, proc_b, cb_b, gain_b, tax_b):
            ws[f"{col}{c_row}"].number_format = FMT_CURRENCY
        # Signed delta format for the Change col.
        ws[f"{DELTA_COL}{c_row}"].number_format = FMT_CURRENCY_DELTA

        # v2.3 C-4: per-column fills mirroring Analyser's per-asset table
        # (sand for proceeds, slate for ordinary/cost, sage for Div 296).
        # Applied to both Panel A and Panel B so the eye sweeps left-to-right
        # by data class, not by panel.
        ws[f"{proc_a}{c_row}"].fill = PROC_DATA_FILL
        ws[f"{cb_a}{c_row}"].fill = ORD_DATA_FILL
        ws[f"{gain_a}{c_row}"].fill = DIV_DATA_FILL
        ws[f"{tax_a}{c_row}"].fill = DIV_DATA_FILL
        ws[f"{proc_b}{c_row}"].fill = PROC_DATA_FILL
        ws[f"{cb_b}{c_row}"].fill = ORD_DATA_FILL
        ws[f"{gain_b}{c_row}"].fill = DIV_DATA_FILL
        ws[f"{tax_b}{c_row}"].fill = DIV_DATA_FILL

    # E3: highlight Div 296 cost base in panel B when it differs from panel A.
    # CF formula uses RELATIVE rows so it shifts per row in the range.
    cb_b_range = f"{PANEL_B_COLS[2]}{DATA_FIRST_ROW}:{PANEL_B_COLS[2]}{DATA_LAST_ROW}"
    diff_rule = FormulaRule(
        formula=[f'AND({PANEL_A_COLS[2]}{DATA_FIRST_ROW}<>"",'
                 f'{PANEL_B_COLS[2]}{DATA_FIRST_ROW}<>'
                 f'{PANEL_A_COLS[2]}{DATA_FIRST_ROW})'],
        fill=COST_BASE_ACCENT_FILL,
    )
    ws.conditional_formatting.add(cb_b_range, diff_rule)

    # v2.3 C-4: trap red-row CF — fires when the asset is in an unrealised
    # loss position (Panel A proceeds < cost base) AND the reset turns it
    # into a positive Div 296 gain (Panel B gain > 0). v2.5 FB-7: column
    # refs made symbolic (was hardcoded "$I" which referenced the wrong
    # column under the old layout; new symbolic refs auto-track the layout).
    panel_b_gain_col = PANEL_B_COLS[3]   # "I" under the v2.5 layout
    trap_range = (
        f"{PANEL_A_COLS[0]}{DATA_FIRST_ROW}:"
        f"{LAST_VISIBLE_COL}{DATA_LAST_ROW}"
    )
    trap_rule = FormulaRule(
        formula=[
            f"AND($A{DATA_FIRST_ROW}<>\"\","
            f"($B{DATA_FIRST_ROW}-$C{DATA_FIRST_ROW})<0,"
            f"${panel_b_gain_col}{DATA_FIRST_ROW}>0)"
        ],
        fill=TRAP_FILL,
    )
    ws.conditional_formatting.add(trap_range, trap_rule)

    # v2.5 FB-7: dropped the green/red CF on the Change column. The
    # FMT_CURRENCY_DELTA format renders negatives as red brackets, which is
    # the consistent visual language across every Change cell on the tab.

    # Overflow note — v2.5 FB-7: mentions tax delta (the new sort metric).
    overflow = ws.cell(
        row=DATA_OVERFLOW_NOTE_ROW, column=1,
        value=(f"Showing top {DISPLAY_ROWS} assets by absolute Div 296 tax "
               f"change when the reset is elected — see the Analyser tab for "
               f"the full register (up to {ASSUMPTIONS.asset_register_rows} rows)."),
    )
    overflow.font = Font(name="Arial", size=9, italic=True, color="666666")
    ws.merge_cells(f"A{DATA_OVERFLOW_NOTE_ROW}:{LAST_VISIBLE_COL}{DATA_OVERFLOW_NOTE_ROW}")


def _build_footer_notes(ws: Worksheet) -> None:
    reminder = ws.cell(
        row=REMINDER_ROW, column=1,
        value=("Note: assets currently in an unrealised-loss position may contribute "
               "Div 296 tax IF you elect the reset that they do not contribute under "
               "the default outcome — see the Analyser tab for per-asset detail."),
    )
    reminder.font = Font(name="Arial", size=9, italic=True, color="666666")
    reminder.alignment = Alignment(wrap_text=True, vertical="top")
    ws.merge_cells(f"A{REMINDER_ROW}:{LAST_VISIBLE_COL}{REMINDER_ROW}")
    ws.row_dimensions[REMINDER_ROW].height = 24

    sort_note = ws.cell(
        row=SORT_NOTE_ROW, column=1,
        value=("Note: per-asset detail shows the top 10 assets where the reset "
               "election moves the Div 296 tax the most (by absolute change, "
               "either direction). See the Analyser tab for the full register."),
    )
    sort_note.font = Font(name="Arial", size=9, italic=True, color="666666")
    sort_note.alignment = Alignment(wrap_text=True, vertical="top")
    ws.merge_cells(f"A{SORT_NOTE_ROW}:{LAST_VISIBLE_COL}{SORT_NOTE_ROW}")
    ws.row_dimensions[SORT_NOTE_ROW].height = 30


def build(wb: Workbook) -> Worksheet:
    ws = wb.create_sheet(SHEET)
    ws.sheet_view.showGridLines = False

    _build_header_block(ws)

    # Section band for the body
    _band(ws, BAND_BODY_ROW,
          "Side-by-side comparison (independent of the master reset toggle)")

    _build_context_strip(ws)

    gain_a_range, gain_b_range, delta_range = _build_per_register_helpers(ws)
    headline_a, headline_b = _build_helpers(ws, gain_a_range, gain_b_range)

    _build_metric_cards(ws, headline_a, headline_b)
    _build_subtotals(ws, headline_a, headline_b, gain_a_range, gain_b_range)
    _build_per_member_breakdown(ws, headline_a, headline_b)
    _build_per_asset_detail(
        ws, gain_a_range, gain_b_range, delta_range, headline_a, headline_b,
    )
    _build_footer_notes(ws)

    # --- Print header watermark (large grey text on every printed page) ---
    ws.oddHeader.center.text = "ILLUSTRATIVE — NOT ADVICE"
    ws.oddHeader.center.size = 28
    ws.oddHeader.center.color = "CCCCCC"

    # --- Print setup: A4 landscape, narrow margins; let height spill if needed ---
    ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.fitToWidth = 1
    # v2.5 FB-3: chart removed — single landscape A4 page is now enough.
    ws.page_setup.fitToHeight = 1
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.print_options.horizontalCentered = True
    ws.page_margins.left = 0.25
    ws.page_margins.right = 0.25
    ws.page_margins.top = 0.5
    ws.page_margins.bottom = 0.4
    ws.print_area = f"A1:{LAST_VISIBLE_COL}{SORT_NOTE_ROW}"

    # --- Column widths (v2.5 FB-7 layout) ---
    # A-E  Panel A + subtotals: A holds long labels; B-D currency values.
    # F-J  Panel B (asset, proceeds, cost base, gain, tax) — same shape as A.
    # K    Change column (tax delta) — rightmost.
    widths = {
        "A": 36, "B": 16, "C": 16, "D": 16, "E": 14,    # Panel A + subtotals
        "F": 26, "G": 14, "H": 16, "I": 16, "J": 14,    # Panel B
        "K": 14,                                         # Change
    }
    for col_letter, w in widths.items():
        ws.column_dimensions[col_letter].width = w

    # Editable header-block cells stay editable under sheet protection.
    for row in (HEADER_FIRM_ROW, HEADER_PREPARED_FOR_ROW,
                HEADER_PREPARED_BY_ROW, HEADER_DATE_ROW):
        ws[f"B{row}"].protection = Protection(locked=False)

    ws.protection.sheet = True
    ws.protection.formatColumns = False
    ws.protection.formatRows = False

    return ws
