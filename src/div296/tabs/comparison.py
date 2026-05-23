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

from openpyxl.chart import BarChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Protection, Side
from openpyxl.workbook import Workbook
from openpyxl.worksheet.worksheet import Worksheet

from div296.assumptions import ASSUMPTIONS
from div296.styles import (
    BODY_FONT, CENTER, FMT_CURRENCY, FMT_PERCENT, INPUT_FILL, INPUT_FONT,
    SECTION_BAND_FILL, SECTION_BAND_FONT, THIN_BOX, TITLE_FONT,
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
CONTEXT_VALUE_ROW = 14

BAND_HEADLINE_ROW = 16
CARD_LABEL_ROW = 17
CARD_VALUE_ROW = 18

BAND_SUBTOTALS_ROW = 20
SUBTOTAL_HEADER_ROW = 21
SUBTOTAL_EARNINGS_ROW = 22
SUBTOTAL_ORD_CGT_ROW = 23
SUBTOTAL_DIV296_ROW = 24
SUBTOTAL_BURDEN_ROW = 25

BAND_DETAIL_ROW = 27
PANEL_TITLE_ROW = 28
PANEL_HEADER_ROW = 29
DATA_FIRST_ROW = 30
DATA_LAST_ROW = DATA_FIRST_ROW + DISPLAY_ROWS - 1   # row 44
DATA_OVERFLOW_NOTE_ROW = DATA_LAST_ROW + 1          # row 45

REMINDER_ROW = DATA_OVERFLOW_NOTE_ROW + 1           # row 36
SORT_NOTE_ROW = REMINDER_ROW + 1                  # row 37

# v1.6: chart anchored inline next to the subtotals (cols F-K, rows 20-27)
# rather than below the data block. Fills the empty right-side region that
# the subtotals table (cols A-D) doesn't use, and gets the chart on page 1.
CHART_ANCHOR_ROW = BAND_SUBTOTALS_ROW                 # row 20
CHART_ANCHOR_COL = "F"
CHART_BOTTOM_ROW = SORT_NOTE_ROW                  # print area ends with the notes

# --- Layout columns ---
PANEL_A_COLS = ("A", "B", "C", "D", "E")   # Asset, Proceeds, Cost base, Adj gain, Tax
DELTA_COL = "F"                             # Per-asset Δ (panel B - panel A adj gain)
PANEL_B_COLS = ("G", "H", "I", "J", "K")   # same shape as panel A
LAST_VISIBLE_COL = PANEL_B_COLS[-1]         # "K"

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


def _build_context_strip(ws: Worksheet) -> None:
    """A small block showing the assumptions driving the headline numbers.
    Computed live from Inputs."""
    # Labels (row 13)
    labels = [
        ("A:C", "Member 1 TSB"),
        ("D:F", "Proportion above $3m"),
        ("G:H", "CGT discount"),
        ("I:K", "$10m / +25% tier"),
    ]
    for merge_range, label in labels:
        start = merge_range.split(":")[0]
        ws[f"{start}{CONTEXT_LABEL_ROW}"] = label
        ws[f"{start}{CONTEXT_LABEL_ROW}"].font = CONTEXT_LABEL_FONT
        ws[f"{start}{CONTEXT_LABEL_ROW}"].alignment = Alignment(horizontal="left", indent=1)
        ws.merge_cells(f"{merge_range.split(':')[0]}{CONTEXT_LABEL_ROW}:"
                       f"{merge_range.split(':')[1]}{CONTEXT_LABEL_ROW}")

    # Values (row 14) — live references to Inputs
    values = [
        ("A:C", f"={INPUTS_SHEET}!B{MEMBERS_FIRST_DATA_ROW}",                 FMT_CURRENCY),
        ("D:F", (f'=IF({INPUTS_SHEET}!E{MEMBERS_FIRST_DATA_ROW}="",'
                 f'{INPUTS_SHEET}!D{MEMBERS_FIRST_DATA_ROW},'
                 f'{INPUTS_SHEET}!E{MEMBERS_FIRST_DATA_ROW})'),               FMT_PERCENT),
        ("G:H", "=discount_on",                                               None),
        ("I:K", "=tier10_on",                                                 None),
    ]
    for merge_range, formula, fmt in values:
        start = merge_range.split(":")[0]
        end = merge_range.split(":")[1]
        cell = ws[f"{start}{CONTEXT_VALUE_ROW}"]
        cell.value = formula
        cell.font = CONTEXT_VALUE_FONT
        cell.alignment = Alignment(horizontal="left", indent=1)
        if fmt:
            cell.number_format = fmt
        ws.merge_cells(f"{start}{CONTEXT_VALUE_ROW}:{end}{CONTEXT_VALUE_ROW}")

    ws.row_dimensions[CONTEXT_VALUE_ROW].height = 22


def _build_per_register_helpers(ws: Worksheet) -> tuple[str, str, str]:
    """Per-register-row helper grid (cols N/O/P) used to sort the per-asset
    detail panel by |Δ| descending.

    N = Scenario A Div 296 adj gain (cost base = original)
    O = Scenario B Div 296 adj gain (cost base = MV)
    P = |Δ| + row-based tiebreaker (so MATCH uniquely identifies the row);
        empty register rows return -1 so LARGE pushes them to the bottom.

    Returns the absolute ranges (e.g. '$N$20:$N$69').
    """
    n_first, n_last = REGISTER_FIRST_DATA_ROW, REGISTER_LAST_DATA_ROW
    for n in range(n_first, n_last + 1):
        proceeds = f"{INPUTS_SHEET}!H{n}"
        orig = f"{INPUTS_SHEET}!D{n}"
        mv = f"{INPUTS_SHEET}!F{n}"
        held = f"{INPUTS_SHEET}!I{n}"
        ws[f"{PER_REG_GAIN_A_COL}{n}"] = _div296_adj_formula(proceeds, orig, held)
        ws[f"{PER_REG_GAIN_B_COL}{n}"] = _div296_adj_formula(proceeds, mv, held)
        # Tiebreaker: a small positive amount that decreases with register row,
        # so earlier-listed assets win ties without ever turning P negative.
        tiebreak = f"({n_last}-ROW())*0.001"
        ws[f"{PER_REG_DELTA_COL}{n}"] = (
            f'=IF({proceeds}="",-1,'
            f'ABS({PER_REG_GAIN_B_COL}{n}-{PER_REG_GAIN_A_COL}{n})+{tiebreak})'
        )
    return (
        f"${PER_REG_GAIN_A_COL}${n_first}:${PER_REG_GAIN_A_COL}${n_last}",
        f"${PER_REG_GAIN_B_COL}${n_first}:${PER_REG_GAIN_B_COL}${n_last}",
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

    cards = [
        ("A:D", "Scenario A — No reset",     f"={headline_a[1:]}",                     CARD_FILL_A),
        ("E:H", "Scenario B — Reset elected", f"={headline_b[1:]}",                    CARD_FILL_B),
        ("I:K", "Net effect (A − B)",        f"={headline_a[1:]}-{headline_b[1:]}",   CARD_FILL_DELTA),
    ]
    for merge_range, label, value_formula, fill in cards:
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
        value_cell.number_format = FMT_CURRENCY
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


def _expand_cols(start: str, end: str) -> list[str]:
    return [chr(c) for c in range(ord(start), ord(end) + 1)]


def _build_subtotals(ws: Worksheet, headline_a: str, headline_b: str,
                     gain_a_range: str, gain_b_range: str) -> None:
    """Four-row subtotal table: earnings, ord CGT, Div 296 tax, total burden."""
    _band(ws, BAND_SUBTOTALS_ROW, "Per-scenario subtotals")

    # Header row
    header_cells = [
        ("A", "Subtotal"),
        ("B", "Scenario A"),
        ("C", "Scenario B"),
        ("D", "Δ (B − A)"),
    ]
    for col, text in header_cells:
        c = ws[f"{col}{SUBTOTAL_HEADER_ROW}"]
        c.value = text
        c.font = SECTION_BAND_FONT
        c.fill = SECTION_BAND_FILL
        c.alignment = CENTER

    # Reference to the Analyser's Ordinary CGT total — this is the same in both
    # scenarios (ordinary CGT doesn't depend on reset election).
    # Analyser B73 = Ordinary CGT payable (reconciliation panel).
    ord_cgt_ref = f"={ANALYSER_SHEET}!B73"

    rows = [
        (SUBTOTAL_EARNINGS_ROW, "Div 296 earnings",
         f"=SUMIF({gain_a_range},\">0\")", f"=SUMIF({gain_b_range},\">0\")"),
        (SUBTOTAL_ORD_CGT_ROW,  "Ordinary CGT (unchanged by reset)",
         ord_cgt_ref, ord_cgt_ref),
        (SUBTOTAL_DIV296_ROW,   "Div 296 tax (headline)",
         f"={headline_a[1:]}", f"={headline_b[1:]}"),
        (SUBTOTAL_BURDEN_ROW,   "TOTAL TAX BURDEN",
         f"=B{SUBTOTAL_ORD_CGT_ROW}+B{SUBTOTAL_DIV296_ROW}",
         f"=C{SUBTOTAL_ORD_CGT_ROW}+C{SUBTOTAL_DIV296_ROW}"),
    ]
    for row, label, a_val, b_val in rows:
        ws[f"A{row}"] = label
        ws[f"B{row}"] = a_val
        ws[f"C{row}"] = b_val
        ws[f"D{row}"] = f"=B{row}-C{row}"
        ws[f"A{row}"].alignment = Alignment(wrap_text=True, vertical="center")
        for col in ("B", "C", "D"):
            ws[f"{col}{row}"].number_format = FMT_CURRENCY
            ws[f"{col}{row}"].alignment = Alignment(horizontal="right", vertical="center")

        # Slightly taller rows so wrapped labels don't clip.
        ws.row_dimensions[row].height = 20

        # Emphasise the total-burden row
        if row == SUBTOTAL_BURDEN_ROW:
            for col in ("A", "B", "C", "D"):
                cell = ws[f"{col}{row}"]
                cell.font = TOTAL_BURDEN_FONT
                cell.fill = TOTAL_BURDEN_FILL
        else:
            ws[f"A{row}"].font = BODY_FONT


def _build_per_asset_detail(
    ws: Worksheet, gain_a_range: str, gain_b_range: str, delta_range: str,
    headline_a: str, headline_b: str,
) -> None:
    _band(ws, BAND_DETAIL_ROW,
          f"Per-asset detail — top {DISPLAY_ROWS} by |Δ (B − A)|")

    # Panel titles row (28)
    panel_a_first, panel_a_last = PANEL_A_COLS[0], PANEL_A_COLS[-1]
    panel_b_first, panel_b_last = PANEL_B_COLS[0], PANEL_B_COLS[-1]
    panel_a_title = ws[f"{panel_a_first}{PANEL_TITLE_ROW}"]
    panel_a_title.value = "Scenario A — No reset"
    panel_a_title.font = SECTION_BAND_FONT
    panel_a_title.fill = SECTION_BAND_FILL
    panel_a_title.alignment = CENTER
    ws.merge_cells(f"{panel_a_first}{PANEL_TITLE_ROW}:{panel_a_last}{PANEL_TITLE_ROW}")

    delta_title = ws[f"{DELTA_COL}{PANEL_TITLE_ROW}"]
    delta_title.value = "Δ"
    delta_title.font = DELTA_FONT
    delta_title.fill = DELTA_HEADER_FILL
    delta_title.alignment = CENTER

    panel_b_title = ws[f"{panel_b_first}{PANEL_TITLE_ROW}"]
    panel_b_title.value = "Scenario B — Reset elected"
    panel_b_title.font = SECTION_BAND_FONT
    panel_b_title.fill = SECTION_BAND_FILL
    panel_b_title.alignment = CENTER
    ws.merge_cells(f"{panel_b_first}{PANEL_TITLE_ROW}:{panel_b_last}{PANEL_TITLE_ROW}")

    # Column sub-headers (29)
    sub_headers = ["Asset", "Proceeds", "Div 296 cost base",
                   "Div 296 adj gain", "Div 296 tax"]
    for col, header in zip(PANEL_A_COLS, sub_headers):
        c = ws[f"{col}{PANEL_HEADER_ROW}"]
        c.value = header
        c.font = SECTION_BAND_FONT
        c.fill = SECTION_BAND_FILL
        c.alignment = CENTER
    delta_sub = ws[f"{DELTA_COL}{PANEL_HEADER_ROW}"]
    delta_sub.value = "gain (B − A)"
    delta_sub.font = DELTA_FONT
    delta_sub.fill = DELTA_HEADER_FILL
    delta_sub.alignment = CENTER
    for col, header in zip(PANEL_B_COLS, sub_headers):
        c = ws[f"{col}{PANEL_HEADER_ROW}"]
        c.value = header
        c.font = SECTION_BAND_FONT
        c.fill = SECTION_BAND_FILL
        c.alignment = CENTER

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

        a_code  = _input("A")
        a_name  = _input("B")
        a_orig  = _input("D")
        a_mv    = _input("F")
        a_proc  = _input("H")

        # Pre-computed gains from the per-register grid (cols N/O).
        gain_a_lookup = f"INDEX(${PER_REG_GAIN_A_COL}:${PER_REG_GAIN_A_COL},{matched})"
        gain_b_lookup = f"INDEX(${PER_REG_GAIN_B_COL}:${PER_REG_GAIN_B_COL},{matched})"

        # Panel A (no reset → cost base = original)
        ws[f"{asset_a}{c_row}"] = f'=IF({matched}=0,"",{a_name}&" ("&{a_code}&")")'
        ws[f"{proc_a}{c_row}"]  = f'=IF({matched}=0,"",{a_proc})'
        ws[f"{cb_a}{c_row}"]    = f'=IF({matched}=0,"",{a_orig})'
        ws[f"{gain_a}{c_row}"]  = f'=IF({matched}=0,"",{gain_a_lookup})'
        ws[f"{tax_a}{c_row}"]   = (
            f'=IF({matched}=0,"",'
            f'IF(SUMIF({gain_a_range},">0")=0,0,'
            f'MAX(0,{gain_a}{c_row})/SUMIF({gain_a_range},">0")*{headline_a}))'
        )

        # Δ column = (panel B adj gain) − (panel A adj gain)
        ws[f"{DELTA_COL}{c_row}"] = (
            f'=IF({matched}=0,"",{gain_b}{c_row}-{gain_a}{c_row})'
        )

        # Panel B (reset elected → cost base = MV)
        ws[f"{asset_b}{c_row}"] = f'=IF({matched}=0,"",{a_name}&" ("&{a_code}&")")'
        ws[f"{proc_b}{c_row}"]  = f'=IF({matched}=0,"",{a_proc})'
        ws[f"{cb_b}{c_row}"]    = f'=IF({matched}=0,"",{a_mv})'
        ws[f"{gain_b}{c_row}"]  = f'=IF({matched}=0,"",{gain_b_lookup})'
        ws[f"{tax_b}{c_row}"]   = (
            f'=IF({matched}=0,"",'
            f'IF(SUMIF({gain_b_range},">0")=0,0,'
            f'MAX(0,{gain_b}{c_row})/SUMIF({gain_b_range},">0")*{headline_b}))'
        )

        # Currency formatting on all numeric cells.
        for col in (proc_a, cb_a, gain_a, tax_a, DELTA_COL, proc_b, cb_b, gain_b, tax_b):
            ws[f"{col}{c_row}"].number_format = FMT_CURRENCY

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

    # Overflow note
    overflow = ws.cell(
        row=DATA_OVERFLOW_NOTE_ROW, column=1,
        value=(f"Showing top {DISPLAY_ROWS} assets by |Δ (B − A)| — see the "
               f"Analyser tab for the full register "
               f"(up to {ASSUMPTIONS.asset_register_rows} rows)."),
    )
    overflow.font = Font(name="Arial", size=9, italic=True, color="666666")
    ws.merge_cells(f"A{DATA_OVERFLOW_NOTE_ROW}:{LAST_VISIBLE_COL}{DATA_OVERFLOW_NOTE_ROW}")


def _build_footer_notes(ws: Worksheet) -> None:
    reminder = ws.cell(
        row=REMINDER_ROW, column=1,
        value=("Note: loss-position assets may contribute Div 296 tax under Scenario B "
               "that they do not under Scenario A — see the Analyser tab for per-asset detail."),
    )
    reminder.font = Font(name="Arial", size=9, italic=True, color="666666")
    reminder.alignment = Alignment(wrap_text=True, vertical="top")
    ws.merge_cells(f"A{REMINDER_ROW}:{LAST_VISIBLE_COL}{REMINDER_ROW}")
    ws.row_dimensions[REMINDER_ROW].height = 24

    sort_note = ws.cell(
        row=SORT_NOTE_ROW, column=1,
        value=("Note: per-asset detail shows the top 10 assets by absolute "
               "Δ (Scenario B − Scenario A) — i.e. those where the reset election "
               "moves the Div 296 gain the most. See the Analyser tab for the full register."),
    )
    sort_note.font = Font(name="Arial", size=9, italic=True, color="666666")
    sort_note.alignment = Alignment(wrap_text=True, vertical="top")
    ws.merge_cells(f"A{SORT_NOTE_ROW}:{LAST_VISIBLE_COL}{SORT_NOTE_ROW}")
    ws.row_dimensions[SORT_NOTE_ROW].height = 30


def _build_chart(ws: Worksheet, *_unused) -> None:
    """v2.0.0: horizontal bar chart of per-asset Δ values (col F of the
    visible per-asset detail block, rows DATA_FIRST_ROW..DATA_LAST_ROW).

    Each bar = one displayed asset; length = (Scenario B − Scenario A)
    gain delta; sorted descending (table already sorts this way via
    LARGE/MATCH). Anchored below the footer notes, full panel width."""
    chart = BarChart()
    chart.type = "bar"          # horizontal
    chart.style = 11
    chart.title = "Per-asset Δ (Scenario B − Scenario A) — top 10 by |Δ|"
    chart.y_axis.title = None
    chart.x_axis.title = "Δ ($)"
    chart.legend = None
    # belt-and-braces; source is visible anyway. openpyxl exposes the
    # plotVisOnly chartSpace flag as `display_blanks`/`visible_cells_only`.
    chart.visible_cells_only = False

    # Data = col F (DELTA_COL), rows DATA_FIRST_ROW..DATA_LAST_ROW.
    # Categories = col A (Asset name + code), same rows.
    data = Reference(
        ws,
        min_col=ord(DELTA_COL) - ord("A") + 1,
        max_col=ord(DELTA_COL) - ord("A") + 1,
        min_row=DATA_FIRST_ROW,
        max_row=DATA_LAST_ROW,
    )
    cats = Reference(
        ws,
        min_col=1,                      # col A — Asset
        max_col=1,
        min_row=DATA_FIRST_ROW,
        max_row=DATA_LAST_ROW,
    )
    chart.add_data(data, titles_from_data=False)
    chart.set_categories(cats)
    chart.dataLabels = DataLabelList(
        showVal=True, showCatName=False, showSerName=False, showLegendKey=False,
    )
    # Sized to fit the full visible panel width (A:K).
    chart.height = 9
    chart.width = 22

    # Anchor BELOW the existing footer notes block (CHART_BOTTOM_ROW + 2),
    # full panel width.
    ws.add_chart(chart, f"A{CHART_BOTTOM_ROW + 2}")


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
    _build_per_asset_detail(
        ws, gain_a_range, gain_b_range, delta_range, headline_a, headline_b,
    )
    _build_footer_notes(ws)
    _build_chart(ws)

    # --- Print header watermark (large grey text on every printed page) ---
    ws.oddHeader.center.text = "ILLUSTRATIVE — NOT ADVICE"
    ws.oddHeader.center.size = 28
    ws.oddHeader.center.color = "CCCCCC"

    # --- Print setup: A4 landscape, narrow margins; let height spill if needed ---
    ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.fitToWidth = 1
    # v2.0.0: allow second page for chart (per-asset detail + chart together
    # can exceed one landscape A4).
    ws.page_setup.fitToHeight = 2
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.print_options.horizontalCentered = True
    ws.page_margins.left = 0.25
    ws.page_margins.right = 0.25
    ws.page_margins.top = 0.5
    ws.page_margins.bottom = 0.4
    CHART_PRINT_BOTTOM = CHART_BOTTOM_ROW + 22       # 2-row gap + ~20-row chart
    ws.print_area = f"A1:{LAST_VISIBLE_COL}{CHART_PRINT_BOTTOM}"

    # --- Column widths ---
    # Panel A: Asset 24 / Proceeds 13 / Cost base 14 / Adj gain 14 / Tax 12
    # Δ: 12
    # Panel B: same as A
    widths = {
        "A": 28, "B": 12, "C": 13, "D": 13, "E": 12,    # Panel A (col A wide enough for subtotal labels)
        "F": 11,                                         # Δ
        "G": 22, "H": 12, "I": 13, "J": 13, "K": 12,    # Panel B
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
