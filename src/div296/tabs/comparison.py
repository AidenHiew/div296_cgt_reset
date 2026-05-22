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
    Rows 47-48   Existing reminder + manual-earnings notes
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
from openpyxl.chart.data_source import (
    AxDataSource, NumData, NumDataSource, NumRef, NumVal,
    StrData, StrRef, StrVal,
)
from openpyxl.chart.label import DataLabelList
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Protection, Side
from openpyxl.workbook import Workbook
from openpyxl.worksheet.worksheet import Worksheet

from div296 import calcs
from div296.assumptions import ASSUMPTIONS
from div296.styles import (
    BODY_FONT, CENTER, FMT_CURRENCY, FMT_PERCENT, INPUT_FILL, INPUT_FONT,
    SECTION_BAND_FILL, SECTION_BAND_FONT, THIN_BOX, TITLE_FONT,
)
from div296.tabs.inputs import (
    MEMBERS_FIRST_DATA_ROW, REGISTER_FIRST_DATA_ROW, SAMPLE_REGISTER_ROWS,
)


SHEET = "Comparison"
INPUTS_SHEET = "'Inputs'"
ANALYSER_SHEET = "'Analyser'"

# How many assets to display on the printable comparison page.
# Reduced from 50 (Inputs register capacity) to a tight tearsheet view.
DISPLAY_ROWS = 15

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

REMINDER_ROW = DATA_OVERFLOW_NOTE_ROW + 2           # row 47
MANUAL_NOTE_ROW = REMINDER_ROW + 1                  # row 48

CHART_ANCHOR_ROW = MANUAL_NOTE_ROW + 2              # row 50
CHART_BOTTOM_ROW = CHART_ANCHOR_ROW + 16            # ~16 rows tall

# --- Layout columns ---
PANEL_A_COLS = ("A", "B", "C", "D", "E")   # Asset, Proceeds, Cost base, Adj gain, Tax
DELTA_COL = "F"                             # Per-asset Δ (panel B - panel A adj gain)
PANEL_B_COLS = ("G", "H", "I", "J", "K")   # same shape as panel A
LAST_VISIBLE_COL = PANEL_B_COLS[-1]         # "K"

# Hidden helper columns (right of the visible panels)
HELPER_COL_A = "L"   # Scenario A working column
HELPER_COL_B = "M"   # Scenario B working column

# Hidden helper rows in cols L/M
HELPER_FUND_EARNINGS_ROW = 1
HELPER_MEMBER_TAX_FIRST_ROW = 2
HELPER_MEMBER_TAX_LAST_ROW = HELPER_MEMBER_TAX_FIRST_ROW + ASSUMPTIONS.member_count - 1
HELPER_HEADLINE_ROW = HELPER_MEMBER_TAX_LAST_ROW + 1   # row 6

# Hidden chart data block (cols L/M, rows 8-9 — labels and computed values)
CHART_LABEL_A_CELL = "L8"
CHART_LABEL_B_CELL = "L9"
CHART_VALUE_A_CELL = "M8"
CHART_VALUE_B_CELL = "M9"

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


# --- Sample-data computation for chart cache + sanity ---

def _sample_scenario_headlines() -> tuple[float, float]:
    """Compute the Scenario A / Scenario B Div 296 headline tax from the
    sample register baked into Inputs. Used to pre-populate the chart's
    cache so headless PDF renderers draw the bars."""
    assets = [
        calcs.Asset(
            code=r[0], name=r[1], quantity=r[2], original_cost_base=r[3],
            total_value=r[4], market_value_30jun2026=r[5],
            valuation_source=r[6], projected_sale_proceeds=r[7],
            held_over_12_months=(r[8] == "Yes"),
        )
        for r in SAMPLE_REGISTER_ROWS
    ]
    member = calcs.Member(tsb=12_000_000.0, split_pct=1.0)
    kwargs = dict(
        assets=assets,
        members=[member],
        discount_on=True,
        discount_rate=ASSUMPTIONS.discount_rate,
        tier10_on=False,
        threshold_1=ASSUMPTIONS.threshold_1,
        threshold_2=ASSUMPTIONS.threshold_2,
        rate_tier1=ASSUMPTIONS.rate_tier1,
        rate_tier2=ASSUMPTIONS.rate_tier2,
    )
    scenario_a = calcs.div296_headline_tax(reset_on=False, **kwargs)
    scenario_b = calcs.div296_headline_tax(reset_on=True, **kwargs)
    return scenario_a, scenario_b


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


def _build_helpers(ws: Worksheet) -> tuple[str, str, str, str]:
    """Hidden helper block in cols L/M. Returns (gain_a_range, gain_b_range,
    headline_a_cell, headline_b_cell) for downstream wiring."""
    # Scenario A panel fund earnings = sum of positive col D values
    panel_a_gain_col = PANEL_A_COLS[3]   # col D = adj gain
    panel_b_gain_col = PANEL_B_COLS[3]   # col J = adj gain
    gain_a_range = f"{panel_a_gain_col}{DATA_FIRST_ROW}:{panel_a_gain_col}{DATA_LAST_ROW}"
    gain_b_range = f"{panel_b_gain_col}{DATA_FIRST_ROW}:{panel_b_gain_col}{DATA_LAST_ROW}"

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

    # Chart data labels & values (hidden)
    ws[CHART_LABEL_A_CELL] = "Scenario A — No reset"
    ws[CHART_LABEL_B_CELL] = "Scenario B — Reset elected"
    ws[CHART_VALUE_A_CELL] = f"={headline_a_cell}"
    ws[CHART_VALUE_B_CELL] = f"={headline_b_cell}"
    for coord in (CHART_VALUE_A_CELL, CHART_VALUE_B_CELL):
        ws[coord].number_format = FMT_CURRENCY

    # Hide helper columns + anything to the right
    for col_letter in (HELPER_COL_A, HELPER_COL_B, "N", "O", "P", "Q"):
        ws.column_dimensions[col_letter].hidden = True

    # Return absolute refs for downstream formulas (e.g. "$L$6").
    abs_a = f"${HELPER_COL_A}${HELPER_HEADLINE_ROW}"
    abs_b = f"${HELPER_COL_B}${HELPER_HEADLINE_ROW}"
    return gain_a_range, gain_b_range, abs_a, abs_b


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
    ws: Worksheet, gain_a_range: str, gain_b_range: str,
    headline_a: str, headline_b: str,
) -> None:
    _band(ws, BAND_DETAIL_ROW,
          f"Per-asset detail (first {DISPLAY_ROWS} assets)")

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

    # Per-row data formulas (only the first DISPLAY_ROWS rows are rendered)
    for offset in range(DISPLAY_ROWS):
        c_row = DATA_FIRST_ROW + offset
        i_row = REGISTER_FIRST_DATA_ROW + offset

        code = f"{INPUTS_SHEET}!A{i_row}"
        name = f"{INPUTS_SHEET}!B{i_row}"
        orig = f"{INPUTS_SHEET}!D{i_row}"
        mv = f"{INPUTS_SHEET}!F{i_row}"
        proceeds = f"{INPUTS_SHEET}!H{i_row}"
        held = f"{INPUTS_SHEET}!I{i_row}"

        # Panel A (no reset → cost base = original)
        asset_a, proc_a, cb_a, gain_a, tax_a = PANEL_A_COLS
        ws[f"{asset_a}{c_row}"] = f'=IF({code}="","",{name}&" ("&{code}&")")'
        ws[f"{proc_a}{c_row}"] = f'=IF({proceeds}="","",{proceeds})'
        ws[f"{cb_a}{c_row}"]   = f'=IF({proceeds}="","",{orig})'
        ws[f"{gain_a}{c_row}"] = _div296_adj_formula(proceeds, orig, held)
        ws[f"{tax_a}{c_row}"]  = (
            f'=IF({proceeds}="","",'
            f'IF(SUMIF({gain_a_range},">0")=0,0,'
            f'MAX(0,{gain_a}{c_row})/SUMIF({gain_a_range},">0")*{headline_a}))'
        )

        # Δ column = (panel B adj gain) − (panel A adj gain)
        asset_b, proc_b, cb_b, gain_b, tax_b = PANEL_B_COLS
        ws[f"{DELTA_COL}{c_row}"] = (
            f'=IF({proceeds}="","",{gain_b}{c_row}-{gain_a}{c_row})'
        )

        # Panel B (reset elected → cost base = MV)
        ws[f"{asset_b}{c_row}"] = f'=IF({code}="","",{name}&" ("&{code}&")")'
        ws[f"{proc_b}{c_row}"]  = f'=IF({proceeds}="","",{proceeds})'
        ws[f"{cb_b}{c_row}"]    = f'=IF({proceeds}="","",{mv})'
        ws[f"{gain_b}{c_row}"]  = _div296_adj_formula(proceeds, mv, held)
        ws[f"{tax_b}{c_row}"]   = (
            f'=IF({proceeds}="","",'
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
        value=(f"Showing first {DISPLAY_ROWS} assets — see the Analyser tab "
               f"for the full register (up to {ASSUMPTIONS.asset_register_rows} rows)."),
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

    manual_note = ws.cell(
        row=MANUAL_NOTE_ROW, column=1,
        value=("Note: Comparison panels always compute from the asset register. "
               "If 'Div 296 earnings source' is set to Manual on Inputs, the override "
               "applies to the Analyser headline only — it is ignored here."),
    )
    manual_note.font = Font(name="Arial", size=9, italic=True, color="666666")
    manual_note.alignment = Alignment(wrap_text=True, vertical="top")
    ws.merge_cells(f"A{MANUAL_NOTE_ROW}:{LAST_VISIBLE_COL}{MANUAL_NOTE_ROW}")
    ws.row_dimensions[MANUAL_NOTE_ROW].height = 30


def _build_chart(ws: Worksheet, headline_a_cell: str, headline_b_cell: str) -> None:
    chart = BarChart()
    chart.type = "col"
    chart.style = 11
    chart.title = "Div 296 tax — Scenario A vs Scenario B"
    chart.y_axis.title = "Div 296 tax ($)"
    chart.x_axis.title = None
    chart.legend = None

    data = Reference(ws, min_col=13, min_row=8, max_col=13, max_row=9)   # M8:M9
    cats = Reference(ws, min_col=12, min_row=8, max_col=12, max_row=9)   # L8:L9
    chart.add_data(data, titles_from_data=False)
    chart.set_categories(cats)
    chart.dataLabels = DataLabelList(
        showVal=True, showCatName=False, showSerName=False, showLegendKey=False,
    )
    chart.height = 8
    chart.width = 22

    scenario_a, scenario_b = _sample_scenario_headlines()
    series = chart.ser[0]
    series.val = NumDataSource(numRef=NumRef(
        f=f"'{SHEET}'!${CHART_VALUE_A_CELL[0]}${CHART_VALUE_A_CELL[1:]}:"
          f"${CHART_VALUE_B_CELL[0]}${CHART_VALUE_B_CELL[1:]}",
        numCache=NumData(
            formatCode='"$"#,##0',
            ptCount=2,
            pt=[NumVal(idx=0, v=scenario_a), NumVal(idx=1, v=scenario_b)],
        ),
    ))
    series.cat = AxDataSource(strRef=StrRef(
        f=f"'{SHEET}'!${CHART_LABEL_A_CELL[0]}${CHART_LABEL_A_CELL[1:]}:"
          f"${CHART_LABEL_B_CELL[0]}${CHART_LABEL_B_CELL[1:]}",
        strCache=StrData(
            ptCount=2,
            pt=[
                StrVal(idx=0, v="Scenario A — No reset"),
                StrVal(idx=1, v="Scenario B — Reset elected"),
            ],
        ),
    ))

    ws.add_chart(chart, f"A{CHART_ANCHOR_ROW}")


def build(wb: Workbook) -> Worksheet:
    ws = wb.create_sheet(SHEET)
    ws.sheet_view.showGridLines = False

    _build_header_block(ws)

    # Section band for the body
    _band(ws, BAND_BODY_ROW,
          "Side-by-side comparison (independent of the master reset toggle)")

    _build_context_strip(ws)

    gain_a_range, gain_b_range, headline_a, headline_b = _build_helpers(ws)

    _build_metric_cards(ws, headline_a, headline_b)
    _build_subtotals(ws, headline_a, headline_b, gain_a_range, gain_b_range)
    _build_per_asset_detail(ws, gain_a_range, gain_b_range, headline_a, headline_b)
    _build_footer_notes(ws)
    _build_chart(ws, headline_a, headline_b)

    # --- Print header watermark (large grey text on every printed page) ---
    ws.oddHeader.center.text = "ILLUSTRATIVE — NOT ADVICE"
    ws.oddHeader.center.size = 28
    ws.oddHeader.center.color = "CCCCCC"

    # --- Print setup: A4 landscape, narrow margins; let height spill if needed ---
    ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0           # may spill to 2 pages — print risk relaxed
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.print_options.horizontalCentered = True
    ws.page_margins.left = 0.25
    ws.page_margins.right = 0.25
    ws.page_margins.top = 0.5
    ws.page_margins.bottom = 0.4
    ws.print_area = f"A1:{LAST_VISIBLE_COL}{CHART_BOTTOM_ROW}"

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

    return ws
