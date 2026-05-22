"""Comparison tab — print-ready landscape A4 side-by-side per spec §8.

Both panels compute independently of the master reset toggle:
  - Scenario A — No reset: Div 296 cost base = original cost base
  - Scenario B — Reset elected: Div 296 cost base = MV 30 Jun 2026

Deviation from spec: the footer is a NEUTRAL "Net effect = A − B" calculation;
no "$X saved" / "$X created" framing, no recommendation language.

Watermark: "ILLUSTRATIVE — NOT ADVICE" is set as the print header (oddHeader)
and also appears as a banner row on screen.

Layout:
    Rows 1–11   Header block (editable: firm, logo, prepared for/by, date)
    Row 14      Section band: "Side-by-side comparison"
    Rows 15–16  Panel headers + column sub-headers
                    A..C  Scenario A panel    (D spacer)
                    E..G  Scenario B panel
    Rows 17–66  50 data rows per panel
    Rows 68–69  Subtotal Div 296 earnings + Div 296 tax (each scenario)
    Rows 71–72  Footer band + neutral net-effect line
    Cols H..M   Hidden: per-scenario fund earnings, 4 member tax cells, headline
"""

from __future__ import annotations

from openpyxl.styles import Alignment, Font, PatternFill, Protection
from openpyxl.utils import get_column_letter
from openpyxl.workbook import Workbook
from openpyxl.worksheet.worksheet import Worksheet

from div296.assumptions import ASSUMPTIONS
from div296.styles import (
    BODY_FONT, CENTER, FMT_CURRENCY, INPUT_FILL, INPUT_FONT,
    SECTION_BAND_FILL, SECTION_BAND_FONT, THIN_BOX, TITLE_FONT,
)
from div296.tabs.inputs import (
    MEMBERS_FIRST_DATA_ROW, REGISTER_FIRST_DATA_ROW,
)


SHEET = "Comparison"
INPUTS_SHEET = "'Inputs'"

# --- Layout constants ---
WATERMARK_ROW = 1
TITLE_ROW = 2
HEADER_FIRM_ROW = 4
HEADER_PREPARED_FOR_ROW = 5
HEADER_PREPARED_BY_ROW = 6
HEADER_DATE_ROW = 7
HEADER_DISCLAIMER_ROW = 8
HEADER_LOGO_ROW = 10

BAND_ROW = 14
PANEL_TITLE_ROW = 15
PANEL_HEADER_ROW = 16
DATA_FIRST_ROW = 17
DATA_LAST_ROW = DATA_FIRST_ROW + ASSUMPTIONS.asset_register_rows - 1   # row 66

SUBTOTAL_EARNINGS_ROW = DATA_LAST_ROW + 2     # row 68
SUBTOTAL_TAX_ROW = SUBTOTAL_EARNINGS_ROW + 1  # row 69

FOOTER_BAND_ROW = SUBTOTAL_TAX_ROW + 2        # row 71
NET_EFFECT_ROW = FOOTER_BAND_ROW + 1          # row 72
REMINDER_ROW = NET_EFFECT_ROW + 2             # row 74

# Hidden helper rows in cols H..M
HELPER_FUND_EARNINGS_ROW = 1
HELPER_MEMBER_TAX_FIRST_ROW = 2
HELPER_MEMBER_TAX_LAST_ROW = HELPER_MEMBER_TAX_FIRST_ROW + ASSUMPTIONS.member_count - 1
HELPER_HEADLINE_ROW = HELPER_MEMBER_TAX_LAST_ROW + 1   # row 6

# Hidden helper column letters
COL_A_EARNINGS = "H"      # Scenario A: fund earnings + members + headline
COL_B_EARNINGS = "I"      # Scenario B: same

# Panel column letters
PANEL_A_ASSET = "A"
PANEL_A_GAIN = "B"
PANEL_A_TAX = "C"
PANEL_B_ASSET = "E"
PANEL_B_GAIN = "F"
PANEL_B_TAX = "G"


def _band(ws: Worksheet, row: int, text: str, last_col_letter: str = "G") -> None:
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
    """Same shape as Analyser's per-member tax, but with the earnings cell parameterised."""
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


def build(wb: Workbook) -> Worksheet:
    ws = wb.create_sheet(SHEET)
    ws.sheet_view.showGridLines = False

    # --- On-screen ILLUSTRATIVE banner ---
    watermark_font = Font(name="Arial", size=14, bold=True, italic=True, color="999999")
    cell = ws.cell(row=WATERMARK_ROW, column=1, value="ILLUSTRATIVE — NOT ADVICE")
    cell.font = watermark_font
    cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.merge_cells(f"A{WATERMARK_ROW}:G{WATERMARK_ROW}")

    # --- Title ---
    title_cell = ws.cell(row=TITLE_ROW, column=1,
                         value="Division 296 Cost Base Reset Model — Comparison")
    title_cell.font = TITLE_FONT
    title_cell.alignment = Alignment(horizontal="center")
    ws.merge_cells(f"A{TITLE_ROW}:G{TITLE_ROW}")

    # --- Header block (editable cells) ---
    header_rows = [
        (HEADER_FIRM_ROW,         "Firm name:",      ""),
        (HEADER_PREPARED_FOR_ROW, "Prepared for:",   ""),
        (HEADER_PREPARED_BY_ROW,  "Prepared by:",    ""),
        (HEADER_DATE_ROW,         "Date:",           ""),
    ]
    for row, label, default in header_rows:
        ws.cell(row=row, column=1, value=label).font = BODY_FONT
        _input_cell(ws, f"B{row}", value=default)
        ws.merge_cells(f"B{row}:D{row}")

    # Disclaimer line (read-only)
    disc = ws.cell(row=HEADER_DISCLAIMER_ROW, column=1,
                   value="Illustrative only — not financial, tax or legal advice.")
    disc.font = Font(name="Arial", size=10, italic=True, color="666666")
    ws.merge_cells(f"A{HEADER_DISCLAIMER_ROW}:G{HEADER_DISCLAIMER_ROW}")

    # Logo placeholder
    logo = ws.cell(row=HEADER_LOGO_ROW, column=6, value="[ logo ]")
    logo.font = Font(name="Arial", size=10, italic=True, color="999999")
    logo.alignment = Alignment(horizontal="center", vertical="center")
    logo.fill = PatternFill("solid", fgColor="F5F5F5")
    ws.merge_cells(f"F{HEADER_LOGO_ROW}:G{HEADER_LOGO_ROW + 2}")

    # --- Body section band ---
    _band(ws, BAND_ROW, "Side-by-side comparison (independent of the master reset toggle)")

    # --- Panel headers ---
    panel_a_header = ws.cell(row=PANEL_TITLE_ROW, column=1, value="Scenario A — No reset")
    panel_b_header = ws.cell(row=PANEL_TITLE_ROW, column=5, value="Scenario B — Reset elected")
    for header in (panel_a_header, panel_b_header):
        header.font = SECTION_BAND_FONT
        header.fill = SECTION_BAND_FILL
        header.alignment = CENTER
    ws.merge_cells(f"A{PANEL_TITLE_ROW}:C{PANEL_TITLE_ROW}")
    ws.merge_cells(f"E{PANEL_TITLE_ROW}:G{PANEL_TITLE_ROW}")

    # Column sub-headers
    col_headers = [
        (PANEL_A_ASSET, "Asset"),
        (PANEL_A_GAIN, "Div 296 adjusted gain"),
        (PANEL_A_TAX, "Div 296 tax"),
        (PANEL_B_ASSET, "Asset"),
        (PANEL_B_GAIN, "Div 296 adjusted gain"),
        (PANEL_B_TAX, "Div 296 tax"),
    ]
    for col_letter, header in col_headers:
        c = ws[f"{col_letter}{PANEL_HEADER_ROW}"]
        c.value = header
        c.font = SECTION_BAND_FONT
        c.fill = SECTION_BAND_FILL
        c.alignment = CENTER

    # --- Hidden helper block (cols H, I; rows 1–6) ---
    # Scenario A: H1 = fund earnings (sum positive col B gains across data rows)
    gain_a_range = f"{PANEL_A_GAIN}{DATA_FIRST_ROW}:{PANEL_A_GAIN}{DATA_LAST_ROW}"
    gain_b_range = f"{PANEL_B_GAIN}{DATA_FIRST_ROW}:{PANEL_B_GAIN}{DATA_LAST_ROW}"
    ws[f"{COL_A_EARNINGS}{HELPER_FUND_EARNINGS_ROW}"] = f'=SUMIF({gain_a_range},">0")'
    ws[f"{COL_B_EARNINGS}{HELPER_FUND_EARNINGS_ROW}"] = f'=SUMIF({gain_b_range},">0")'

    # Member-tax rows
    for i in range(ASSUMPTIONS.member_count):
        helper_row = HELPER_MEMBER_TAX_FIRST_ROW + i
        inputs_row = MEMBERS_FIRST_DATA_ROW + i
        a_earnings = f"$H${HELPER_FUND_EARNINGS_ROW}"
        b_earnings = f"$I${HELPER_FUND_EARNINGS_ROW}"
        ws[f"{COL_A_EARNINGS}{helper_row}"] = _member_tax_formula(inputs_row, a_earnings)
        ws[f"{COL_B_EARNINGS}{helper_row}"] = _member_tax_formula(inputs_row, b_earnings)

    # Headlines
    ws[f"{COL_A_EARNINGS}{HELPER_HEADLINE_ROW}"] = (
        f"=SUM({COL_A_EARNINGS}{HELPER_MEMBER_TAX_FIRST_ROW}:"
        f"{COL_A_EARNINGS}{HELPER_MEMBER_TAX_LAST_ROW})"
    )
    ws[f"{COL_B_EARNINGS}{HELPER_HEADLINE_ROW}"] = (
        f"=SUM({COL_B_EARNINGS}{HELPER_MEMBER_TAX_FIRST_ROW}:"
        f"{COL_B_EARNINGS}{HELPER_MEMBER_TAX_LAST_ROW})"
    )

    # Hide helper columns
    for col_letter in (COL_A_EARNINGS, COL_B_EARNINGS, "J", "K", "L", "M"):
        ws.column_dimensions[col_letter].hidden = True

    # --- Per-asset data rows (50 per panel) ---
    headline_a = f"$H${HELPER_HEADLINE_ROW}"
    headline_b = f"$I${HELPER_HEADLINE_ROW}"

    for offset in range(ASSUMPTIONS.asset_register_rows):
        c_row = DATA_FIRST_ROW + offset
        i_row = REGISTER_FIRST_DATA_ROW + offset
        code = f"{INPUTS_SHEET}!A{i_row}"
        name = f"{INPUTS_SHEET}!B{i_row}"
        orig = f"{INPUTS_SHEET}!D{i_row}"
        mv = f"{INPUTS_SHEET}!F{i_row}"
        proceeds = f"{INPUTS_SHEET}!H{i_row}"
        held = f"{INPUTS_SHEET}!I{i_row}"

        # Scenario A panel (no reset → original cost base)
        ws[f"{PANEL_A_ASSET}{c_row}"] = f'=IF({code}="","",{name}&" ("&{code}&")")'
        ws[f"{PANEL_A_GAIN}{c_row}"] = _div296_adj_formula(proceeds, orig, held)
        ws[f"{PANEL_A_TAX}{c_row}"] = (
            f'=IF({proceeds}="","",'
            f'IF(SUMIF({gain_a_range},">0")=0,0,'
            f'MAX(0,{PANEL_A_GAIN}{c_row})/SUMIF({gain_a_range},">0")*{headline_a}))'
        )

        # Scenario B panel (reset → MV 30 Jun 2026)
        ws[f"{PANEL_B_ASSET}{c_row}"] = f'=IF({code}="","",{name}&" ("&{code}&")")'
        ws[f"{PANEL_B_GAIN}{c_row}"] = _div296_adj_formula(proceeds, mv, held)
        ws[f"{PANEL_B_TAX}{c_row}"] = (
            f'=IF({proceeds}="","",'
            f'IF(SUMIF({gain_b_range},">0")=0,0,'
            f'MAX(0,{PANEL_B_GAIN}{c_row})/SUMIF({gain_b_range},">0")*{headline_b}))'
        )

        for col_letter in (PANEL_A_GAIN, PANEL_A_TAX, PANEL_B_GAIN, PANEL_B_TAX):
            ws[f"{col_letter}{c_row}"].number_format = FMT_CURRENCY

    # --- Subtotals ---
    # Earnings subtotal sits under the GAIN column header; tax subtotal under TAX.
    ws.cell(row=SUBTOTAL_EARNINGS_ROW, column=1,
            value="Subtotal — Div 296 earnings").font = SECTION_BAND_FONT
    ws.cell(row=SUBTOTAL_EARNINGS_ROW, column=1).fill = SECTION_BAND_FILL
    ws[f"{PANEL_A_GAIN}{SUBTOTAL_EARNINGS_ROW}"] = f"={COL_A_EARNINGS}{HELPER_FUND_EARNINGS_ROW}"
    ws[f"{PANEL_B_GAIN}{SUBTOTAL_EARNINGS_ROW}"] = f"={COL_B_EARNINGS}{HELPER_FUND_EARNINGS_ROW}"

    ws.cell(row=SUBTOTAL_TAX_ROW, column=1,
            value="Subtotal — Div 296 tax (headline)").font = SECTION_BAND_FONT
    ws.cell(row=SUBTOTAL_TAX_ROW, column=1).fill = SECTION_BAND_FILL
    ws[f"{PANEL_A_TAX}{SUBTOTAL_TAX_ROW}"] = f"={COL_A_EARNINGS}{HELPER_HEADLINE_ROW}"
    ws[f"{PANEL_B_TAX}{SUBTOTAL_TAX_ROW}"] = f"={COL_B_EARNINGS}{HELPER_HEADLINE_ROW}"

    for col_letter in (PANEL_A_GAIN, PANEL_B_GAIN):
        cell = ws[f"{col_letter}{SUBTOTAL_EARNINGS_ROW}"]
        cell.number_format = FMT_CURRENCY
        cell.font = SECTION_BAND_FONT
        cell.fill = SECTION_BAND_FILL
    for col_letter in (PANEL_A_TAX, PANEL_B_TAX):
        cell = ws[f"{col_letter}{SUBTOTAL_TAX_ROW}"]
        cell.number_format = FMT_CURRENCY
        cell.font = SECTION_BAND_FONT
        cell.fill = SECTION_BAND_FILL

    # --- Neutral footer (NO recommendation language) ---
    _band(ws, FOOTER_BAND_ROW, "Footer")
    ws.cell(row=NET_EFFECT_ROW, column=1,
            value="Net effect of electing the reset (= Scenario A − Scenario B)").font = BODY_FONT
    ws[f"{PANEL_A_TAX}{NET_EFFECT_ROW}"] = (
        f"={COL_A_EARNINGS}{HELPER_HEADLINE_ROW}-{COL_B_EARNINGS}{HELPER_HEADLINE_ROW}"
    )
    ws[f"{PANEL_A_TAX}{NET_EFFECT_ROW}"].number_format = FMT_CURRENCY
    ws[f"{PANEL_A_TAX}{NET_EFFECT_ROW}"].font = Font(name="Arial", size=11, bold=True)

    reminder = ws.cell(
        row=REMINDER_ROW, column=1,
        value=("Note: loss-position assets may contribute Div 296 tax under Scenario B "
               "that they do not under Scenario A — see the Analyser tab for per-asset detail."),
    )
    reminder.font = Font(name="Arial", size=9, italic=True, color="666666")
    ws.merge_cells(f"A{REMINDER_ROW}:G{REMINDER_ROW}")

    manual_note = ws.cell(
        row=REMINDER_ROW + 1, column=1,
        value=("Note: Comparison panels always compute from the asset register. "
               "If 'Div 296 earnings source' is set to Manual on Inputs, the override "
               "applies to the Analyser headline only — it is ignored here."),
    )
    manual_note.font = Font(name="Arial", size=9, italic=True, color="666666")
    ws.merge_cells(f"A{REMINDER_ROW + 1}:G{REMINDER_ROW + 1}")

    # --- Print header watermark (large gray text on every printed page) ---
    ws.oddHeader.center.text = "ILLUSTRATIVE — NOT ADVICE"
    ws.oddHeader.center.size = 28
    ws.oddHeader.center.color = "CCCCCC"

    # --- Print setup: A4 landscape, fit-to-1-page-wide ---
    ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.print_options.horizontalCentered = True
    ws.page_margins.left = 0.5
    ws.page_margins.right = 0.5
    ws.page_margins.top = 0.75
    ws.page_margins.bottom = 0.5
    ws.print_area = f"A1:G{REMINDER_ROW}"

    # --- Column widths ---
    widths = {
        "A": 28, "B": 18, "C": 16,
        "D": 3,
        "E": 28, "F": 18, "G": 16,
    }
    for col_letter, w in widths.items():
        ws.column_dimensions[col_letter].width = w

    # Editable header-block cells unlocked so they remain user-editable
    # under sheet protection.
    for row in (HEADER_FIRM_ROW, HEADER_PREPARED_FOR_ROW,
                HEADER_PREPARED_BY_ROW, HEADER_DATE_ROW):
        ws[f"B{row}"].protection = Protection(locked=False)

    ws.protection.sheet = True

    return ws
