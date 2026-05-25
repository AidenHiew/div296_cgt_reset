"""Inputs tab — the only data-entry sheet.

Layout (v1.7 — Manual earnings dropped; control panel now has 3 toggles):

    Row 1       Title
    Row 2       Sample-data badge (warns staff before client share)
    Row 4       Band: 1. Control panel
    Rows 5-7    3 control levers (named ranges anchored here)
    Row 9       Band: 2. Members
    Row 10      Members header row
    Rows 11-14  4 member rows
    Row 16      Split-% sum check row
    Row 18      Band: 3. Asset register
    Row 19      Register header row
    Rows 20-69  50 register data rows
    Row 71      Band: 4. Advanced assumptions
    Rows 72-79  8 assumption rows (named ranges anchored here)

Downstream tabs read this layout via the constants below — do not
change row numbers without grepping for the constant names first.
"""

from __future__ import annotations

from openpyxl.comments import Comment
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.styles import Alignment, Font, PatternFill, Protection
from openpyxl.utils import get_column_letter
from openpyxl.workbook import Workbook
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.worksheet import Worksheet

from div296 import named_ranges as nr
from div296.assumptions import ASSUMPTIONS
from div296.styles import (
    BODY_FONT, CENTER, FMT_CURRENCY, FMT_INT, FMT_PERCENT, FMT_PERCENT_3,
    FMT_TEXT, INPUT_FILL, INPUT_FONT, SECTION_BAND_FILL,
    SECTION_BAND_FONT, THIN_BOX, TITLE_FONT,
)


# Cell-layout constants — single source of truth.
SHEET = "Inputs"

# --- Zone 1: Control panel (rows 5-7) ---
CONTROL_ROWS = {
    nr.RESET_ON:        ("Reset election",                        5, "B5", ["ON", "OFF"], "ON"),
    nr.TIER10_ON:       ("$10m / +25% tier",                      6, "B6", ["ON", "OFF"], "OFF"),
    nr.DISCOUNT_ON:     ("CGT discount",                          7, "B7", ["ON", "OFF"], "ON"),
}

# --- Zone 2: Members (rows 9-16) ---
MEMBERS_HEADER_ROW = 10
MEMBERS_FIRST_DATA_ROW = 11
MEMBER_HEADERS = [
    "Member",
    "TSB ($)",
    "Split % of fund earnings",
    "Proportion above $3m (auto)",
    "Proportion override (optional)",
]
MEMBERS_TOTAL_ROW = MEMBERS_FIRST_DATA_ROW + ASSUMPTIONS.member_count    # row 15 (v2.3)
SPLIT_CHECK_ROW = MEMBERS_FIRST_DATA_ROW + ASSUMPTIONS.member_count + 1   # row 16

# --- Zone 3: Asset register (rows 18-69) ---
REGISTER_HEADER_ROW = 19
REGISTER_FIRST_DATA_ROW = 20
REGISTER_LAST_DATA_ROW = REGISTER_FIRST_DATA_ROW + ASSUMPTIONS.asset_register_rows - 1   # row 71

REGISTER_HEADERS = [
    ("Asset code", FMT_TEXT),
    ("Asset name", FMT_TEXT),
    ("Quantity", FMT_INT),
    ("Original cost base", FMT_CURRENCY),
    ("Current market value (as at today)", FMT_CURRENCY),
    ("Market value at 30 Jun 2026", FMT_CURRENCY),
    ("Valuation source / date", FMT_TEXT),
    ("Projected sale proceeds", FMT_CURRENCY),
    ("Held > 12 months?", FMT_TEXT),
]

SAMPLE_REGISTER_ROWS = [
    ("P1", "Commercial property",   1,     800_000, 2_400_000, 2_400_000,
     "Independent val, 30/06/26",   2_600_000, "Yes"),
    ("S1", "Listed shares parcel",  5_000, 300_000, 520_000,   520_000,
     "ASX close 30/06/26",          600_000,   "Yes"),
    ("L1", "Loss-making holding",   2_000, 500_000, 100_000,   100_000,
     "Independent val, 30/06/26",   200_000,   "Yes"),
]

# --- Zone 4: Advanced assumptions (rows 73-81) ---
ADV_BAND_ROW = REGISTER_LAST_DATA_ROW + 2     # row 73
ADV_FIRST_ROW = ADV_BAND_ROW + 1              # row 74

ADV_ROWS = [
    (nr.RATE_TIER1,        "Div 296 additional rate — tier 1 ($3m–$10m)",   ASSUMPTIONS.rate_tier1,             FMT_PERCENT),
    (nr.RATE_TIER2,        "Div 296 additional rate — tier 2 (above $10m)", ASSUMPTIONS.rate_tier2,             FMT_PERCENT),
    (nr.THRESHOLD_1,       "Threshold 1",                                   ASSUMPTIONS.threshold_1,            FMT_CURRENCY),
    (nr.THRESHOLD_2,       "Threshold 2",                                   ASSUMPTIONS.threshold_2,            FMT_CURRENCY),
    (nr.DISCOUNT_RATE,     "CGT discount rate (1/3 = 33.333%)",             ASSUMPTIONS.discount_rate,          FMT_PERCENT_3),
    (nr.FUND_CGT_RATE,     "Fund CGT rate (accumulation phase)",            ASSUMPTIONS.fund_cgt_rate,          FMT_PERCENT),
    (nr.INDEXATION_INCR_1, "Indexation increment — threshold 1",            ASSUMPTIONS.indexation_increment_1, FMT_CURRENCY),
    (nr.INDEXATION_INCR_2, "Indexation increment — threshold 2",            ASSUMPTIONS.indexation_increment_2, FMT_CURRENCY),
]


def _band(ws: Worksheet, row: int, text: str, last_col_letter: str = "I") -> None:
    ws.cell(row=row, column=1, value=text).font = SECTION_BAND_FONT
    ws.merge_cells(f"A{row}:{last_col_letter}{row}")
    for col_idx in range(1, ord(last_col_letter) - ord("A") + 2):
        ws.cell(row=row, column=col_idx).fill = SECTION_BAND_FILL


def _input_cell(ws: Worksheet, coord: str, value=None, number_format: str | None = None) -> None:
    cell = ws[coord]
    if value is not None:
        cell.value = value
    cell.font = INPUT_FONT
    cell.fill = INPUT_FILL
    cell.border = THIN_BOX
    # Unlocked so it stays editable when the sheet is protected.
    cell.protection = Protection(locked=False)
    if number_format:
        cell.number_format = number_format


def _define_name(wb: Workbook, name: str, coord: str) -> None:
    """Create a workbook-scoped defined name pointing at Inputs!coord."""
    ref = f"'{SHEET}'!${coord[0]}${coord[1:]}" if coord[1:].isdigit() else f"'{SHEET}'!{coord}"
    wb.defined_names[name] = DefinedName(name=name, attr_text=ref)


def build(wb: Workbook) -> Worksheet:
    ws = wb.create_sheet(SHEET)

    # --- Row 1: Title ---
    ws["A1"] = "Division 296 Cost Base Reset Model — Inputs"
    ws["A1"].font = TITLE_FONT
    ws.merge_cells("A1:I1")

    # --- Row 2: Sample-data badge ---
    badge = ws.cell(
        row=2, column=1,
        value=("⚠  Sample data preloaded — overwrite every cell with your fund's "
               "actual figures before sharing with a client."),
    )
    badge.font = Font(name="Arial", size=10, bold=True, italic=True, color="8A6D00")
    badge.fill = PatternFill("solid", fgColor="FFF4CE")
    badge.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.merge_cells("A2:I2")
    ws.row_dimensions[2].height = 22

    # --- Row 3: TSB diagnostic (v2.2.0 — "does Div 296 even apply?") ---
    # Auto-checks the highest member TSB against the $3m threshold and shows
    # a green / amber line. Conditional formatting paints the row.
    last_member_row = MEMBERS_FIRST_DATA_ROW + ASSUMPTIONS.member_count - 1
    max_tsb_expr = f"MAX(B{MEMBERS_FIRST_DATA_ROW}:B{last_member_row})"
    diag_formula = (
        f'=IF({max_tsb_expr}=0,'
        f'"➤  Enter member TSBs below to see whether Div 296 applies to this fund.",'
        f'IF({max_tsb_expr}<=threshold_1,'
        f'"✓  All members are below the $3m TSB threshold — Div 296 does not currently apply to this fund.",'
        f'"⚠  At least one member has a TSB above $3m — Div 296 applies. See the Comparison tab for the modelled impact."))'
    )
    diag = ws.cell(row=3, column=1, value=diag_formula)
    diag.font = Font(name="Arial", size=10, bold=True, color="666666")
    diag.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.merge_cells("A3:I3")
    ws.row_dimensions[3].height = 22

    # Green when no member exceeds $3m.
    green_rule = FormulaRule(
        formula=[f"AND({max_tsb_expr}>0,{max_tsb_expr}<=threshold_1)"],
        fill=PatternFill("solid", fgColor="E1F5EE"),
    )
    # Amber when at least one member is above $3m.
    amber_rule = FormulaRule(
        formula=[f"{max_tsb_expr}>threshold_1"],
        fill=PatternFill("solid", fgColor="FBE9E9"),
    )
    ws.conditional_formatting.add("A3:I3", green_rule)
    ws.conditional_formatting.add("A3:I3", amber_rule)

    # --- Zone 1: Control panel ---
    _band(ws, 4, "1. Control panel (the demo levers)")
    for name, (label, row, coord, options, default) in CONTROL_ROWS.items():
        ws.cell(row=row, column=1, value=label).font = BODY_FONT
        _input_cell(ws, coord, value=default)
        _define_name(wb, name, coord)
        if options is not None:
            dv = DataValidation(type="list", formula1=f'"{",".join(options)}"', allow_blank=False)
            ws.add_data_validation(dv)
            dv.add(coord)
        # v2.2.0: scope comment on the Reset election toggle.
        if name == nr.RESET_ON:
            ws[coord].comment = Comment(
                ("This toggle affects the Analyser tab only.\n\n"
                 "The Comparison tab always shows both paths (default vs "
                 "election) regardless of what is selected here, so you can "
                 "compare side-by-side without flipping back and forth."),
                "v2.2 UX",
            )

    # --- Zone 2: Members (moved up — short, fund-level context) ---
    _band(ws, MEMBERS_HEADER_ROW - 1, "2. Members")
    for col_idx, header in enumerate(MEMBER_HEADERS, start=1):
        c = ws.cell(row=MEMBERS_HEADER_ROW, column=col_idx, value=header)
        c.font = SECTION_BAND_FONT
        c.fill = SECTION_BAND_FILL
        c.alignment = CENTER

    for i in range(ASSUMPTIONS.member_count):
        row = MEMBERS_FIRST_DATA_ROW + i
        ws.cell(row=row, column=1, value=f"Member {i+1}").font = BODY_FONT
        is_first = i == 0
        # Defaults: row 1 populated with §12 sample (single member, TSB $12m, split 100%);
        # rows 2–4 left blank.
        _input_cell(ws, f"B{row}", value=12_000_000 if is_first else None, number_format=FMT_CURRENCY)
        _input_cell(ws, f"C{row}", value=1.0 if is_first else None, number_format=FMT_PERCENT)
        # Auto proportion = MAX(0, (TSB - threshold_1) / TSB), guarded for blank TSB
        prop_formula = f"=IF(B{row}>0, MAX(0,(B{row}-threshold_1)/B{row}), 0)"
        prop_cell = ws.cell(row=row, column=4, value=prop_formula)
        prop_cell.number_format = FMT_PERCENT
        _input_cell(ws, f"E{row}", value=None, number_format=FMT_PERCENT)

    # v2.3: Member TSB total row — combined TSB across members for quick read.
    last_member_data_row = MEMBERS_FIRST_DATA_ROW + ASSUMPTIONS.member_count - 1
    total_label = ws.cell(row=MEMBERS_TOTAL_ROW, column=1, value="Total")
    total_label.font = Font(name="Arial", size=10, bold=True, color="1D3B34")
    total_label.alignment = Alignment(horizontal="left", indent=1)
    total_tsb = ws.cell(
        row=MEMBERS_TOTAL_ROW, column=2,
        value=f"=SUM(B{MEMBERS_FIRST_DATA_ROW}:B{last_member_data_row})",
    )
    total_tsb.font = Font(name="Arial", size=10, bold=True, color="1D3B34")
    total_tsb.number_format = FMT_CURRENCY
    total_tsb.border = THIN_BOX
    total_tsb.fill = PatternFill("solid", fgColor="EFF5F3")  # light teal band
    # Apply the same light teal band to cols C-E for visual continuity.
    for col_idx in range(3, 6):
        c = ws.cell(row=MEMBERS_TOTAL_ROW, column=col_idx)
        c.fill = PatternFill("solid", fgColor="EFF5F3")

    # Member-split sanity row + visual flag (spec §4 Zone 3).
    ws.cell(row=SPLIT_CHECK_ROW, column=1, value="Split % sum (must equal 100%)").font = BODY_FONT
    check_formula = (
        f"=SUM(C{MEMBERS_FIRST_DATA_ROW}:C{MEMBERS_FIRST_DATA_ROW + ASSUMPTIONS.member_count - 1})"
    )
    chk = ws.cell(row=SPLIT_CHECK_ROW, column=3, value=check_formula)
    chk.number_format = FMT_PERCENT
    # Red when not 100%, green when 100%.
    chk_range = f"C{SPLIT_CHECK_ROW}"
    ws.conditional_formatting.add(
        chk_range,
        CellIsRule(operator="notEqual", formula=["1"], fill=PatternFill("solid", fgColor="FBE9E9")),
    )
    ws.conditional_formatting.add(
        chk_range,
        CellIsRule(operator="equal", formula=["1"], fill=PatternFill("solid", fgColor="E1F5EE")),
    )

    # --- Zone 3: Asset register (50 rows; sample data preloaded) ---
    _band(
        ws, REGISTER_HEADER_ROW - 1,
        f"3. Asset register (50 rows; sample data pre-loaded in rows "
        f"{REGISTER_FIRST_DATA_ROW}–{REGISTER_FIRST_DATA_ROW + len(SAMPLE_REGISTER_ROWS) - 1})",
    )
    for col_idx, (header, _fmt) in enumerate(REGISTER_HEADERS, start=1):
        c = ws.cell(row=REGISTER_HEADER_ROW, column=col_idx, value=header)
        c.font = SECTION_BAND_FONT
        c.fill = SECTION_BAND_FILL
        c.alignment = CENTER

    # Held>12mo dropdown shared across all 50 rows.
    held_dv = DataValidation(type="list", formula1='"Yes,No"', allow_blank=True)
    ws.add_data_validation(held_dv)

    for offset in range(ASSUMPTIONS.asset_register_rows):
        row = REGISTER_FIRST_DATA_ROW + offset
        sample = SAMPLE_REGISTER_ROWS[offset] if offset < len(SAMPLE_REGISTER_ROWS) else None
        for col_idx, (_header, fmt) in enumerate(REGISTER_HEADERS, start=1):
            coord = ws.cell(row=row, column=col_idx).coordinate
            value = sample[col_idx - 1] if sample else None
            _input_cell(ws, coord, value=value, number_format=fmt)
        held_dv.add(f"I{row}")

    # v2.2.0: Loss-position highlight — tint any register row whose current
    # market value (col E) is below its original cost base (col D). Cosmetic
    # only; nothing in the calc engine reads from this.
    reg_first = REGISTER_FIRST_DATA_ROW
    reg_last = REGISTER_LAST_DATA_ROW
    loss_rule = FormulaRule(
        formula=[f'AND($D{reg_first}<>"",$E{reg_first}<>"",$E{reg_first}<$D{reg_first})'],
        fill=PatternFill("solid", fgColor="FBE9E9"),
    )
    ws.conditional_formatting.add(f"A{reg_first}:I{reg_last}", loss_rule)

    # --- Zone 4: Advanced assumptions (set-once constants — bottom) ---
    _band(ws, ADV_BAND_ROW, "4. Advanced assumptions (set once)")
    for i, (name, label, value, fmt) in enumerate(ADV_ROWS):
        row = ADV_FIRST_ROW + i
        ws.cell(row=row, column=1, value=label).font = BODY_FONT
        coord = f"B{row}"
        _input_cell(ws, coord, value=value, number_format=fmt)
        _define_name(wb, name, coord)

    # --- Column widths + freeze ---
    widths = [32, 26, 10, 18, 18, 26, 26, 20, 18]
    for col_idx, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(col_idx)].width = w
    # Freeze below the title + sample badge + TSB diagnostic so they remain sticky.
    ws.freeze_panes = "A4"

    # --- Sheet protection (tamper-evident, passwordless) ---
    # Input cells were individually unlocked via _input_cell();
    # everything else is locked by default.
    ws.protection.sheet = True
    ws.protection.formatColumns = False
    ws.protection.formatRows = False
    ws.protection.selectLockedCells = False
    ws.protection.selectUnlockedCells = False

    return ws
