"""Inputs tab — the only data-entry sheet.

Three zones (spec §4):
  1. Control panel (5 levers per Zone 1)
  2. Asset register (50 rows × 9 columns per Zone 2)
  3. Members & advanced assumptions (4 members + assumption cells per Zone 3)

Pre-loaded with the spec §12 sample data: 3 assets, single member TSB $12m.

Cell layout (used by named_ranges and downstream tabs — do not change without
also updating named_ranges.py and any formulas that reference these cells):
    Control panel:  B4..B8
    Register:       headers row 12, data rows 13..62 (50 rows)
    Members:        row 66..69; advanced assumptions row 72..78
"""

from __future__ import annotations

from openpyxl.styles import Protection
from openpyxl.utils import get_column_letter
from openpyxl.workbook import Workbook
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.worksheet import Worksheet

from div296 import named_ranges as nr
from div296.assumptions import ASSUMPTIONS
from div296.styles import (
    BODY_FONT, CENTER, FMT_CURRENCY, FMT_INT, FMT_PERCENT, FMT_PERCENT_3,
    FMT_TEXT, INPUT_FILL, INPUT_FONT, LEFT, SECTION_BAND_FILL,
    SECTION_BAND_FONT, THIN_BOX, TITLE_FONT,
)


# Cell-layout constants — single source of truth.
SHEET = "Inputs"

CONTROL_ROWS = {
    nr.RESET_ON:        ("Reset election", 4, "B4", ["ON", "OFF"], "ON"),
    nr.TIER10_ON:       ("$10m / +25% tier", 5, "B5", ["ON", "OFF"], "OFF"),
    nr.DISCOUNT_ON:     ("CGT discount", 6, "B6", ["ON", "OFF"], "ON"),
    nr.EARNINGS_SOURCE: ("Div 296 earnings source", 7, "B7", ["Auto", "Manual"], "Auto"),
    nr.MANUAL_EARNINGS: ("Manual earnings (used only if Manual)", 8, "B8", None, None),
}

REGISTER_HEADER_ROW = 12
REGISTER_FIRST_DATA_ROW = 13
REGISTER_LAST_DATA_ROW = REGISTER_FIRST_DATA_ROW + ASSUMPTIONS.asset_register_rows - 1

REGISTER_HEADERS = [
    ("Asset code", FMT_TEXT),
    ("Asset name", FMT_TEXT),
    ("Quantity", FMT_INT),
    ("Original cost base", FMT_CURRENCY),
    ("Total value", FMT_CURRENCY),
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

MEMBERS_HEADER_ROW = 65
MEMBERS_FIRST_DATA_ROW = 66
MEMBER_HEADERS = [
    "Member",
    "TSB ($)",
    "Split % of fund earnings",
    "Proportion above $3m (auto)",
    "Proportion override (optional)",
]

ADV_ROWS = [
    (nr.RATE_TIER1,        "Div 296 additional rate — tier 1 ($3m–$10m)", ASSUMPTIONS.rate_tier1,        FMT_PERCENT),
    (nr.RATE_TIER2,        "Div 296 additional rate — tier 2 (above $10m)", ASSUMPTIONS.rate_tier2,      FMT_PERCENT),
    (nr.THRESHOLD_1,       "Threshold 1",                                 ASSUMPTIONS.threshold_1,        FMT_CURRENCY),
    (nr.THRESHOLD_2,       "Threshold 2",                                 ASSUMPTIONS.threshold_2,        FMT_CURRENCY),
    (nr.DISCOUNT_RATE,     "CGT discount rate (1/3 = 33.333%)",           ASSUMPTIONS.discount_rate,      FMT_PERCENT_3),
    (nr.FUND_CGT_RATE,     "Fund CGT rate (accumulation phase)",          ASSUMPTIONS.fund_cgt_rate,      FMT_PERCENT),
    (nr.INDEXATION_INCR_1, "Indexation increment — threshold 1",          ASSUMPTIONS.indexation_increment_1, FMT_CURRENCY),
    (nr.INDEXATION_INCR_2, "Indexation increment — threshold 2",          ASSUMPTIONS.indexation_increment_2, FMT_CURRENCY),
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

    # --- Title ---
    ws["A1"] = "Division 296 Cost Base Reset Model — Inputs"
    ws["A1"].font = TITLE_FONT
    ws.merge_cells("A1:I1")

    # --- Zone 1: Control panel ---
    _band(ws, 3, "1. Control panel (the demo levers)")
    for name, (label, row, coord, options, default) in CONTROL_ROWS.items():
        ws.cell(row=row, column=1, value=label).font = BODY_FONT
        _input_cell(ws, coord, value=default,
                    number_format=FMT_CURRENCY if name == nr.MANUAL_EARNINGS else None)
        _define_name(wb, name, coord)
        if options is not None:
            dv = DataValidation(type="list", formula1=f'"{",".join(options)}"', allow_blank=False)
            ws.add_data_validation(dv)
            dv.add(coord)

    # --- Zone 2: Asset register ---
    _band(ws, 11, "2. Asset register (50 rows; sample data pre-loaded in rows 13–15)")
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

    # --- Zone 3: Members ---
    _band(ws, MEMBERS_HEADER_ROW - 1, "3. Members & advanced assumptions")
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

    # Member-split sanity row + visual flag (spec §4 Zone 3).
    split_check_row = MEMBERS_FIRST_DATA_ROW + ASSUMPTIONS.member_count + 1
    ws.cell(row=split_check_row, column=1, value="Split % sum (must equal 100%)").font = BODY_FONT
    check_formula = (
        f"=SUM(C{MEMBERS_FIRST_DATA_ROW}:C{MEMBERS_FIRST_DATA_ROW + ASSUMPTIONS.member_count - 1})"
    )
    chk = ws.cell(row=split_check_row, column=3, value=check_formula)
    chk.number_format = FMT_PERCENT
    # Red when not 100%, green when 100%.
    from openpyxl.formatting.rule import CellIsRule
    from openpyxl.styles import PatternFill as _PF
    chk_range = f"C{split_check_row}"
    ws.conditional_formatting.add(
        chk_range,
        CellIsRule(operator="notEqual", formula=["1"], fill=_PF("solid", fgColor="FBE9E9")),
    )
    ws.conditional_formatting.add(
        chk_range,
        CellIsRule(operator="equal", formula=["1"], fill=_PF("solid", fgColor="E1F5EE")),
    )

    # --- Advanced assumptions ---
    adv_first_row = split_check_row + 2
    _band(ws, adv_first_row - 1, "Advanced assumptions (set once)")
    for i, (name, label, value, fmt) in enumerate(ADV_ROWS):
        row = adv_first_row + i
        ws.cell(row=row, column=1, value=label).font = BODY_FONT
        coord = f"B{row}"
        _input_cell(ws, coord, value=value, number_format=fmt)
        _define_name(wb, name, coord)

    # --- Column widths + freeze ---
    widths = [10, 26, 10, 18, 18, 22, 26, 20, 18]
    for col_idx, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(col_idx)].width = w
    ws.freeze_panes = "A2"

    # --- Sheet protection (tamper-evident, passwordless) ---
    # Input cells were individually unlocked via _input_cell();
    # everything else is locked by default.
    ws.protection.sheet = True
    ws.protection.selectLockedCells = False
    ws.protection.selectUnlockedCells = False

    return ws
