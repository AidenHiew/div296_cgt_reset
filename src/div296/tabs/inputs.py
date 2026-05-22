"""Inputs tab — the only data-entry sheet.

Three zones:
  1. Control panel (5 levers per spec §4 Zone 1)
  2. Asset register (50 rows, 9 columns per spec §4 Zone 2)
  3. Members & advanced assumptions (4 members + assumption cells per §4 Zone 3)

Populated in v1.0. v0.1 writes only the title banner.
"""

from openpyxl.workbook import Workbook
from openpyxl.worksheet.worksheet import Worksheet

from div296.styles import TITLE_FONT


def build(wb: Workbook) -> Worksheet:
    ws = wb.create_sheet("Inputs")
    ws["A1"] = "Division 296 Cost Base Reset Model — Inputs"
    ws["A1"].font = TITLE_FONT
    ws.column_dimensions["A"].width = 32
    ws.freeze_panes = "A2"
    return ws
