"""Analyser tab — full 9-column audit trail.

Per spec §5: mirror-lever strip at top, one row per asset (50 rows),
columns 1–9 (Asset, Proceeds, Original cost base, Div 296 cost base,
Ordinary taxable gain, Ordinary CGT, Div 296 adjusted gain, Div 296 tax,
Reset impact), totals row, reconciliation panel.

Per-asset Div 296 tax (col 8) = pro-rata of the headline member-attributed
Div 296 total: (this asset's positive adjusted gain ÷ sum of positive
adjusted gains) × headline. Per-asset sum always ties to the headline.

Populated in v1.0. v0.1 writes only the title banner.
"""

from openpyxl.workbook import Workbook
from openpyxl.worksheet.worksheet import Worksheet

from div296.styles import TITLE_FONT


def build(wb: Workbook) -> Worksheet:
    ws = wb.create_sheet("Analyser")
    ws["A1"] = "Division 296 Cost Base Reset Model — Analyser"
    ws["A1"].font = TITLE_FONT
    ws.column_dimensions["A"].width = 32
    ws.freeze_panes = "A2"
    return ws
