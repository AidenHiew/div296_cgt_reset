"""Comparison tab — print-ready landscape A4 side-by-side.

Per spec §8: header block (firm, logo, prepared for/by, date, disclaimer),
two panels (Scenario A no-reset / Scenario B reset elected), neutral
net-effect footer.

Deviation from spec: footer is a NEUTRAL "Net effect = Scenario A − Scenario B = $X"
calculation only. No "$X saved" / "$X created" framing, no recommendation
language. Adds a diagonal "ILLUSTRATIVE — NOT ADVICE" watermark across
the print area.

Populated in v1.0. v0.1 writes only the title banner.
"""

from openpyxl.workbook import Workbook
from openpyxl.worksheet.worksheet import Worksheet

from div296.styles import TITLE_FONT


def build(wb: Workbook) -> Worksheet:
    ws = wb.create_sheet("Comparison")
    ws["A1"] = "Division 296 Cost Base Reset Model — Comparison"
    ws["A1"].font = TITLE_FONT
    ws.column_dimensions["A"].width = 32
    ws.sheet_view.showGridLines = False
    return ws
