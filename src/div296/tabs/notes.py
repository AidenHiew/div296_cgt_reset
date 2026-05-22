"""Notes tab — terminology, caveats, valuation log, provenance.

Per spec §10, with the additional factual disclosures listed in
README.md (loss-offset divergence, pension-phase exclusion, reset-OFF
realised-only, wash sale / Part IVA, transaction costs, alternative
levers). Hidden provenance cell carries build_version / build_date /
git short-SHA.

Populated in v1.0. v0.1 writes only the title banner + a provenance stub.
"""

import datetime as _dt
import subprocess as _sp

from openpyxl.workbook import Workbook
from openpyxl.worksheet.worksheet import Worksheet

from div296 import __version__ as _ver
from div296.styles import TITLE_FONT


def _git_short_sha() -> str:
    try:
        out = _sp.run(
            ["git", "rev-parse", "--short", "HEAD"],
            capture_output=True, text=True, check=False, timeout=2,
        )
        return out.stdout.strip() or "uncommitted"
    except Exception:
        return "unknown"


def build(wb: Workbook) -> Worksheet:
    ws = wb.create_sheet("Notes")
    ws["A1"] = "Division 296 Cost Base Reset Model — Notes"
    ws["A1"].font = TITLE_FONT
    ws.column_dimensions["A"].width = 32

    ws["A50"] = "build_version"
    ws["B50"] = _ver
    ws["A51"] = "build_date"
    ws["B51"] = _dt.date.today().isoformat()
    ws["A52"] = "git_short_sha"
    ws["B52"] = _git_short_sha()
    for row in (50, 51, 52):
        ws.row_dimensions[row].hidden = True
    return ws
