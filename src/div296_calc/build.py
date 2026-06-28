"""Build the Ongoing Division 296 Calculator workbook.

    python -m div296_calc.build
    python -m div296_calc.build --no-validate

Writes dist/ongoing_calculator/Div_296_Ongoing_Calculator_v<version>.xlsx
(Calculator, Notes).
The recalc gate is STRICT: ANY Excel error cell fails the build (no
skip-list). The workbook is small enough that the pure-Python `formulas`
engine recalculates it without the OOM the year-one model hits.
"""

from __future__ import annotations

import argparse
import re
from pathlib import Path

from openpyxl import Workbook

from div296_calc import __version__
from div296_calc.tabs import calculator, notes

EXCEL_ERROR_RE = re.compile(r"#(REF|DIV/0|VALUE|NAME|NULL|NUM|N/A)!?\b")
AUTHOR = "Aiden Hiew"
TITLE = "Ongoing Division 296 Calculator"


def _dist_dir() -> Path:
    return Path(__file__).resolve().parents[2] / "dist" / "ongoing_calculator"


def build_workbook() -> Workbook:
    wb = Workbook()
    wb.remove(wb.active)
    calculator.build(wb)
    notes.build(wb)
    props = wb.properties
    props.creator = props.lastModifiedBy = AUTHOR
    props.title = props.subject = TITLE
    props.description = f"v{__version__} — ongoing Div 296 calculator"
    for ws in wb.worksheets:
        ws.oddFooter.left.text = f"Prepared by {AUTHOR}"
        ws.oddFooter.left.size = 8
        ws.oddFooter.right.text = f"v{__version__}  |  Page &P of &N"
        ws.oddFooter.right.size = 8
    return wb


def validate_recalc(xlsx_path: Path) -> list[str]:
    """Strict recalc — return every cell that resolves to an Excel error."""
    import formulas  # noqa: PLC0415

    xl = formulas.ExcelModel().loads(str(xlsx_path)).finish()
    sol = xl.calculate()
    errors: list[str] = []
    for key, cell in sol.items():
        if "'!" not in str(key):
            continue
        try:
            value = cell.value
        except Exception as exc:  # noqa: BLE001
            errors.append(f"{key} -> <unreadable: {exc!r}>")
            continue
        if EXCEL_ERROR_RE.search(str(value)):
            errors.append(f"{key} -> {value}")
    return errors


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Build the ongoing Div 296 calculator.")
    parser.add_argument("--output", "-o", type=Path, default=None)
    parser.add_argument("--no-validate", action="store_true")
    args = parser.parse_args(argv)

    out = args.output or (_dist_dir() / f"Div_296_Ongoing_Calculator_v{__version__}.xlsx")
    out.parent.mkdir(parents=True, exist_ok=True)
    build_workbook().save(out)
    print(f"Wrote {out}")

    if args.no_validate:
        return 0
    try:
        errors = validate_recalc(out)
    except ImportError:
        print("Skipped recalc: `formulas` not installed (pip install -e .[dev]).")
        return 0
    if errors:
        print(f"FAILED recalc — {len(errors)} error cell(s):")
        for e in errors[:20]:
            print(f"  {e}")
        return 1
    print("Recalc validation: OK (no Excel error cells).")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
