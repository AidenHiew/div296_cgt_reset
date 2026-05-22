"""Entrypoint: build the Division 296 Cost Base Reset Model workbook.

    python -m div296.build
    python -m div296.build --no-validate   # skip live recalc check

Writes dist/Division_296_Model_v<version>.xlsx with the 4 tabs in spec
order: Inputs, Analyser, Comparison, Notes.

By default, after writing the file the build runs a pure-Python recalc
via the `formulas` package and fails with a non-zero exit if any cell
resolves to an Excel error sentinel (#REF!, #DIV/0!, #VALUE!, #NAME?,
#NULL!, #NUM!, #N/A). This is the regression gate that catches a broken
formula at build time rather than at the client's screen.
"""

from __future__ import annotations

import argparse
import re
from pathlib import Path

from openpyxl import Workbook

from div296 import __version__
from div296.tabs import analyser, comparison, inputs, notes


EXCEL_ERROR_RE = re.compile(r"#(REF|DIV/0|VALUE|NAME|NULL|NUM|N/A)!?\b")


def _dist_dir() -> Path:
    return Path(__file__).resolve().parents[2] / "dist"


def build_workbook() -> Workbook:
    wb = Workbook()
    # openpyxl creates a default "Sheet" — remove it so tab order is exact.
    default = wb.active
    wb.remove(default)

    inputs.build(wb)
    analyser.build(wb)
    comparison.build(wb)
    notes.build(wb)
    return wb


def validate_recalc(xlsx_path: Path) -> list[str]:
    """Recalc the workbook and return a list of cell keys with error values.

    Empty list = clean. Raises ImportError if `formulas` is not installed —
    callers decide whether to skip or fail in that case.
    """
    import formulas  # noqa: PLC0415 — optional dep, imported on demand

    xl = formulas.ExcelModel().loads(str(xlsx_path)).finish()
    sol = xl.calculate()

    errors: list[str] = []
    for key, cell in sol.items():
        try:
            value = cell.value
        except Exception:
            continue
        # value may be ndarray, scalar, str, or formulas error object
        s = str(value)
        if EXCEL_ERROR_RE.search(s):
            errors.append(f"{key} -> {s}")
    return errors


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Build the Div 296 workbook.")
    parser.add_argument(
        "--output", "-o",
        type=Path,
        default=None,
        help="Output path (default: dist/Division_296_Model_v<version>.xlsx)",
    )
    parser.add_argument(
        "--no-validate",
        action="store_true",
        help="Skip the post-build recalc validation (faster; not recommended).",
    )
    args = parser.parse_args(argv)

    out_path = args.output or (_dist_dir() / f"Division_296_Model_v{__version__}.xlsx")
    out_path.parent.mkdir(parents=True, exist_ok=True)

    wb = build_workbook()
    wb.save(out_path)
    print(f"Wrote {out_path}")

    if args.no_validate:
        return 0

    try:
        errors = validate_recalc(out_path)
    except ImportError:
        print("Skipped recalc validation: `formulas` not installed (pip install -e .[dev]).")
        return 0

    if errors:
        print(f"FAILED recalc validation — {len(errors)} cell(s) with Excel error values:")
        for err in errors[:20]:
            print(f"  {err}")
        if len(errors) > 20:
            print(f"  ... and {len(errors) - 20} more")
        return 1

    print("Recalc validation: OK (no Excel error cells).")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
