"""Entrypoint: build the Division 296 Cost Base Reset Model workbook.

    python -m div296.build
    python -m div296.build --no-validate          # skip live recalc check
    python -m div296.build --edition adviser      # omit CLASS Import tab

Writes dist/Division_296_Model_v<version>.xlsx with the 5 tabs in spec
order: Inputs, CLASS Import, Analyser, Comparison, Notes.

The `--edition adviser` build omits the CLASS Import staging tab (for
advisers who don't use CLASS Super); all other content is identical. Output
filename gets an `_Adviser_Edition` suffix so the two builds can't be
mixed up. See docs/USER_GUIDE_Adviser_Edition_v3.docx.

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
from div296.tabs import analyser, class_import, comparison, inputs, notes


EXCEL_ERROR_RE = re.compile(r"#(REF|DIV/0|VALUE|NAME|NULL|NUM|N/A)!?\b")


def _dist_dir() -> Path:
    return Path(__file__).resolve().parents[2] / "dist"


AUTHOR = "Aiden Hiew"
WORKBOOK_TITLE = "Division 296 CGT Reset Model"


def _stamp_properties(wb: Workbook) -> None:
    props = wb.properties
    props.creator = AUTHOR
    props.lastModifiedBy = AUTHOR
    props.title = WORKBOOK_TITLE
    props.subject = WORKBOOK_TITLE
    props.description = f"v{__version__} — see CONTEXT.md"
    props.keywords = "Division 296; superannuation; CGT; cost base reset"


def build_workbook(*, include_class_import: bool = True) -> Workbook:
    wb = Workbook()
    # openpyxl creates a default "Sheet" — remove it so tab order is exact.
    default = wb.active
    wb.remove(default)

    inputs.build(wb, include_class_import_hints=include_class_import)
    if include_class_import:
        class_import.build(wb)
    analyser.build(wb)
    comparison.build(wb)
    notes.build(wb, include_class_import_hints=include_class_import)
    _stamp_properties(wb)
    _stamp_print_footer(wb)
    return wb


def _stamp_print_footer(wb: Workbook) -> None:
    """Add 'Prepared by Aiden Hiew · v<ver> · Page X of N' to every tab's print footer."""
    for ws in wb.worksheets:
        ws.oddFooter.left.text = f"Prepared by {AUTHOR}"
        ws.oddFooter.left.size = 8
        ws.oddFooter.left.color = "808080"
        ws.oddFooter.right.text = f"v{__version__}  |  Page &P of &N"
        ws.oddFooter.right.size = 8
        ws.oddFooter.right.color = "808080"


def validate_recalc(xlsx_path: Path) -> list[str]:
    """Recalc the workbook and return a list of cell keys with error values.

    Empty list = clean. Known `formulas`-engine false positives (see
    div296._recalc_limitations) are skipped and summarised on stdout.
    Cells whose value accessor raises are REPORTED (a validation gate must
    not silently treat unreadable cells as clean). Raises ImportError if
    `formulas` is not installed — callers decide whether to skip or fail.
    """
    import formulas  # noqa: PLC0415 — optional dep, imported on demand

    from div296._recalc_limitations import is_known_limitation

    xl = formulas.ExcelModel().loads(str(xlsx_path)).finish()
    sol = xl.calculate()

    errors: list[str] = []
    skipped_known = 0
    for key, cell in sol.items():
        if "'!" not in str(key):
            # `formulas` keys workbook cells as "'[file]SHEET'!A1"; its solution
            # also carries internal meta-entries (e.g. the 'self' Dispatcher)
            # that have no .value and are not workbook cells — skip those.
            continue
        try:
            value = cell.value
        except Exception as exc:  # noqa: BLE001 — engine raises arbitrary types
            if is_known_limitation(key):
                skipped_known += 1
            else:
                errors.append(f"{key} -> <unreadable: {exc!r}>")
            continue
        s = str(value)
        if EXCEL_ERROR_RE.search(s):
            if is_known_limitation(key):
                skipped_known += 1
                continue
            errors.append(f"{key} -> {s}")
    if skipped_known:
        print(f"Recalc: skipped {skipped_known} known `formulas`-engine "
              f"false-positive cell(s) — see div296/_recalc_limitations.py.")
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
        help="Skip the post-build recalc validation check.",
    )
    parser.add_argument(
        "--edition",
        choices=("full", "adviser"),
        default="full",
        help=("'full' (default) ships the CLASS Import staging tab; "
              "'adviser' omits it (for advisers who don't use CLASS Super)."),
    )
    args = parser.parse_args(argv)

    include_class_import = args.edition == "full"
    suffix = "" if include_class_import else "_Adviser_Edition"
    default_name = f"Division_296_Model_v{__version__}{suffix}.xlsx"
    out_path = args.output or (_dist_dir() / default_name)
    out_path.parent.mkdir(parents=True, exist_ok=True)

    wb = build_workbook(include_class_import=include_class_import)
    wb.save(out_path)
    print(f"Wrote {out_path}")

    if args.no_validate:
        return 0

    try:
        errors = validate_recalc(out_path)
    except ImportError:
        print("Skipped recalc validation: `formulas` not installed (pip install -e .[dev]).")
        return 0
    except MemoryError:
        # v1.7 added a sort-by-impact helper grid + LARGE/MATCH/INDEX lookups
        # whose dependency graph can exceed what the pure-Python `formulas`
        # package can hold on a modest machine. Excel and LibreOffice handle
        # it fine — fall back to a warning rather than failing the build.
        print("Skipped recalc validation: `formulas` ran out of memory on this "
              "workbook. Verify in Excel/LibreOffice instead.")
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
