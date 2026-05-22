"""Entrypoint: build the Division 296 Cost Base Reset Model workbook.

    python -m div296.build

Writes dist/Division_296_Model_v<version>.xlsx with the 4 tabs in spec
order: Inputs, Analyser, Comparison, Notes.
"""

from __future__ import annotations

import argparse
from pathlib import Path

from openpyxl import Workbook

from div296 import __version__
from div296.tabs import analyser, comparison, inputs, notes


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


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Build the Div 296 workbook.")
    parser.add_argument(
        "--output", "-o",
        type=Path,
        default=None,
        help="Output path (default: dist/Division_296_Model_v<version>.xlsx)",
    )
    args = parser.parse_args(argv)

    out_path = args.output or (_dist_dir() / f"Division_296_Model_v{__version__}.xlsx")
    out_path.parent.mkdir(parents=True, exist_ok=True)

    wb = build_workbook()
    wb.save(out_path)
    print(f"Wrote {out_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
