"""Integration tests: build the workbook, optionally recalc, read it back.

Placeholder in v0.1: confirms the build runs and produces a 4-tab file
in the expected order. v1.0 will add LibreOffice headless recalc and
assert the reconciliation totals against §12.
"""

from pathlib import Path

from openpyxl import load_workbook

from div296.build import build_workbook


EXPECTED_TABS = ["Inputs", "Analyser", "Comparison", "Notes"]


def test_workbook_has_four_tabs_in_spec_order(tmp_path: Path):
    out = tmp_path / "out.xlsx"
    wb = build_workbook()
    wb.save(out)

    reopened = load_workbook(out)
    assert reopened.sheetnames == EXPECTED_TABS
