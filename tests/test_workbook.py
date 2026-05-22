"""Integration tests: build the workbook, optionally recalc, read it back.

v0.1 confirmed the build runs and produces 4 tabs in spec order.
v1.0 adds verification that the Inputs tab has the right structure
and that every named range resolves to a real cell. LibreOffice headless
recalc and full reconciliation assertions land with chunk 6.
"""

from pathlib import Path

from openpyxl import load_workbook

from div296 import named_ranges as nr
from div296.build import build_workbook


EXPECTED_TABS = ["Inputs", "Analyser", "Comparison", "Notes"]


def test_workbook_has_four_tabs_in_spec_order(tmp_path: Path):
    out = tmp_path / "out.xlsx"
    wb = build_workbook()
    wb.save(out)
    reopened = load_workbook(out)
    assert reopened.sheetnames == EXPECTED_TABS


def test_every_named_range_resolves(tmp_path: Path):
    out = tmp_path / "out.xlsx"
    wb = build_workbook()
    wb.save(out)
    reopened = load_workbook(out)

    missing = [name for name in nr.ALL_NAMES if name not in reopened.defined_names]
    assert not missing, f"Named ranges not found in workbook: {missing}"

    for name in nr.ALL_NAMES:
        dn = reopened.defined_names[name]
        destinations = list(dn.destinations)
        assert destinations, f"Named range {name!r} has no destinations"
        for sheet_name, ref in destinations:
            assert sheet_name in reopened.sheetnames, f"{name} points at missing sheet {sheet_name}"
            ws = reopened[sheet_name]
            # If the reference resolves at all, openpyxl returns a cell.
            cell = ws[ref.replace("$", "")]
            assert cell is not None


def test_inputs_sample_data_preloaded(tmp_path: Path):
    """The three §12 sample assets are pre-loaded in rows 13–15."""
    out = tmp_path / "out.xlsx"
    wb = build_workbook()
    wb.save(out)
    ws = load_workbook(out)["Inputs"]

    # Row 13: Commercial property
    assert ws["A13"].value == "P1"
    assert ws["B13"].value == "Commercial property"
    assert ws["D13"].value == 800_000
    assert ws["F13"].value == 2_400_000
    assert ws["H13"].value == 2_600_000
    assert ws["I13"].value == "Yes"

    # Row 14: Listed shares parcel
    assert ws["A14"].value == "S1"
    assert ws["D14"].value == 300_000
    assert ws["F14"].value == 520_000

    # Row 15: Loss-making holding
    assert ws["A15"].value == "L1"
    assert ws["D15"].value == 500_000
    assert ws["F15"].value == 100_000

    # Sole member: TSB $12m, split 100%
    assert ws["B66"].value == 12_000_000
    assert ws["C66"].value == 1.0


def test_control_panel_defaults(tmp_path: Path):
    out = tmp_path / "out.xlsx"
    wb = build_workbook()
    wb.save(out)
    ws = load_workbook(out)["Inputs"]

    assert ws["B4"].value == "ON"     # Reset election
    assert ws["B5"].value == "OFF"    # $10m / +25% tier
    assert ws["B6"].value == "ON"     # CGT discount
    assert ws["B7"].value == "Auto"   # earnings source
    assert ws["B8"].value is None     # manual earnings
