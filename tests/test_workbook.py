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


# --- Analyser tab (chunk 3) ----------------------------------------------

def _analyser(tmp_path: Path):
    out = tmp_path / "out.xlsx"
    wb = build_workbook()
    wb.save(out)
    return load_workbook(out)["Analyser"]


class TestAnalyser:
    def test_mirror_levers_reference_inputs(self, tmp_path: Path):
        ws = _analyser(tmp_path)
        for row in range(4, 9):
            assert ws.cell(row=row, column=2).value == f"='Inputs'!B{row}"

    def test_fund_earnings_formula_uses_named_ranges(self, tmp_path: Path):
        ws = _analyser(tmp_path)
        f = ws["B11"].value
        assert 'earnings_source="Manual"' in f
        assert "manual_earnings" in f
        assert 'SUMIF(G20:G69,">0")' in f

    def test_headline_is_sum_of_member_tax(self, tmp_path: Path):
        ws = _analyser(tmp_path)
        assert ws["B16"].value == "=SUM(B12:B15)"

    def test_member_tax_formula_references_inputs_member_rows(self, tmp_path: Path):
        ws = _analyser(tmp_path)
        # Member 1 at row 12 should reference Inputs!B66 / C66 / D66 / E66.
        f = ws["B12"].value
        for ref in ("'Inputs'!B66", "'Inputs'!C66", "'Inputs'!D66", "'Inputs'!E66"):
            assert ref in f, f"member 1 formula missing {ref}"
        assert "rate_tier1" in f
        assert "rate_tier2" in f
        assert "threshold_1" in f
        assert "threshold_2" in f

    def test_first_data_row_columns(self, tmp_path: Path):
        ws = _analyser(tmp_path)
        # Row 20 = first asset (Inputs row 13).
        a, b, c, d, e, f, g, h, i = (ws.cell(row=20, column=col).value for col in range(1, 10))
        assert "'Inputs'!A13" in a and "'Inputs'!B13" in a
        assert "'Inputs'!H13" in b
        assert "'Inputs'!D13" in c
        assert 'reset_on="ON"' in d and "'Inputs'!F13" in d
        assert "discount_rate" in e and "'Inputs'!I13" in e
        assert "fund_cgt_rate" in f and "E20" in f         # uses col 5 of same row
        assert "discount_rate" in g and 'reset_on="ON"' in g
        assert "$B$16" in h and 'SUMIF(G20:G69,">0")' in h
        assert "J20" in i and "K20" in i                    # helper cols

    def test_helper_columns_hidden(self, tmp_path: Path):
        ws = _analyser(tmp_path)
        assert ws.column_dimensions["J"].hidden
        assert ws.column_dimensions["K"].hidden

    def test_totals_row_sums_correct_ranges(self, tmp_path: Path):
        ws = _analyser(tmp_path)
        assert ws["B70"].value == "=SUM(B20:B69)"
        assert ws["F70"].value == "=SUM(F20:F69)"
        assert ws["G70"].value == "=SUM(G20:G69)"
        assert ws["H70"].value == "=SUM(H20:H69)"

    def test_reconciliation_panel(self, tmp_path: Path):
        ws = _analyser(tmp_path)
        assert ws["A73"].value == "Ordinary CGT payable"
        assert ws["B73"].value == "=SUM(F20:F69)"
        assert ws["A74"].value == "Div 296 tax payable (headline)"
        assert ws["B74"].value == "=B16"
        assert ws["A75"].value == "Capital losses carried forward"
        # CF formula sums MAX(0, original − proceeds) across all 50 register rows.
        cf = ws["B75"].value
        assert cf.startswith("=")
        assert cf.count("MAX(0,'Inputs'!D") == 50  # one MAX per register row

    def test_trap_conditional_formatting_applied(self, tmp_path: Path):
        ws = _analyser(tmp_path)
        cf_rules = list(ws.conditional_formatting._cf_rules.items())
        assert cf_rules, "no conditional formatting rules on Analyser"
        # At least one rule covering the per-asset block.
        ranges = [str(r[0]) for r in cf_rules]
        assert any("A20" in r and "I69" in r for r in ranges)


# --- Comparison tab (chunk 4) --------------------------------------------

def _comparison(tmp_path: Path):
    out = tmp_path / "out.xlsx"
    wb = build_workbook()
    wb.save(out)
    return load_workbook(out)["Comparison"]


class TestComparison:
    def test_watermark_banner_and_print_header(self, tmp_path: Path):
        ws = _comparison(tmp_path)
        assert ws["A1"].value == "ILLUSTRATIVE — NOT ADVICE"
        # Print-header watermark for every printed page
        assert "ILLUSTRATIVE" in ws.oddHeader.center.text

    def test_panel_headers(self, tmp_path: Path):
        ws = _comparison(tmp_path)
        assert ws["A15"].value == "Scenario A — No reset"
        assert ws["E15"].value == "Scenario B — Reset elected"

    def test_panel_a_uses_original_cost_base(self, tmp_path: Path):
        """Scenario A panel must compute using original CB (Inputs!D), not MV (F)."""
        ws = _comparison(tmp_path)
        f = ws["B17"].value
        assert "'Inputs'!D13" in f
        assert "'Inputs'!F13" not in f

    def test_panel_b_uses_market_value(self, tmp_path: Path):
        """Scenario B panel must compute using MV (Inputs!F), not original CB (D)."""
        ws = _comparison(tmp_path)
        f = ws["F17"].value
        assert "'Inputs'!F13" in f
        assert "'Inputs'!D13" not in f

    def test_panels_independent_of_master_reset_toggle(self, tmp_path: Path):
        """Neither panel formula may reference the reset_on named range."""
        ws = _comparison(tmp_path)
        for row in range(17, 67):
            for col_letter in ("B", "C", "F", "G"):
                f = ws[f"{col_letter}{row}"].value
                if f and isinstance(f, str):
                    assert "reset_on" not in f, (
                        f"{col_letter}{row} references reset_on; panels must be independent"
                    )

    def test_helper_columns_hidden(self, tmp_path: Path):
        ws = _comparison(tmp_path)
        for col_letter in ("H", "I", "J", "K", "L", "M"):
            assert ws.column_dimensions[col_letter].hidden, f"col {col_letter} should be hidden"

    def test_subtotals_reference_helpers(self, tmp_path: Path):
        ws = _comparison(tmp_path)
        # Subtotal earnings: =H1 / =I1
        assert ws["C68"].value == "=H1"
        assert ws["G68"].value == "=I1"
        # Subtotal tax (headline): =H6 / =I6
        assert ws["C69"].value == "=H6"
        assert ws["G69"].value == "=I6"

    def test_neutral_net_effect_footer(self, tmp_path: Path):
        """Footer is a NEUTRAL net-effect calculation — no recommendation language anywhere."""
        ws = _comparison(tmp_path)
        # Net effect label
        label = ws["A72"].value
        assert "Net effect" in label
        # Forbidden words anywhere on the sheet (recommendation language).
        forbidden = ("saved", "saves", "created", "creates", "should",
                     "recommend", "we suggest")
        for row in ws.iter_rows(values_only=True):
            for v in row:
                if isinstance(v, str):
                    for word in forbidden:
                        assert word.lower() not in v.lower(), (
                            f"Comparison tab contains forbidden recommendation word "
                            f"{word!r} in cell value {v!r}"
                        )
        # Net effect formula references both headlines.
        net_formula = ws["C72"].value
        assert net_formula == "=H6-I6"

    def test_print_setup_landscape_a4(self, tmp_path: Path):
        ws = _comparison(tmp_path)
        assert ws.page_setup.orientation == ws.ORIENTATION_LANDSCAPE
        assert int(ws.page_setup.paperSize) == int(ws.PAPERSIZE_A4)
        assert int(ws.page_setup.fitToWidth) == 1
        assert ws.print_area is not None and "G" in ws.print_area
