"""Integration test: build → recalc with `formulas` → assert §12 totals.

This is the strongest verification gate we have: the live Excel formulas
in the built `.xlsx` are recalculated by the pure-Python `formulas`
package and the resulting numbers are asserted against the spec §12
acceptance table.

If this test passes, the Excel formulas are demonstrated to reproduce
the spec numbers when opened in Excel/LibreOffice.

Skips cleanly (with an explanatory message) if `formulas` isn't
installed — the rest of the suite still runs.
"""

from __future__ import annotations


import numpy as np
import pytest

from div296.build import build_workbook


formulas = pytest.importorskip(
    "formulas",
    reason="`formulas` package required for live recalc test — pip install -e .[dev]",
)


# --- helpers --------------------------------------------------------------

def _unwrap(value):
    """`formulas` returns 2-D ndarray cells; unwrap to a scalar."""
    if isinstance(value, np.ndarray):
        return float(value.flat[0])
    return float(value)


@pytest.fixture(scope="module")
def recalc_solution(tmp_path_factory) -> dict:
    out = tmp_path_factory.mktemp("recalc") / "model.xlsx"
    wb = build_workbook()
    wb.save(out)

    xl = formulas.ExcelModel().loads(str(out)).finish()
    sol = xl.calculate()

    file_token = f"[{out.name}]"

    def at(sheet: str, cell: str):
        key = f"'{file_token}{sheet.upper()}'!{cell}"
        return _unwrap(sol[key].value)

    return {
        # Inputs (v1.5: member 1 moved to row 13)
        "member1_proportion": at("Inputs", "D13"),
        # Analyser headline + reconciliation
        "headline_div296_tax": at("Analyser", "B16"),
        "ord_cgt_payable": at("Analyser", "B73"),
        "div296_tax_payable": at("Analyser", "B74"),
        "cf_losses": at("Analyser", "B75"),
        # Analyser totals row (v1.5: Ord CGT total moved from F70 to E70)
        "totals_proceeds": at("Analyser", "B70"),
        "totals_ord_cgt": at("Analyser", "E70"),
        "totals_div296_adj_gain": at("Analyser", "G70"),
        "totals_div296_tax": at("Analyser", "H70"),
        # Comparison subtotals (v1.5: moved to top — rows 22-25)
        "scenario_a_earnings": at("Comparison", "B22"),
        "scenario_b_earnings": at("Comparison", "C22"),
        "scenario_a_ord_cgt": at("Comparison", "B23"),
        "scenario_a_tax": at("Comparison", "B24"),
        "scenario_b_tax": at("Comparison", "C24"),
        "scenario_a_burden": at("Comparison", "B25"),
        "scenario_b_burden": at("Comparison", "C25"),
        # Net effect = scenario A − scenario B (Div 296 tax delta), in card I18
        "net_effect": at("Comparison", "I18"),
    }


# --- §12 acceptance assertions ------------------------------------------

def _r(x: float) -> int:
    return int(round(x))


def test_member_proportion_75pct(recalc_solution):
    assert round(recalc_solution["member1_proportion"], 4) == 0.75


def test_headline_div296_tax_28500(recalc_solution):
    assert _r(recalc_solution["headline_div296_tax"]) == 28_500


def test_ordinary_cgt_payable_210000(recalc_solution):
    assert _r(recalc_solution["ord_cgt_payable"]) == 210_000


def test_div296_tax_payable_28500(recalc_solution):
    assert _r(recalc_solution["div296_tax_payable"]) == 28_500


def test_capital_losses_carried_forward_300000(recalc_solution):
    assert _r(recalc_solution["cf_losses"]) == 300_000


def test_analyser_totals(recalc_solution):
    s = recalc_solution
    assert _r(s["totals_proceeds"]) == 3_400_000          # 2.6m + 0.6m + 0.2m
    assert _r(s["totals_ord_cgt"]) == 210_000
    assert _r(s["totals_div296_adj_gain"]) == 253_333
    assert _r(s["totals_div296_tax"]) == 28_500           # pro-rata sum ties to headline


def test_comparison_scenario_a_no_reset_157500(recalc_solution):
    """Spec §12: reset OFF, fund earnings 1.4m × 75% × 15% = 157,500."""
    assert _r(recalc_solution["scenario_a_tax"]) == 157_500


def test_comparison_scenario_b_reset_28500(recalc_solution):
    assert _r(recalc_solution["scenario_b_tax"]) == 28_500


def test_comparison_net_effect_129000(recalc_solution):
    """Spec §12: 157,500 − 28,500 = 129,000."""
    assert _r(recalc_solution["net_effect"]) == 129_000


def test_comparison_earnings_subtotals(recalc_solution):
    """v1.5: Div 296 earnings on row 22, equal to sum of positive adj gains."""
    s = recalc_solution
    assert _r(s["scenario_a_earnings"]) == 1_400_000   # 1.2m + 0.2m + 0 (loss floored)
    assert _r(s["scenario_b_earnings"]) == 253_333     # 133,333 + 53,333 + 66,667


def test_comparison_ordinary_cgt_unchanged_by_reset(recalc_solution):
    """v1.5: Ord CGT row is the same value in both scenarios — pulls from Analyser."""
    assert _r(recalc_solution["scenario_a_ord_cgt"]) == 210_000


def test_comparison_total_tax_burden(recalc_solution):
    """v1.5: TOTAL TAX BURDEN = ord CGT + Div 296 tax for each scenario."""
    s = recalc_solution
    assert _r(s["scenario_a_burden"]) == 367_500   # 210,000 + 157,500
    assert _r(s["scenario_b_burden"]) == 238_500   # 210,000 + 28,500


def test_build_validate_recalc_finds_no_errors(tmp_path):
    """Build-time recalc gate: no cell resolves to #REF!/#DIV/0!/#VALUE!/etc."""
    from div296.build import build_workbook, validate_recalc

    out = tmp_path / "model.xlsx"
    build_workbook().save(out)
    errors = validate_recalc(out)
    assert errors == [], f"Expected no Excel error cells, got: {errors[:5]}"
