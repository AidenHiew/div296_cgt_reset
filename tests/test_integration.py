"""Integration test: build → recalc with `formulas` → assert §12 totals.

This is the strongest verification gate we have: the live Excel formulas
in the built `.xlsx` are recalculated by the pure-Python `formulas`
package and the resulting numbers are asserted against the v3.1 §12
acceptance numbers (see [tests/test_calcs.py](tests/test_calcs.py) for
the calc-engine truth source).

Updated for v3.0/v3.1 layout: cell references are derived from the
layout constants in `analyser.py` / `comparison.py` rather than
hardcoded, so this test won't drift again when rows shift.

Note on acceptance numbers
--------------------------
`test_calcs.py` uses a SINGLE-MEMBER fixture (sole TSB $12m) for §12.
The built workbook seeds TWO members (M1=$12m, M2=$3.5m, total $15.5m)
which spreads the Div 296 earnings across both members' band1/band2
shares. So the headline tax figures here differ from those in
`test_calcs.py`. Both are correct for their respective member configs —
this test is the live-workbook truth source.

Known `formulas` package limitation
-----------------------------------
The v3.1 Reconciliation helpers (M70/N70/O70) use a SUMPRODUCT pattern
with boolean array math that the pure-Python `formulas` recalc engine
cannot evaluate (it returns #VALUE!). This cascades to B71 (Fund
Ord CGT), B73 (CF losses), and Comparison rows that pull from B71. Real
Excel and LibreOffice compute these correctly — verified by hand and by
the unit tests in `test_calcs.py`. Cells in `KNOWN_FORMULAS_LIMITATIONS`
are excluded from the recalc-validate assertion below.

Skips cleanly (with an explanatory message) if `formulas` isn't
installed — the rest of the suite still runs. Also skips on MemoryError
since the per-register sort grid (cols N/O/P × 50 rows of LARGE/MATCH/
INDEX) can exceed what the pure-Python `formulas` engine holds.
"""

from __future__ import annotations


import numpy as np
import pytest
from openpyxl.utils import get_column_letter

from div296._recalc_limitations import KNOWN_FORMULAS_LIMITATIONS
from div296.build import build_workbook
from div296.tabs import analyser as A_TAB
from div296.tabs import comparison as C_TAB


formulas = pytest.importorskip(
    "formulas",
    reason="`formulas` package required for live recalc test — pip install -e .[dev]",
)

# Mark every test in this module as `slow` so dev runs can opt out:
#     pytest -m "not slow"
pytestmark = pytest.mark.slow


# --- helpers --------------------------------------------------------------

def _unwrap(value):
    """`formulas` returns 2-D ndarray cells; unwrap to a scalar float.

    Some cells legitimately resolve to strings (e.g. the empty-string
    fallback `IF(...,"")` or the "—" placeholder), or to XlError objects
    when the `formulas` package can't evaluate the formula (see
    KNOWN_FORMULAS_LIMITATIONS). Return as-is in those cases so the
    caller can detect and skip.
    """
    if isinstance(value, np.ndarray):
        value = value.flat[0]
    if isinstance(value, str):
        return value
    try:
        return float(value)
    except (TypeError, ValueError):
        return value  # e.g. XlError → caller will skip


def _col(idx: int) -> str:
    return get_column_letter(idx)


# Derived cell coordinates — single source of truth, computed from the
# tab layout constants so this test stays correct under future row shifts.

# Analyser
A_HEADLINE_NORESET = f"{_col(A_TAB.FUND_NORESET_COL)}{A_TAB.HEADLINE_ROW}"   # C13
A_HEADLINE_ELECTED = f"{_col(A_TAB.FUND_ELECTED_COL)}{A_TAB.HEADLINE_ROW}"   # D13
A_RECON_DIV296 = f"B{A_TAB.RECON_DIV296_ROW}"          # B72
A_TOTAL_PROCEEDS = f"{_col(A_TAB.PROCEEDS_COL)}{A_TAB.TOTALS_ROW}"        # C67
A_TOTAL_DIV296_GAIN = f"{_col(A_TAB.DIV_GAIN_COL)}{A_TAB.TOTALS_ROW}"     # H67
A_TOTAL_DIV296_TAX = f"{_col(A_TAB.DIV_TAX_COL)}{A_TAB.TOTALS_ROW}"       # I67

# Comparison subtotals — cols: A label, B no-reset, C elected, D diff
C_EARN_A = f"B{C_TAB.SUBTOTAL_EARNINGS_ROW}"           # B27
C_EARN_B = f"C{C_TAB.SUBTOTAL_EARNINGS_ROW}"           # C27
C_DIV296_A = f"B{C_TAB.SUBTOTAL_DIV296_ROW}"           # B29
C_DIV296_B = f"C{C_TAB.SUBTOTAL_DIV296_ROW}"           # C29

# Comparison "Net effect" headline card — third card spans I:K, value row = CARD_VALUE_ROW.
C_NET_EFFECT = f"I{C_TAB.CARD_VALUE_ROW}"              # I22 (signed: elected − no-reset)



@pytest.fixture(scope="module")
def recalc_solution(tmp_path_factory) -> dict:
    out = tmp_path_factory.mktemp("recalc") / "model.xlsx"
    wb = build_workbook()
    wb.save(out)

    try:
        xl = formulas.ExcelModel().loads(str(out)).finish()
        sol = xl.calculate()
    except MemoryError:
        # Comparison's per-register sort grid (cols N/O/P with 50 rows of
        # LARGE/MATCH/INDEX) exceeds what the pure-Python `formulas` engine
        # holds on modest hardware. Verified via LibreOffice headless render
        # — see PDF export script.
        pytest.skip(
            "`formulas` package OOM'd recalculating the v3.1 workbook. "
            "Verify §12 numbers via Excel/LibreOffice manually."
        )

    file_token = f"[{out.name}]"

    def at(sheet: str, cell: str):
        key = f"'{file_token}{sheet.upper()}'!{cell}"
        return _unwrap(sol[key].value)

    return {
        "sol": sol,
        "file_token": file_token,
        # Analyser — fund summary (side-by-side scenarios on row 13)
        "headline_div296_tax_noreset": at("Analyser", A_HEADLINE_NORESET),
        "headline_div296_tax_elected": at("Analyser", A_HEADLINE_ELECTED),
        # Analyser — Reconciliation (Div 296 row only; Ord CGT + CF losses
        # are in KNOWN_FORMULAS_LIMITATIONS).
        "div296_tax_payable": at("Analyser", A_RECON_DIV296),
        # Analyser — totals row 67 (col F omitted in v3.1; info-only)
        "totals_proceeds": at("Analyser", A_TOTAL_PROCEEDS),
        "totals_div296_adj_gain": at("Analyser", A_TOTAL_DIV296_GAIN),
        "totals_div296_tax": at("Analyser", A_TOTAL_DIV296_TAX),
        # Comparison — per-scenario subtotals (rows 27-30; Ord CGT + burden
        # are in KNOWN_FORMULAS_LIMITATIONS).
        "scenario_a_earnings": at("Comparison", C_EARN_A),
        "scenario_b_earnings": at("Comparison", C_EARN_B),
        "scenario_a_tax": at("Comparison", C_DIV296_A),
        "scenario_b_tax": at("Comparison", C_DIV296_B),
        # Comparison — net-effect card (signed: elected − no-reset)
        "net_effect_signed": at("Comparison", C_NET_EFFECT),
    }


# --- §12 acceptance assertions ------------------------------------------
#
# Acceptance numbers below are computed for the seeded workbook (M1=$12m,
# M2=$3.5m, total $15.5m TSB; §12 three-asset register). Derivation:
#
#   Total TSB = 15.5m
#   M1: split = 12/15.5 = 0.7742;  band1 = 7/12;  band2 = 2/12
#   M2: split =  3.5/15.5 = 0.2258; band1 = 0.5/3.5 ≈ 0.1429;  band2 = 0
#
#   Per-member rate factor (split × (band1×15% + band2×25%)):
#     M1: 0.7742 × (7/12×15% + 2/12×25%) = 0.7742 × 0.12917 = 0.10000
#     M2: 0.2258 × (0.5/3.5 × 15%)       = 0.2258 × 0.02143 = 0.00484
#     Combined fund factor:                                ≈ 0.10484
#
#   No-reset earnings (v3.1, intra-year netted): 1,200,000 + 200,000 − 300,000 = 1,100,000
#   No-reset headline tax: 1,100,000 × 0.10484 = 115,323
#   Elected earnings (cost base = MV): 133,333 + 53,333 + 66,667 = 253,333
#   Elected headline tax: 253,333 × 0.10484 =                       26,559
#   Net effect (signed, elected − no-reset): 26,559 − 115,323 =    −88,763

def _r(x) -> int:
    if not isinstance(x, (int, float)):
        pytest.fail(f"Expected numeric, got {type(x).__name__}: {x!r}")
    return int(round(x))


def test_headline_div296_tax_elected_26559(recalc_solution):
    assert _r(recalc_solution["headline_div296_tax_elected"]) == 26_559


def test_headline_div296_tax_noreset_115323(recalc_solution):
    assert _r(recalc_solution["headline_div296_tax_noreset"]) == 115_323


def test_recon_div296_tax_payable_matches_headline(recalc_solution):
    """Recon panel echoes the elected-reset headline."""
    assert _r(recalc_solution["div296_tax_payable"]) == 26_559


def test_analyser_totals(recalc_solution):
    s = recalc_solution
    assert _r(s["totals_proceeds"]) == 3_400_000          # 2.6m + 0.6m + 0.2m
    assert _r(s["totals_div296_adj_gain"]) == 253_333     # elected-reset scenario
    assert _r(s["totals_div296_tax"]) == 26_559           # pro-rata sum ties to headline


def test_comparison_scenario_a_no_reset_115323(recalc_solution):
    """v3.1: no-reset Div 296 tax for 2-member fund = $115,323."""
    assert _r(recalc_solution["scenario_a_tax"]) == 115_323


def test_comparison_scenario_b_reset_26559(recalc_solution):
    assert _r(recalc_solution["scenario_b_tax"]) == 26_559


def test_comparison_net_effect_signed(recalc_solution):
    """Card is SIGNED (elected − no-reset): 26,559 − 115,323 = −88,763.
    Negative = electing the reset reduces total Div 296 tax."""
    assert _r(recalc_solution["net_effect_signed"]) == -88_763


def test_comparison_earnings_subtotals(recalc_solution):
    """v3.1: MAX(0, SUM(...)) — intra-year netting, floored at zero."""
    s = recalc_solution
    assert _r(s["scenario_a_earnings"]) == 1_100_000   # 1.2m + 0.2m − 0.3m
    assert _r(s["scenario_b_earnings"]) == 253_333     # 133,333 + 53,333 + 66,667


def test_build_validate_recalc_finds_no_unexpected_errors(tmp_path):
    """Build-time recalc gate: no UNEXPECTED cell error.

    The v3.1 Reconciliation SUMPRODUCT helpers (M70/N70/O70) and the cells
    that depend on them are filtered out — they're a known `formulas`
    package limitation (boolean-array SUMPRODUCT not supported), not a real
    Excel-side bug. See module docstring.
    """
    from div296.build import build_workbook, validate_recalc

    out = tmp_path / "model.xlsx"
    build_workbook().save(out)
    try:
        errors = validate_recalc(out)
    except MemoryError:
        pytest.skip(
            "`formulas` package OOM'd recalculating the v3.1 workbook for "
            "validate_recalc. Same root cause as the fixture skip path."
        )

    unexpected = [
        e for e in errors
        if not any(skip in e for skip in KNOWN_FORMULAS_LIMITATIONS)
    ]
    assert unexpected == [], (
        f"Expected no Excel error cells beyond the SUMPRODUCT helpers, "
        f"got: {unexpected[:5]}"
    )
