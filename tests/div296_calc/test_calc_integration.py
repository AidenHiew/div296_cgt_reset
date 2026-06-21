"""Slow integration: build → recalc with `formulas` → assert sample-fund cells.

The shipped coherent sample fund (Alice pooled 60% / $4M; Bob pooled 40% /
$12.9M; one pool) is the live-workbook truth source. The four kernel goldens
(Jess/Emma/Jill/Jamal) are pure-layer unit tests in test_calc_engine.py — the
two-tier path is exercised live here via Bob's >$10M balance.

Hand-derivation:
  CGT helper: gains>12m 300k, losses 60k → net = (300k−60k)×2/3 = 160,000
  POOLED TOTAL = 120k+30k+80k+0+160k − 40k = 350,000
  Alice: earnings 0.6×350k = 210,000; TSB_REF 4.0M; band1 = 1/4 = .25
         tax = 210,000 × .25 × .15 = 7,875
  Bob:   earnings 0.4×350k = 140,000; TSB_REF 12.9M
         band1 = 7/12.9, band2 = 2.9/12.9
         tax = 140,000 × (7/12.9×.15 + 2.9/12.9×.25) ≈ 19,263.57
  FUND TOTAL ≈ 27,138.57
"""

from __future__ import annotations

import numpy as np
import pytest

from div296_calc.build import build_workbook
from div296_calc.tabs import calculator as C

formulas = pytest.importorskip(
    "formulas", reason="`formulas` required — pip install -e .[dev]")
pytestmark = pytest.mark.slow


def _unwrap(v):
    if isinstance(v, np.ndarray):
        v = v.flat[0]
    if isinstance(v, str):
        return v
    try:
        return float(v)
    except (TypeError, ValueError):
        return v


@pytest.fixture(scope="module")
def sol(tmp_path_factory):
    out = tmp_path_factory.mktemp("calc") / "calc.xlsx"
    build_workbook().save(out)
    xl = formulas.ExcelModel().loads(str(out)).finish()
    s = xl.calculate()
    token = f"[{out.name}]"

    def at(cell):
        return _unwrap(s[f"'{token}CALCULATOR'!{cell}"].value)

    return at


def _r(x):
    assert isinstance(x, (int, float)), f"non-numeric: {x!r}"
    return int(round(x))


def test_net_capital_gain_cell(sol):
    assert _r(sol(f"B{C.NET_CG_ROW}")) == 160_000


def test_pooled_total_cell(sol):
    assert _r(sol(f"B{C.POOLED_TOTAL_ROW}")) == 350_000


def test_alice_total_tax(sol):
    assert _r(sol(f"B{C.TOTAL_TAX_ROW}")) == 7_875


def test_bob_two_tier_total_tax(sol):
    assert _r(sol(f"C{C.TOTAL_TAX_ROW}")) == 19_264
    assert _r(sol(f"C{C.TIER2_TAX_ROW}")) > 0       # >$10M slice taxed live


def test_fund_total_tax(sol):
    assert _r(sol(f"B{C.FUND_TOTAL_ROW}")) == 27_139


def test_recalc_gate_zero_error_cells(tmp_path):
    from div296_calc.build import build_workbook, validate_recalc
    out = tmp_path / "calc.xlsx"
    build_workbook().save(out)
    assert validate_recalc(out) == []
