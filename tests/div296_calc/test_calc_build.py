"""Structural tests for the built workbook (fast — no recalc)."""

import pytest

from div296_calc import _formulas as F
from div296_calc.build import build_workbook
from div296_calc.tabs import calculator as C


@pytest.fixture(scope="module")
def wb():
    return build_workbook()


def test_two_tabs_in_order(wb):
    assert wb.sheetnames == ["Calculator", "Notes"]


def test_named_ranges_defined(wb):
    for name in ("rate_tier1", "rate_tier2", "discount_rate",
                 "t1_sel", "t2_sel", "greater_of_sel"):
        assert name in wb.defined_names


def test_pooled_total_cell_formula(wb):
    ws = wb["Calculator"]
    c = ws[f"B{C.POOLED_TOTAL_ROW}"].value
    assert c == F.pooled_total_formula(
        f"B{C.DIVIDENDS_ROW}", f"B{C.INTEREST_ROW}", f"B{C.RENT_ROW}",
        f"B{C.OTHER_ROW}", f"B{C.NET_CG_ROW}", f"B{C.EXPENSES_ROW}",
    )


def test_net_cg_cell_pulls_from_cgt_helper(wb):
    ws = wb["Calculator"]
    assert ws[f"B{C.NET_CG_ROW}"].value == F.cgt_net_formula(
        f"B{C.CGT_OVER_ROW}", f"B{C.CGT_UNDER_ROW}", f"B{C.CGT_LOSS_ROW}")


def test_member1_tier2_tax_formula(wb):
    ws = wb["Calculator"]
    col = "B"   # member 1
    assert ws[f"{col}{C.TIER2_TAX_ROW}"].value == F.tier2_tax_formula(
        f"{col}{C.NET_EARNINGS_ROW}", f"{col}{C.BAND2_ROW}",
        f"{col}{C.TSB_REF_ROW}", "t1_sel")


def test_fund_total_is_sum_of_member_tax(wb):
    ws = wb["Calculator"]
    val = ws[f"B{C.FUND_TOTAL_ROW}"].value
    assert val == f"=SUM(B{C.TOTAL_TAX_ROW}:E{C.TOTAL_TAX_ROW})"


def test_sample_fund_seeded(wb):
    ws = wb["Calculator"]
    # Alice (col B) 60% / $4M closing; Bob (col C) 40% / $12.9M closing.
    assert ws[f"B{C.CLOSING_TSB_ROW}"].value == 4_000_000
    assert ws[f"C{C.CLOSING_TSB_ROW}"].value == 12_900_000
    assert ws[f"B{C.SHARE_ROW}"].value == 0.6
    assert ws[f"C{C.SHARE_ROW}"].value == 0.4
    assert ws[f"B{C.DIVIDENDS_ROW}"].value == 120_000


def test_sheet_protected_with_unlocked_inputs(wb):
    ws = wb["Calculator"]
    assert ws.protection.sheet is True
    # A seeded input cell is unlocked; a result cell stays locked.
    assert ws[f"B{C.CLOSING_TSB_ROW}"].protection.locked is False
    assert ws[f"B{C.TOTAL_TAX_ROW}"].protection.locked is True


def test_notes_tab_has_law_basis_and_caveats(wb):
    text = "\n".join(
        str(c.value) for row in wb["Notes"].iter_rows() for c in row if c.value)
    assert "Royal Assent" in text
    assert "thresholds verified" in text.lower()
    assert "not advice" in text.lower() or "not personal advice" in text.lower()
    assert "carry-forward" in text.lower() or "carried forward" in text.lower()


def test_build_main_writes_file(tmp_path):
    from div296_calc.build import main
    out = tmp_path / "calc.xlsx"
    rc = main(["-o", str(out), "--no-validate"])
    assert rc == 0
    assert out.exists()
