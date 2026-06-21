"""Pure-engine tests for div296_calc.calcs."""

import pytest

from div296_calc.assumptions import RATE_TIER1, RATE_TIER2
from div296_calc.calcs import member_div296_tax, tsb_ref

T1, T2 = 3_000_000, 10_000_000


def _tax(net, tsb):
    return member_div296_tax(net, tsb, T1, T2, RATE_TIER1, RATE_TIER2)


# --- mandatory >$10M two-tier anchor (the footgun guard) ---
def test_emma_two_tier_above_10m():
    assert _tax(840_000, 12_900_000) == pytest.approx(115_581.40, abs=0.01)


def test_jess_single_tier():
    assert _tax(476_625, 4_500_000) == pytest.approx(23_831.25, abs=0.01)


def test_jill_just_above_3m():
    assert _tax(100_000, 3_100_000) == pytest.approx(483.87, abs=0.01)


# --- boundaries ---
def test_tsb_exactly_3m_is_zero():
    assert _tax(500_000, 3_000_000) == 0.0


def test_tsb_exactly_10m_is_tier1_only():
    # band1 = (10M-3M)/10M = 0.7, band2 = 0
    assert _tax(100_000, 10_000_000) == pytest.approx(100_000 * 0.7 * 0.15)


def test_tsb_below_3m_with_earnings_is_zero():
    assert _tax(100_000, 2_500_000) == 0.0


def test_zero_tsb_is_zero():
    assert _tax(100_000, 0) == 0.0


def test_zero_or_negative_net_is_zero():
    assert _tax(0, 4_000_000) == 0.0
    assert _tax(-50_000, 4_000_000) == 0.0


# --- TSB_REF rule ---
def test_tsb_ref_2026_27_uses_closing():
    from div296_calc.calcs import Member
    m = Member(name="A", opening_tsb=5_000_000, closing_tsb=4_000_000, share=1.0)
    assert tsb_ref(m, use_greater_of=False) == 4_000_000


def test_tsb_ref_greater_of_when_opening_higher():
    from div296_calc.calcs import Member
    m = Member(name="A", opening_tsb=5_000_000, closing_tsb=4_000_000, share=1.0)
    assert tsb_ref(m, use_greater_of=True) == 5_000_000


def test_tsb_ref_greater_of_when_closing_higher():
    from div296_calc.calcs import Member
    m = Member(name="A", opening_tsb=4_000_000, closing_tsb=6_000_000, share=1.0)
    assert tsb_ref(m, use_greater_of=True) == 6_000_000
