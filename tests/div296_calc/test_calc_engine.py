"""Pure-engine tests for div296_calc.calcs."""

import pytest

from div296_calc.assumptions import RATE_TIER1, RATE_TIER2
from div296_calc.calcs import Member, member_div296_tax, tsb_ref

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


# --- CGT netting helper (s102-5) ---
from div296_calc.assumptions import DISCOUNT_RATE  # noqa: E402
from div296_calc.calcs import CgtInputs, CgtResult, net_capital_gain  # noqa: E402


def test_cgt_long_held_gain_gets_one_third_discount():
    r = net_capital_gain(CgtInputs(gross_gains_over_12m=300_000,
                                   gross_gains_under_12m=0,
                                   capital_losses=0), DISCOUNT_RATE)
    assert r.net_realised_cg == pytest.approx(200_000)   # 300k × 2/3
    assert r.unused_capital_loss == 0


def test_cgt_losses_hit_non_discount_gains_first():
    # 60k loss eats the 50k short-held gain, then 10k of the long-held gain.
    # remaining long-held = 240k → ×2/3 = 160k.
    r = net_capital_gain(CgtInputs(gross_gains_over_12m=250_000,
                                   gross_gains_under_12m=50_000,
                                   capital_losses=60_000), DISCOUNT_RATE)
    assert r.net_realised_cg == pytest.approx(160_000)
    assert r.unused_capital_loss == 0


def test_cgt_losses_exceeding_all_gains_carry_forward():
    r = net_capital_gain(CgtInputs(gross_gains_over_12m=100_000,
                                   gross_gains_under_12m=20_000,
                                   capital_losses=200_000), DISCOUNT_RATE)
    assert r.net_realised_cg == 0
    assert r.unused_capital_loss == pytest.approx(80_000)   # 200k − 120k


def test_cgt_result_is_frozen_dataclass():
    r = net_capital_gain(CgtInputs(0, 0, 0), DISCOUNT_RATE)
    assert isinstance(r, CgtResult)


# --- pooled total + per-member allocation ---
from div296_calc.calcs import PooledIncome, member_earnings, pooled_total  # noqa: E402


def test_pooled_total_sums_income_less_expenses():
    p = PooledIncome(dividends_grossed=120_000, interest=30_000, rent=80_000,
                     other=0, net_realised_cg=160_000, deductible_expenses=40_000)
    assert pooled_total(p) == 350_000


def test_pooled_total_can_go_negative():
    p = PooledIncome(dividends_grossed=10_000, interest=0, rent=0, other=0,
                     net_realised_cg=0, deductible_expenses=60_000)
    assert pooled_total(p) == -50_000


def test_member_earnings_pooled_is_share_times_pool():
    m = Member("A", 4_000_000, 4_000_000, share=0.6)
    assert member_earnings(m, 350_000) == pytest.approx(210_000)


def test_member_earnings_override_wins_and_ignores_pool():
    m = Member("A", 4_000_000, 4_000_000, share=0.6, override=90_000)
    assert member_earnings(m, 350_000) == 90_000


def test_member_earnings_override_zero_is_honoured_not_blank():
    m = Member("A", 4_000_000, 4_000_000, share=0.6, override=0.0)
    assert member_earnings(m, 350_000) == 0.0


def test_member_earnings_negative_pool_allocates_negative():
    m = Member("A", 4_000_000, 4_000_000, share=0.5)
    assert member_earnings(m, -50_000) == pytest.approx(-25_000)


# --- per-member result + fund roll-up + share guard ---
from div296_calc.assumptions import thresholds_for  # noqa: E402
from div296_calc.calcs import (  # noqa: E402
    compute_fund,
    compute_member,
    fund_total_tax,
    pooled_share_status,
)

YT = thresholds_for("2026-27")


def _member_result(m, pool):
    return compute_member(m, pool, YT, RATE_TIER1, RATE_TIER2)


def test_compute_member_below_threshold_status():
    m = Member("Below", 2_500_000, 2_500_000, share=1.0)
    r = _member_result(m, 100_000)
    assert r.tax == 0.0
    assert r.below_threshold is True
    assert r.new_loss == 0.0


def test_compute_member_emma_via_override():
    m = Member("Emma", 12_900_000, 12_900_000, share=0.0, override=840_000)
    r = _member_result(m, 0)
    assert r.tax == pytest.approx(115_581.40, abs=0.01)
    assert r.tier2_tax > 0          # the >$10M slice is taxed
    assert r.below_threshold is False


def test_compute_member_prior_loss_partially_absorbs():
    m = Member("P", 4_000_000, 4_000_000, share=1.0, prior_loss=100_000)
    r = _member_result(m, 200_000)            # net = 100,000
    assert r.net_earnings == pytest.approx(100_000)
    assert r.tax == pytest.approx(100_000 * 0.25 * 0.15)   # band1 = 1M/4M


def test_compute_member_prior_loss_exceeds_earnings_carries_forward():
    m = Member("L", 4_000_000, 4_000_000, share=1.0, prior_loss=250_000)
    r = _member_result(m, 200_000)            # net = -50,000
    assert r.tax == 0.0
    assert r.new_loss == pytest.approx(50_000)


def test_jamal_negative_earnings_carry_forward():
    m = Member("Jamal", 5_000_000, 5_000_000, share=1.0, override=-200_000)
    r = _member_result(m, 0)
    assert r.tax == 0.0
    assert r.new_loss == pytest.approx(200_000)


def test_compute_fund_rolls_up_sum_including_zero_member():
    members = [
        Member("Above", 4_000_000, 4_000_000, share=0.5),
        Member("Below", 2_000_000, 2_000_000, share=0.5),
    ]
    results = compute_fund(members, 400_000, YT, RATE_TIER1, RATE_TIER2)
    assert len(results) == 2
    # Above: earnings 200k, band1 = 1M/4M = .25 → 200k×.25×.15 = 7,500
    assert fund_total_tax(results) == pytest.approx(7_500)


def test_pooled_share_status_100_percent_ok():
    members = [Member("A", 4e6, 4e6, share=0.6), Member("B", 5e6, 5e6, share=0.4)]
    status = pooled_share_status(members)
    assert status.pooled_count == 2
    assert status.share_sum == pytest.approx(1.0)
    assert status.ok is True
    assert status.all_segregated is False


def test_pooled_share_status_soft_warn_when_not_100():
    members = [Member("A", 4e6, 4e6, share=0.6), Member("B", 5e6, 5e6, share=0.3)]
    status = pooled_share_status(members)
    assert status.ok is False


def test_pooled_share_status_excludes_overridden_members():
    members = [
        Member("Pooled", 4e6, 4e6, share=1.0),
        Member("Seg", 5e6, 5e6, share=0.0, override=50_000),
    ]
    status = pooled_share_status(members)
    assert status.pooled_count == 1
    assert status.share_sum == pytest.approx(1.0)
    assert status.ok is True


def test_pooled_share_status_all_overridden_is_suppressed():
    members = [
        Member("S1", 4e6, 4e6, share=0.0, override=10_000),
        Member("S2", 5e6, 5e6, share=0.0, override=20_000),
    ]
    status = pooled_share_status(members)
    assert status.pooled_count == 0
    assert status.all_segregated is True
    assert status.ok is True          # no false warning
