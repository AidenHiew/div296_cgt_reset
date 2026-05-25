"""Spec §12 acceptance-number tests.

Locked scenario: single member, TSB $12,000,000, all assets held > 12mo,
discount ON, $10m tier OFF. Three assets per §12 table.

All expected numbers are whole-dollar rounded as per §12 (the dollar
amounts in the spec are rounded; we round our floats the same way).
"""

from __future__ import annotations

import pytest

from div296 import __version__
from div296.assumptions import ASSUMPTIONS
from div296.calcs import (
    Asset,
    Member,
    carry_forward_loss,
    div296_adjusted_gain,
    div296_fund_earnings,
    div296_headline_tax,
    ordinary_cgt,
    ordinary_taxable_gain,
    per_asset_div296_tax,
)


A = ASSUMPTIONS


# --- §12 sample register --------------------------------------------------

PROPERTY = Asset(
    code="P1",
    name="Commercial property",
    quantity=1,
    original_cost_base=800_000,
    current_market_value=2_400_000,
    market_value_30jun2026=2_400_000,
    valuation_source="Independent val, 30/06/26",
    projected_sale_proceeds=2_600_000,
    held_over_12_months=True,
)

SHARES = Asset(
    code="S1",
    name="Listed shares parcel",
    quantity=5_000,
    original_cost_base=300_000,
    current_market_value=520_000,
    market_value_30jun2026=520_000,
    valuation_source="ASX close 30/06/26",
    projected_sale_proceeds=600_000,
    held_over_12_months=True,
)

LOSS = Asset(
    code="L1",
    name="Loss-making holding",
    quantity=2_000,
    original_cost_base=500_000,
    current_market_value=100_000,
    market_value_30jun2026=100_000,
    valuation_source="Independent val, 30/06/26",
    projected_sale_proceeds=200_000,
    held_over_12_months=True,
)

REGISTER = [PROPERTY, SHARES, LOSS]

SOLE_MEMBER = Member(tsb=12_000_000, split_pct=1.0)
MEMBERS = [SOLE_MEMBER]


# --- helpers --------------------------------------------------------------

def _r(x: float) -> int:
    """Whole-dollar rounding matching §12 (banker's rounding via round())."""
    return int(round(x))


# --- package sanity -------------------------------------------------------

def test_package_version():
    assert __version__ == "2.3.0"


# --- §12 scenario: reset ON, discount ON, tier OFF -----------------------

class TestResetOn:
    reset_on = True
    discount_on = True
    tier10_on = False

    def test_member_proportion(self):
        # (12m - 3m) / 12m = 0.75
        from div296.calcs import member_proportion_above_3m
        assert member_proportion_above_3m(SOLE_MEMBER, A.threshold_1) == pytest.approx(0.75)

    # Ordinary taxable gain (col 5)
    def test_property_ordinary_taxable_gain(self):
        # (2,600,000 - 800,000) × 2/3 = 1,200,000
        v = ordinary_taxable_gain(PROPERTY, self.discount_on, A.discount_rate)
        assert _r(v) == 1_200_000

    def test_shares_ordinary_taxable_gain(self):
        # (600,000 - 300,000) × 2/3 = 200,000
        v = ordinary_taxable_gain(SHARES, self.discount_on, A.discount_rate)
        assert _r(v) == 200_000

    def test_loss_ordinary_taxable_gain_is_full_loss(self):
        # (200,000 - 500,000) = -300,000, losses not discounted
        v = ordinary_taxable_gain(LOSS, self.discount_on, A.discount_rate)
        assert _r(v) == -300_000

    # Ordinary CGT (col 6) — per-asset silo
    def test_property_ordinary_cgt(self):
        v = ordinary_cgt(PROPERTY, self.discount_on, A.discount_rate, A.fund_cgt_rate)
        assert _r(v) == 180_000

    def test_shares_ordinary_cgt(self):
        v = ordinary_cgt(SHARES, self.discount_on, A.discount_rate, A.fund_cgt_rate)
        assert _r(v) == 30_000

    def test_loss_ordinary_cgt_is_zero(self):
        v = ordinary_cgt(LOSS, self.discount_on, A.discount_rate, A.fund_cgt_rate)
        assert _r(v) == 0

    def test_total_ordinary_cgt_payable(self):
        total = sum(
            ordinary_cgt(a, self.discount_on, A.discount_rate, A.fund_cgt_rate)
            for a in REGISTER
        )
        assert _r(total) == 210_000

    # Div 296 adjusted gain (col 7) — reset ON, cost base = MV 30 Jun 2026
    def test_property_div296_adjusted_gain(self):
        # (2,600,000 - 2,400,000) × 2/3 = 133,333.33...
        v = div296_adjusted_gain(PROPERTY, self.reset_on, self.discount_on, A.discount_rate)
        assert _r(v) == 133_333

    def test_shares_div296_adjusted_gain(self):
        # (600,000 - 520,000) × 2/3 = 53,333.33...
        v = div296_adjusted_gain(SHARES, self.reset_on, self.discount_on, A.discount_rate)
        assert _r(v) == 53_333

    def test_loss_div296_adjusted_gain_is_the_trap(self):
        # (200,000 - 100,000) × 2/3 = 66,666.67  ← reset creates a Div 296 gain
        v = div296_adjusted_gain(LOSS, self.reset_on, self.discount_on, A.discount_rate)
        assert _r(v) == 66_667

    # Div 296 fund earnings
    def test_div296_fund_earnings_total(self):
        v = div296_fund_earnings(REGISTER, self.reset_on, self.discount_on, A.discount_rate)
        assert _r(v) == 253_333

    # Headline Div 296 tax (member-attributed)
    def test_div296_headline_tax(self):
        v = div296_headline_tax(
            REGISTER, MEMBERS,
            self.reset_on, self.discount_on, A.discount_rate,
            self.tier10_on, A.threshold_1, A.threshold_2,
            A.rate_tier1, A.rate_tier2,
        )
        assert _r(v) == 28_500

    # Per-asset Div 296 tax — pro-rata of headline
    def test_per_asset_div296_pro_rata_property(self):
        headline = div296_headline_tax(
            REGISTER, MEMBERS,
            self.reset_on, self.discount_on, A.discount_rate,
            self.tier10_on, A.threshold_1, A.threshold_2,
            A.rate_tier1, A.rate_tier2,
        )
        v = per_asset_div296_tax(
            PROPERTY, REGISTER, headline,
            self.reset_on, self.discount_on, A.discount_rate,
        )
        assert _r(v) == 15_000

    def test_per_asset_div296_pro_rata_shares(self):
        headline = div296_headline_tax(
            REGISTER, MEMBERS,
            self.reset_on, self.discount_on, A.discount_rate,
            self.tier10_on, A.threshold_1, A.threshold_2,
            A.rate_tier1, A.rate_tier2,
        )
        v = per_asset_div296_tax(
            SHARES, REGISTER, headline,
            self.reset_on, self.discount_on, A.discount_rate,
        )
        assert _r(v) == 6_000

    def test_per_asset_div296_pro_rata_loss(self):
        headline = div296_headline_tax(
            REGISTER, MEMBERS,
            self.reset_on, self.discount_on, A.discount_rate,
            self.tier10_on, A.threshold_1, A.threshold_2,
            A.rate_tier1, A.rate_tier2,
        )
        v = per_asset_div296_tax(
            LOSS, REGISTER, headline,
            self.reset_on, self.discount_on, A.discount_rate,
        )
        assert _r(v) == 7_500

    def test_per_asset_div296_sum_ties_to_headline(self):
        """Locked decision: pro-rata sum always reconciles to headline."""
        headline = div296_headline_tax(
            REGISTER, MEMBERS,
            self.reset_on, self.discount_on, A.discount_rate,
            self.tier10_on, A.threshold_1, A.threshold_2,
            A.rate_tier1, A.rate_tier2,
        )
        per_asset_sum = sum(
            per_asset_div296_tax(
                a, REGISTER, headline,
                self.reset_on, self.discount_on, A.discount_rate,
            )
            for a in REGISTER
        )
        assert per_asset_sum == pytest.approx(headline)

    # Carry-forward losses (reconciliation panel)
    def test_carry_forward_losses(self):
        total = sum(carry_forward_loss(a) for a in REGISTER)
        assert _r(total) == 300_000


# --- §12 scenario: reset OFF, discount ON, tier OFF ----------------------

class TestResetOff:
    reset_on = False
    discount_on = True
    tier10_on = False

    def test_property_div296_adjusted_gain(self):
        # (2,600,000 - 800,000) × 2/3 = 1,200,000
        v = div296_adjusted_gain(PROPERTY, self.reset_on, self.discount_on, A.discount_rate)
        assert _r(v) == 1_200_000

    def test_shares_div296_adjusted_gain(self):
        # (600,000 - 300,000) × 2/3 = 200,000
        v = div296_adjusted_gain(SHARES, self.reset_on, self.discount_on, A.discount_rate)
        assert _r(v) == 200_000

    def test_loss_div296_adjusted_gain_is_loss(self):
        # (200,000 - 500,000) = -300,000 raw, losses not discounted
        v = div296_adjusted_gain(LOSS, self.reset_on, self.discount_on, A.discount_rate)
        assert _r(v) == -300_000

    def test_div296_fund_earnings_floors_loss_asset(self):
        # 1,200,000 + 200,000 + max(0, -300,000) = 1,400,000
        v = div296_fund_earnings(REGISTER, self.reset_on, self.discount_on, A.discount_rate)
        assert _r(v) == 1_400_000

    def test_div296_headline_tax(self):
        # 1,400,000 × 75% × 15% = 157,500
        v = div296_headline_tax(
            REGISTER, MEMBERS,
            self.reset_on, self.discount_on, A.discount_rate,
            self.tier10_on, A.threshold_1, A.threshold_2,
            A.rate_tier1, A.rate_tier2,
        )
        assert _r(v) == 157_500


# --- Comparison footer self-check ----------------------------------------

def test_net_effect_of_electing_reset():
    """Spec §12 footer: 157,500 − 28,500 = 129,000."""
    common = dict(
        discount_on=True,
        discount_rate=A.discount_rate,
        tier10_on=False,
        threshold_1=A.threshold_1,
        threshold_2=A.threshold_2,
        rate_tier1=A.rate_tier1,
        rate_tier2=A.rate_tier2,
    )
    off = div296_headline_tax(REGISTER, MEMBERS, reset_on=False, **common)
    on = div296_headline_tax(REGISTER, MEMBERS, reset_on=True, **common)
    assert _r(off) == 157_500
    assert _r(on) == 28_500
    assert _r(off - on) == 129_000


# --- Tier 2 ON scenario (spec §7) ----------------------------------------

class TestTierOn:
    """Member with TSB > $10m, tier ON: tax splits across the two bands."""

    def test_tier_on_collapses_to_tier_off_when_tsb_under_10m(self):
        """TSB $5m → band2 = 0; tier ON tax equals tier OFF tax."""
        m = Member(tsb=5_000_000, split_pct=1.0)
        earnings = 100_000
        from div296.calcs import div296_tax_for_member
        off = div296_tax_for_member(
            earnings, m,
            tier10_on=False, threshold_1=A.threshold_1, threshold_2=A.threshold_2,
            rate_tier1=A.rate_tier1, rate_tier2=A.rate_tier2,
        )
        on = div296_tax_for_member(
            earnings, m,
            tier10_on=True, threshold_1=A.threshold_1, threshold_2=A.threshold_2,
            rate_tier1=A.rate_tier1, rate_tier2=A.rate_tier2,
        )
        assert _r(off) == _r(on)
        # Manual check: (5m-3m)/5m × 100k × 15% = 40% × 100k × 15% = 6,000
        assert _r(on) == 6_000

    def test_tier_on_with_tsb_above_10m(self):
        """TSB $15m: band1 = (15-10)+(10-3) ... wait, band1 = MIN(15,10)-3 = 7m / 15m;
        band2 = (15-10) / 15m = 5m / 15m. Tax = E × 7/15 × 15% + E × 5/15 × 25%."""
        m = Member(tsb=15_000_000, split_pct=1.0)
        earnings = 100_000
        from div296.calcs import div296_tax_for_member
        on = div296_tax_for_member(
            earnings, m,
            tier10_on=True, threshold_1=A.threshold_1, threshold_2=A.threshold_2,
            rate_tier1=A.rate_tier1, rate_tier2=A.rate_tier2,
        )
        expected = (
            100_000 * (7_000_000 / 15_000_000) * 0.15
            + 100_000 * (5_000_000 / 15_000_000) * 0.25
        )
        assert _r(on) == _r(expected)

    def test_tier_on_zero_tax_when_tsb_at_threshold_1(self):
        """TSB = $3m exactly: band1 = 0, band2 = 0 → tax = 0."""
        m = Member(tsb=3_000_000, split_pct=1.0)
        from div296.calcs import div296_tax_for_member
        on = div296_tax_for_member(
            100_000, m,
            tier10_on=True, threshold_1=A.threshold_1, threshold_2=A.threshold_2,
            rate_tier1=A.rate_tier1, rate_tier2=A.rate_tier2,
        )
        assert _r(on) == 0

    def test_zero_tsb_returns_zero_tax(self):
        m = Member(tsb=0, split_pct=1.0)
        from div296.calcs import div296_tax_for_member
        result = div296_tax_for_member(
            100_000, m,
            tier10_on=False, threshold_1=A.threshold_1, threshold_2=A.threshold_2,
            rate_tier1=A.rate_tier1, rate_tier2=A.rate_tier2,
        )
        assert result == 0.0
