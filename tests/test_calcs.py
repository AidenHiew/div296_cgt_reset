"""Spec §12 acceptance-number tests (re-baselined for v3.1 loss netting).

Locked scenario: single member, TSB $12,000,000, all assets held > 12mo.
Three assets per §12 table.

v3.0 API simplification (breaking vs v2.x):
- `discount_on` and `tier10_on` parameters removed throughout — the calc engine
  always applies the 1/3 discount when `held_over_12_months=True` and always
  uses the two-band tax formula.
- `Member.proportion_override` field and `member_proportion_above_3m` function
  removed.

v3.1 numerical change (breaking vs v3.0):
- Ordinary CGT and Div 296 fund earnings now NET capital gains and losses
  intra-year at the fund level (s102-5 ITAA 1997). The §12 scenario contains
  a $300,000 loss asset (L1), so the no-reset headline numbers shift:
    * Div 296 fund earnings (no reset): $1,400,000 → $1,100,000
    * Div 296 headline tax (no reset):  $180,833   → $142,083
    * Ordinary CGT payable (fund):      $210,000   → $180,000
    * Carry-forward losses (fund):      $300,000   → $0 (gains absorb)
  Reset-on numbers are unchanged (the loss asset becomes a Div 296 gain
  under reset, so no losses to net).

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
    carry_forward_loss_fund,
    div296_adjusted_gain,
    div296_fund_earnings,
    div296_headline_tax,
    div296_tax_for_member,
    ordinary_cgt,
    ordinary_cgt_fund,
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
    assert __version__ == "3.1.0"


# --- §12 scenario: reset ON (elected) ------------------------------------

class TestResetOn:
    """Reset elected: Div 296 cost base = MV at 30 Jun 2026.

    v3.0: tier_on is always True (toggle removed). Discount applies iff held>12mo
    (toggle removed). These tests verify the v2.6 acceptance numbers under the
    simplified API."""

    reset_on = True

    # Ordinary taxable gain (col 5)
    def test_property_ordinary_taxable_gain(self):
        # (2,600,000 - 800,000) × 2/3 = 1,200,000
        v = ordinary_taxable_gain(PROPERTY, A.discount_rate)
        assert _r(v) == 1_200_000

    def test_shares_ordinary_taxable_gain(self):
        # (600,000 - 300,000) × 2/3 = 200,000
        v = ordinary_taxable_gain(SHARES, A.discount_rate)
        assert _r(v) == 200_000

    def test_loss_ordinary_taxable_gain_is_full_loss(self):
        # (200,000 - 500,000) = -300,000, losses not discounted
        v = ordinary_taxable_gain(LOSS, A.discount_rate)
        assert _r(v) == -300_000

    # Ordinary CGT (per-asset, standalone diagnostic view — v3.1: NOT the real tax)
    def test_property_ordinary_cgt_standalone(self):
        v = ordinary_cgt(PROPERTY, A.discount_rate, A.fund_cgt_rate)
        assert _r(v) == 180_000

    def test_shares_ordinary_cgt_standalone(self):
        v = ordinary_cgt(SHARES, A.discount_rate, A.fund_cgt_rate)
        assert _r(v) == 30_000

    def test_loss_ordinary_cgt_standalone_is_zero(self):
        v = ordinary_cgt(LOSS, A.discount_rate, A.fund_cgt_rate)
        assert _r(v) == 0

    # Fund Ordinary CGT (v3.1: real tax after intra-year netting per s102-5)
    def test_fund_ordinary_cgt_payable(self):
        # gross gains (all held > 12m): $1.8m + $300k = $2.1m
        # gross losses: $300k
        # losses absorbed by discount gains (no non-discount gains): d_after = $1.8m
        # net taxable = $1.8m × 2/3 = $1.2m
        # CGT = $1.2m × 15% = $180,000
        v = ordinary_cgt_fund(REGISTER, A.discount_rate, A.fund_cgt_rate)
        assert _r(v) == 180_000

    # Div 296 adjusted gain (col 7) — reset ON, cost base = MV 30 Jun 2026
    def test_property_div296_adjusted_gain(self):
        # (2,600,000 - 2,400,000) × 2/3 = 133,333.33...
        v = div296_adjusted_gain(PROPERTY, self.reset_on, A.discount_rate)
        assert _r(v) == 133_333

    def test_shares_div296_adjusted_gain(self):
        # (600,000 - 520,000) × 2/3 = 53,333.33...
        v = div296_adjusted_gain(SHARES, self.reset_on, A.discount_rate)
        assert _r(v) == 53_333

    def test_loss_div296_adjusted_gain_is_the_trap(self):
        # (200,000 - 100,000) × 2/3 = 66,666.67  ← reset creates a Div 296 gain
        v = div296_adjusted_gain(LOSS, self.reset_on, A.discount_rate)
        assert _r(v) == 66_667

    # Div 296 fund earnings
    def test_div296_fund_earnings_total(self):
        v = div296_fund_earnings(REGISTER, self.reset_on, A.discount_rate)
        assert _r(v) == 253_333

    # Headline Div 296 tax (member-attributed)
    def test_div296_headline_tax(self):
        # earnings = $253,333.33; two-band for $12m member:
        # band1 = ($10m-$3m)/$12m = 7/12;  band2 = ($12m-$10m)/$12m = 2/12.
        # tax = $253,333.33 × (7/12×15% + 2/12×25%) = $32,722.22
        v = div296_headline_tax(
            REGISTER, MEMBERS,
            self.reset_on, A.discount_rate,
            A.threshold_1, A.threshold_2,
            A.rate_tier1, A.rate_tier2,
        )
        assert _r(v) == 32_722

    # Per-asset Div 296 tax — pro-rata of headline
    def test_per_asset_div296_pro_rata_property(self):
        headline = div296_headline_tax(
            REGISTER, MEMBERS,
            self.reset_on, A.discount_rate,
            A.threshold_1, A.threshold_2,
            A.rate_tier1, A.rate_tier2,
        )
        v = per_asset_div296_tax(
            PROPERTY, REGISTER, headline,
            self.reset_on, A.discount_rate,
        )
        # 133,333/253,333 × 32,722 = 17,222
        assert _r(v) == 17_222

    def test_per_asset_div296_pro_rata_shares(self):
        headline = div296_headline_tax(
            REGISTER, MEMBERS,
            self.reset_on, A.discount_rate,
            A.threshold_1, A.threshold_2,
            A.rate_tier1, A.rate_tier2,
        )
        v = per_asset_div296_tax(
            SHARES, REGISTER, headline,
            self.reset_on, A.discount_rate,
        )
        # 53,333/253,333 × 32,722 = 6,889
        assert _r(v) == 6_889

    def test_per_asset_div296_pro_rata_loss(self):
        headline = div296_headline_tax(
            REGISTER, MEMBERS,
            self.reset_on, A.discount_rate,
            A.threshold_1, A.threshold_2,
            A.rate_tier1, A.rate_tier2,
        )
        v = per_asset_div296_tax(
            LOSS, REGISTER, headline,
            self.reset_on, A.discount_rate,
        )
        # 66,667/253,333 × 32,722 = 8,611  (reset turns the loss asset into a Div 296 gain)
        assert _r(v) == 8_611

    def test_per_asset_div296_sum_ties_to_headline(self):
        """Locked decision: pro-rata sum always reconciles to headline."""
        headline = div296_headline_tax(
            REGISTER, MEMBERS,
            self.reset_on, A.discount_rate,
            A.threshold_1, A.threshold_2,
            A.rate_tier1, A.rate_tier2,
        )
        per_asset_sum = sum(
            per_asset_div296_tax(
                a, REGISTER, headline,
                self.reset_on, A.discount_rate,
            )
            for a in REGISTER
        )
        assert per_asset_sum == pytest.approx(headline)

    # Carry-forward losses — per-asset (informational sum) vs fund (real)
    def test_per_asset_carry_forward_losses_gross_sum(self):
        """Per-asset gross loss sum — INFORMATIONAL only (v3.1)."""
        total = sum(carry_forward_loss(a) for a in REGISTER)
        assert _r(total) == 300_000

    def test_fund_carry_forward_loss_is_zero_when_gains_absorb(self):
        """v3.1: fund-level carry-forward = MAX(0, gross_losses - gross_gains).
        §12 has $2.1m gross gains and $300k gross losses → gains absorb losses,
        carry-forward = $0."""
        v = carry_forward_loss_fund(REGISTER)
        assert _r(v) == 0


# --- §12 scenario: reset OFF (no election) -------------------------------

class TestResetOff:
    """No election: Div 296 cost base = original cost base.

    v3.0: same API as the elected scenario — only the reset_on bool flips."""

    reset_on = False

    def test_property_div296_adjusted_gain(self):
        # (2,600,000 - 800,000) × 2/3 = 1,200,000
        v = div296_adjusted_gain(PROPERTY, self.reset_on, A.discount_rate)
        assert _r(v) == 1_200_000

    def test_shares_div296_adjusted_gain(self):
        # (600,000 - 300,000) × 2/3 = 200,000
        v = div296_adjusted_gain(SHARES, self.reset_on, A.discount_rate)
        assert _r(v) == 200_000

    def test_loss_div296_adjusted_gain_is_loss(self):
        # (200,000 - 500,000) = -300,000 raw, losses not discounted
        v = div296_adjusted_gain(LOSS, self.reset_on, A.discount_rate)
        assert _r(v) == -300_000

    def test_div296_fund_earnings_nets_loss_against_gains(self):
        # v3.1: 1,200,000 + 200,000 + (-300,000) = 1,100,000 (intra-year netting)
        v = div296_fund_earnings(REGISTER, self.reset_on, A.discount_rate)
        assert _r(v) == 1_100_000

    def test_div296_headline_tax(self):
        # v3.1: 1,100,000 × (7/12 × 15% + 2/12 × 25%) = 142,083
        v = div296_headline_tax(
            REGISTER, MEMBERS,
            self.reset_on, A.discount_rate,
            A.threshold_1, A.threshold_2,
            A.rate_tier1, A.rate_tier2,
        )
        assert _r(v) == 142_083


# --- Comparison footer self-check ----------------------------------------

def test_net_effect_of_electing_reset():
    """v3.1 (intra-year netting): 142,083 − 32,722 = 109,361 saved by electing reset.

    Under no-reset, the §12 LOSS asset's $300k loss nets against the gains,
    dropping earnings from $1.4m (v3.0 floor) to $1.1m (v3.1 netted).
    Under reset, the loss asset becomes a $66,667 Div 296 gain — no losses
    to net, so the reset headline is unchanged from v3.0.
    """
    common = dict(
        discount_rate=A.discount_rate,
        threshold_1=A.threshold_1,
        threshold_2=A.threshold_2,
        rate_tier1=A.rate_tier1,
        rate_tier2=A.rate_tier2,
    )
    off = div296_headline_tax(REGISTER, MEMBERS, reset_on=False, **common)
    on = div296_headline_tax(REGISTER, MEMBERS, reset_on=True, **common)
    assert _r(off) == 142_083
    assert _r(on) == 32_722
    assert _r(off - on) == 109_361


# --- Two-band edge cases (spec §7) ---------------------------------------

class TestTwoBand:
    """Two-band tax behaviour. v3.0: always two-band (no toggle)."""

    def test_tsb_under_10m_only_band1(self):
        """TSB $5m → band2 = 0; tax uses band1 only = (5-3)/5 × earnings × 15%."""
        m = Member(tsb=5_000_000, split_pct=1.0)
        earnings = 100_000
        v = div296_tax_for_member(
            earnings, m,
            A.threshold_1, A.threshold_2, A.rate_tier1, A.rate_tier2,
        )
        # Manual check: (5m-3m)/5m × 100k × 15% = 40% × 100k × 15% = 6,000
        assert _r(v) == 6_000

    def test_tsb_above_10m_splits_across_bands(self):
        """TSB $15m: band1 = (10-3)/15 = 7/15; band2 = (15-10)/15 = 5/15."""
        m = Member(tsb=15_000_000, split_pct=1.0)
        earnings = 100_000
        v = div296_tax_for_member(
            earnings, m,
            A.threshold_1, A.threshold_2, A.rate_tier1, A.rate_tier2,
        )
        expected = (
            100_000 * (7_000_000 / 15_000_000) * 0.15
            + 100_000 * (5_000_000 / 15_000_000) * 0.25
        )
        assert _r(v) == _r(expected)

    def test_tsb_at_threshold_1_zero_tax(self):
        """TSB = $3m exactly: band1 = 0, band2 = 0 → tax = 0."""
        m = Member(tsb=3_000_000, split_pct=1.0)
        v = div296_tax_for_member(
            100_000, m,
            A.threshold_1, A.threshold_2, A.rate_tier1, A.rate_tier2,
        )
        assert _r(v) == 0

    def test_zero_tsb_returns_zero_tax(self):
        m = Member(tsb=0, split_pct=1.0)
        v = div296_tax_for_member(
            100_000, m,
            A.threshold_1, A.threshold_2, A.rate_tier1, A.rate_tier2,
        )
        assert v == 0.0


# --- v3.0 API-break sanity ------------------------------------------------

def test_member_no_proportion_override_field():
    """v3.0 removed `Member.proportion_override`. Constructing with it must fail."""
    with pytest.raises(TypeError):
        Member(tsb=1_000_000, split_pct=1.0, proportion_override=0.5)  # type: ignore[call-arg]


# --- v3.1 capital-loss netting (intra-year, fund-level) -------------------

class TestLossNettingV31:
    """v3.1 intra-year netting for both ordinary CGT and Div 296.

    See calcs.py module docstring "v3.1 capital-loss netting" decisions.
    """

    def test_user_scenario_gain_100k_loss_200k(self):
        """User's worked example: $100k gain (held>12m) + $200k loss.

        Expected:
        - Div 296 fund earnings = MAX(0, $66,667 + -$200,000) = $0
        - Div 296 headline tax  = $0
        - Fund Ordinary CGT     = $0 (gross gain $100k < gross loss $200k)
        - Carry-forward (fund)  = MAX(0, $200k - $100k) = $100,000
        """
        gain = Asset(
            code="G", name="Gain", quantity=1,
            original_cost_base=0, current_market_value=100_000,
            market_value_30jun2026=100_000, valuation_source="",
            projected_sale_proceeds=100_000, held_over_12_months=True,
        )
        loss = Asset(
            code="L", name="Loss", quantity=1,
            original_cost_base=200_000, current_market_value=0,
            market_value_30jun2026=0, valuation_source="",
            projected_sale_proceeds=0, held_over_12_months=True,
        )
        register = [gain, loss]
        member = Member(tsb=5_000_000, split_pct=1.0)

        assert _r(div296_fund_earnings(register, False, A.discount_rate)) == 0
        assert _r(div296_headline_tax(
            register, [member], False, A.discount_rate,
            A.threshold_1, A.threshold_2, A.rate_tier1, A.rate_tier2,
        )) == 0
        assert _r(ordinary_cgt_fund(register, A.discount_rate, A.fund_cgt_rate)) == 0
        assert _r(carry_forward_loss_fund(register)) == 100_000

    def test_all_discount_net_positive(self):
        """All-held-over-12m: $300k gain absorbs $100k loss.

        Expected ordinary_cgt_fund:
        - discount_gains=$300k, nondiscount_gains=$0, gross_losses=$100k
        - losses to non-discount first: nd_after=$0, losses_remaining=$100k
        - d_after = $300k - $100k = $200k
        - net_taxable = $0 + $200k × 2/3 = $133,333
        - CGT = $133,333 × 15% = $20,000
        """
        gain = Asset(
            code="G", name="Gain", quantity=1,
            original_cost_base=0, current_market_value=300_000,
            market_value_30jun2026=300_000, valuation_source="",
            projected_sale_proceeds=300_000, held_over_12_months=True,
        )
        loss = Asset(
            code="L", name="Loss", quantity=1,
            original_cost_base=100_000, current_market_value=0,
            market_value_30jun2026=0, valuation_source="",
            projected_sale_proceeds=0, held_over_12_months=True,
        )
        assert _r(ordinary_cgt_fund(
            [gain, loss], A.discount_rate, A.fund_cgt_rate
        )) == 20_000

    def test_loss_priority_nondiscount_first(self):
        """Mixed holding: $100k non-discount gain + $100k discount gain + $80k loss.

        Losses apply to non-discount gains first (taxpayer-favourable):
        - nondiscount_gains=$100k, gross_losses=$80k
        - nd_after = $100k - $80k = $20k
        - losses_remaining = $0
        - d_after = $100k - $0 = $100k
        - net_taxable = $20k + $100k × 2/3 = $86,667
        - CGT = $86,667 × 15% = $13,000
        """
        nd_gain = Asset(
            code="ND", name="Short-held gain", quantity=1,
            original_cost_base=0, current_market_value=100_000,
            market_value_30jun2026=100_000, valuation_source="",
            projected_sale_proceeds=100_000, held_over_12_months=False,
        )
        d_gain = Asset(
            code="D", name="Long-held gain", quantity=1,
            original_cost_base=0, current_market_value=100_000,
            market_value_30jun2026=100_000, valuation_source="",
            projected_sale_proceeds=100_000, held_over_12_months=True,
        )
        loss = Asset(
            code="L", name="Loss", quantity=1,
            original_cost_base=80_000, current_market_value=0,
            market_value_30jun2026=0, valuation_source="",
            projected_sale_proceeds=0, held_over_12_months=True,
        )
        assert _r(ordinary_cgt_fund(
            [nd_gain, d_gain, loss], A.discount_rate, A.fund_cgt_rate
        )) == 13_000

    def test_div296_earnings_floored_at_zero(self):
        """Pure regression: net negative adjusted gains → 0, not negative."""
        loss_only = Asset(
            code="L", name="Loss", quantity=1,
            original_cost_base=100_000, current_market_value=0,
            market_value_30jun2026=0, valuation_source="",
            projected_sale_proceeds=0, held_over_12_months=True,
        )
        assert div296_fund_earnings(
            [loss_only], False, A.discount_rate
        ) == 0.0

    def test_carry_forward_loss_fund_when_losses_exceed_gains(self):
        """Fund-level carry-forward when losses exceed gains."""
        gain = Asset(
            code="G", name="Gain", quantity=1,
            original_cost_base=0, current_market_value=50_000,
            market_value_30jun2026=50_000, valuation_source="",
            projected_sale_proceeds=50_000, held_over_12_months=True,
        )
        loss = Asset(
            code="L", name="Loss", quantity=1,
            original_cost_base=200_000, current_market_value=0,
            market_value_30jun2026=0, valuation_source="",
            projected_sale_proceeds=0, held_over_12_months=True,
        )
        assert _r(carry_forward_loss_fund([gain, loss])) == 150_000

    def test_carry_forward_loss_fund_zero_when_gains_dominate(self):
        """No carry-forward when gross gains exceed gross losses."""
        gain = Asset(
            code="G", name="Gain", quantity=1,
            original_cost_base=0, current_market_value=500_000,
            market_value_30jun2026=500_000, valuation_source="",
            projected_sale_proceeds=500_000, held_over_12_months=True,
        )
        loss = Asset(
            code="L", name="Loss", quantity=1,
            original_cost_base=100_000, current_market_value=0,
            market_value_30jun2026=0, valuation_source="",
            projected_sale_proceeds=0, held_over_12_months=True,
        )
        assert carry_forward_loss_fund([gain, loss]) == 0.0
