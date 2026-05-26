"""Pure-Python mirror of every Excel formula.

The Excel workbook is the deliverable; this module exists so the test
suite can verify the spec §12 acceptance numbers without Excel/LibreOffice.
Both the workbook formulas and this module must produce identical results
for any input.

Locked decisions (see README.md):
- Ordinary CGT is per-asset siloed (no in-year loss offset).
- Div 296 fund earnings = sum of POSITIVE adjusted gains (asset-level floor).
- Per-asset Div 296 tax = pro-rata share of the headline member-attributed total.
- Pension phase is NOT modelled (assumes 100% accumulation, 15% fund CGT rate).

v3.0 simplification (breaking API change vs v2.x):
- `tier10_on` parameter removed — Tier 2 ($10m / +25%) is enacted law and
  always applied; the v2.x toggle had no Bill-correct use case.
- `discount_on` parameter removed — the 1/3 CGT discount applies iff the asset
  was held > 12 months; the per-asset `held_over_12_months` field is the
  legally meaningful control.
- `Member.proportion_override` field removed — provably dead in the default
  config since v2.5 (only ever fed the now-removed tier_off branch).
- `member_proportion_above_3m` function removed — only ever called from the
  deleted tier_off branch.

`reset_on` parameter is retained — it is a scenario selector ("which cost
base do I use for this calc?"), not a global toggle. Both the "if no reset"
and "if elected to reset" scenarios remain computable for the Comparison tab
side-by-side and Analyser fund summary side-by-side.
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import Sequence


@dataclass
class Asset:
    code: str
    name: str
    quantity: float
    original_cost_base: float
    current_market_value: float
    market_value_30jun2026: float
    valuation_source: str
    projected_sale_proceeds: float
    held_over_12_months: bool


@dataclass
class Member:
    tsb: float
    split_pct: float                          # 0.0–1.0


# ---- per-asset gain calculations (spec §5 col 5 and col 7) ----

def _apply_discount(raw_gain: float, asset: Asset, discount_rate: float) -> float:
    """Discount applies to gains only, never to losses.

    1/3 super-fund CGT discount applies iff the asset was held > 12 months
    (per s115-25 ITAA 1997). The v2.x `discount_on` global toggle is removed
    in v3.0 — it had no Bill-correct use case.
    """
    if raw_gain <= 0:
        return raw_gain
    if asset.held_over_12_months:
        return raw_gain * (1 - discount_rate)
    return raw_gain


def ordinary_raw_gain(asset: Asset) -> float:
    return asset.projected_sale_proceeds - asset.original_cost_base


def ordinary_taxable_gain(asset: Asset, discount_rate: float) -> float:
    return _apply_discount(ordinary_raw_gain(asset), asset, discount_rate)


def div296_cost_base(asset: Asset, reset_on: bool) -> float:
    """Scenario-dependent cost base. `reset_on=True` → MV at 30 Jun 2026;
    `reset_on=False` → original cost base. Both scenarios are needed for the
    side-by-side comparison."""
    return asset.market_value_30jun2026 if reset_on else asset.original_cost_base


def div296_raw_gain(asset: Asset, reset_on: bool) -> float:
    return asset.projected_sale_proceeds - div296_cost_base(asset, reset_on)


def div296_adjusted_gain(
    asset: Asset, reset_on: bool, discount_rate: float
) -> float:
    return _apply_discount(div296_raw_gain(asset, reset_on), asset, discount_rate)


# ---- per-asset ordinary CGT (spec §5 col 6) ----

def ordinary_cgt(
    asset: Asset, discount_rate: float, fund_cgt_rate: float
) -> float:
    """Per-asset silo: max(0, taxable gain) × fund CGT rate. Losses contribute $0."""
    return max(0.0, ordinary_taxable_gain(asset, discount_rate)) * fund_cgt_rate


def carry_forward_loss(asset: Asset) -> float:
    """Magnitude of any raw ordinary capital loss (positive number)."""
    return max(0.0, -ordinary_raw_gain(asset))


# ---- Div 296 fund earnings (locked: sum of positive adjusted gains) ----

def div296_fund_earnings(
    assets: Sequence[Asset], reset_on: bool, discount_rate: float
) -> float:
    return sum(
        max(0.0, div296_adjusted_gain(a, reset_on, discount_rate))
        for a in assets
    )


# ---- per-member Div 296 tax (spec §7) ----

def div296_tax_for_member(
    earnings: float,
    member: Member,
    threshold_1: float,
    threshold_2: float,
    rate_tier1: float,
    rate_tier2: float,
) -> float:
    """Tax on the member's share of earnings using two-band proportion.

    v3.0: always two-band. The v2.x `tier10_on` toggle and single-tier
    fallback are removed — the $10m / +25% tier is enacted law and the
    single-tier path was never Bill-correct.
    """
    e = earnings * member.split_pct
    if e <= 0 or member.tsb <= 0:
        return 0.0
    band1 = max(0.0, min(member.tsb, threshold_2) - threshold_1) / member.tsb
    band2 = max(0.0, member.tsb - threshold_2) / member.tsb
    return e * band1 * rate_tier1 + e * band2 * rate_tier2


def div296_headline_tax(
    assets: Sequence[Asset],
    members: Sequence[Member],
    reset_on: bool,
    discount_rate: float,
    threshold_1: float,
    threshold_2: float,
    rate_tier1: float,
    rate_tier2: float,
) -> float:
    """Sum of per-member Div 296 tax."""
    earnings = div296_fund_earnings(assets, reset_on, discount_rate)
    return sum(
        div296_tax_for_member(
            earnings, m, threshold_1, threshold_2, rate_tier1, rate_tier2
        )
        for m in members
    )


# ---- per-asset Div 296 tax — pro-rata of headline (locked decision) ----

def per_asset_div296_tax(
    asset: Asset,
    all_assets: Sequence[Asset],
    headline_tax: float,
    reset_on: bool,
    discount_rate: float,
) -> float:
    """Each asset's share of the headline = its positive adjusted gain ÷ sum of positives."""
    my_gain = max(0.0, div296_adjusted_gain(asset, reset_on, discount_rate))
    total = sum(
        max(0.0, div296_adjusted_gain(a, reset_on, discount_rate))
        for a in all_assets
    )
    if total <= 0:
        return 0.0
    return my_gain / total * headline_tax
