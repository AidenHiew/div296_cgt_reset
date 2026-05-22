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
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import Optional, Sequence


@dataclass
class Asset:
    code: str
    name: str
    quantity: float
    original_cost_base: float
    total_value: float
    market_value_30jun2026: float
    valuation_source: str
    projected_sale_proceeds: float
    held_over_12_months: bool


@dataclass
class Member:
    tsb: float
    split_pct: float                          # 0.0–1.0
    proportion_override: Optional[float] = None


# ---- per-asset gain calculations (spec §5 col 5 and col 7) ----

def _apply_discount(raw_gain: float, asset: Asset, discount_on: bool, discount_rate: float) -> float:
    """Discount applies to gains only, never to losses."""
    if raw_gain <= 0:
        return raw_gain
    if discount_on and asset.held_over_12_months:
        return raw_gain * (1 - discount_rate)
    return raw_gain


def ordinary_raw_gain(asset: Asset) -> float:
    return asset.projected_sale_proceeds - asset.original_cost_base


def ordinary_taxable_gain(asset: Asset, discount_on: bool, discount_rate: float) -> float:
    return _apply_discount(ordinary_raw_gain(asset), asset, discount_on, discount_rate)


def div296_cost_base(asset: Asset, reset_on: bool) -> float:
    return asset.market_value_30jun2026 if reset_on else asset.original_cost_base


def div296_raw_gain(asset: Asset, reset_on: bool) -> float:
    return asset.projected_sale_proceeds - div296_cost_base(asset, reset_on)


def div296_adjusted_gain(
    asset: Asset, reset_on: bool, discount_on: bool, discount_rate: float
) -> float:
    return _apply_discount(div296_raw_gain(asset, reset_on), asset, discount_on, discount_rate)


# ---- per-asset ordinary CGT (spec §5 col 6) ----

def ordinary_cgt(
    asset: Asset, discount_on: bool, discount_rate: float, fund_cgt_rate: float
) -> float:
    """Per-asset silo: max(0, taxable gain) × fund CGT rate. Losses contribute $0."""
    return max(0.0, ordinary_taxable_gain(asset, discount_on, discount_rate)) * fund_cgt_rate


def carry_forward_loss(asset: Asset) -> float:
    """Magnitude of any raw ordinary capital loss (positive number)."""
    return max(0.0, -ordinary_raw_gain(asset))


# ---- member proportion above $3m (spec §4 Zone 3) ----

def member_proportion_above_3m(
    member: Member, threshold_1: float
) -> float:
    if member.proportion_override is not None:
        return member.proportion_override
    if member.tsb <= 0:
        return 0.0
    return max(0.0, (member.tsb - threshold_1) / member.tsb)


# ---- Div 296 fund earnings (locked: sum of positive adjusted gains) ----

def div296_fund_earnings(
    assets: Sequence[Asset], reset_on: bool, discount_on: bool, discount_rate: float
) -> float:
    return sum(
        max(0.0, div296_adjusted_gain(a, reset_on, discount_on, discount_rate))
        for a in assets
    )


# ---- per-member Div 296 tax (spec §7) ----

def div296_tax_for_member(
    earnings: float,
    member: Member,
    tier10_on: bool,
    threshold_1: float,
    threshold_2: float,
    rate_tier1: float,
    rate_tier2: float,
) -> float:
    """Tax on the member's share of earnings using one- or two-tier proportion."""
    e = earnings * member.split_pct
    if e <= 0 or member.tsb <= 0:
        return 0.0
    if not tier10_on:
        p = member_proportion_above_3m(member, threshold_1)
        return e * p * rate_tier1
    # Tier on: split earnings across two bands by balance proportion.
    band1 = max(0.0, min(member.tsb, threshold_2) - threshold_1) / member.tsb
    band2 = max(0.0, member.tsb - threshold_2) / member.tsb
    return e * band1 * rate_tier1 + e * band2 * rate_tier2


def div296_headline_tax(
    assets: Sequence[Asset],
    members: Sequence[Member],
    reset_on: bool,
    discount_on: bool,
    discount_rate: float,
    tier10_on: bool,
    threshold_1: float,
    threshold_2: float,
    rate_tier1: float,
    rate_tier2: float,
    earnings_source: str = "Auto",
    manual_earnings: float = 0.0,
) -> float:
    """Sum of per-member Div 296 tax. Honours the Auto/Manual earnings switch."""
    if earnings_source == "Manual":
        earnings = manual_earnings
    else:
        earnings = div296_fund_earnings(assets, reset_on, discount_on, discount_rate)
    return sum(
        div296_tax_for_member(
            earnings, m, tier10_on, threshold_1, threshold_2, rate_tier1, rate_tier2
        )
        for m in members
    )


# ---- per-asset Div 296 tax — pro-rata of headline (locked decision) ----

def per_asset_div296_tax(
    asset: Asset,
    all_assets: Sequence[Asset],
    headline_tax: float,
    reset_on: bool,
    discount_on: bool,
    discount_rate: float,
) -> float:
    """Each asset's share of the headline = its positive adjusted gain ÷ sum of positives."""
    my_gain = max(0.0, div296_adjusted_gain(asset, reset_on, discount_on, discount_rate))
    total = sum(
        max(0.0, div296_adjusted_gain(a, reset_on, discount_on, discount_rate))
        for a in all_assets
    )
    if total <= 0:
        return 0.0
    return my_gain / total * headline_tax
