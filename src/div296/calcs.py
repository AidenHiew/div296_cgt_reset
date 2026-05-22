"""Pure-Python mirror of every Excel formula.

The Excel workbook is the deliverable; this module exists so the test
suite can verify the spec §12 acceptance numbers without spinning up
Excel/LibreOffice. Both the workbook formulas and this module must
produce identical results for any input.

Populated in v1.0. Stubs only in v0.1.
"""

from dataclasses import dataclass
from typing import Optional


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
    split_pct: float
    proportion_override: Optional[float] = None


def ordinary_taxable_gain(asset: Asset, discount_on: bool, discount_rate: float) -> float:
    raise NotImplementedError("v1.0")


def div296_adjusted_gain(
    asset: Asset, reset_on: bool, discount_on: bool, discount_rate: float
) -> float:
    raise NotImplementedError("v1.0")


def member_proportion_above_3m(member: Member, threshold_1: float) -> float:
    raise NotImplementedError("v1.0")


def div296_tax_for_member(
    earnings: float,
    member: Member,
    tier10_on: bool,
    threshold_1: float,
    threshold_2: float,
    rate_tier1: float,
    rate_tier2: float,
) -> float:
    raise NotImplementedError("v1.0")
