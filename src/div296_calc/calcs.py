"""Pure-Python Division 296 ongoing-year calc engine.

The Excel workbook is the deliverable; this module mirrors every formula
so the test suite can verify the numbers without Excel. NO openpyxl import
in this file (or assumptions.py) — the calc layer is openpyxl-free.

Tax kernel uses the disjoint-SLICE form: 15% on the $3M–$10M slice, 25%
(=15%+extra 10%) on the slice above $10M. This is algebraically identical
to the law's "15% above $3M plus extra 10% above $10M"; pairing the slice
band2 with 0.10 instead of 0.25 undercounts (the documented footgun).
"""

from __future__ import annotations

from collections.abc import Sequence
from dataclasses import dataclass


@dataclass(frozen=True)
class Member:
    name: str
    opening_tsb: float
    closing_tsb: float
    share: float                      # 0.0–1.0 of pooled earnings
    prior_loss: float = 0.0           # prior-year Div 296 carried-forward loss (>=0)
    override: float | None = None     # segregated earnings; None ⇒ pooled


def tsb_ref(member: Member, use_greater_of: bool) -> float:
    """The balance the proportions are measured against.

    2026-27: closing TSB. 2027-28+: greater of opening/closing (anti-avoidance).
    """
    if use_greater_of:
        return max(member.opening_tsb, member.closing_tsb)
    return member.closing_tsb


def member_div296_tax(
    net_earnings: float,
    tsb_ref_value: float,
    threshold_1: float,
    threshold_2: float,
    rate_tier1: float,
    rate_tier2: float,
) -> float:
    """Div 296 tax on a member's net Div 296 earnings.

    Returns 0 when there are no positive net earnings or the balance is at
    or below threshold_1.
    """
    if net_earnings <= 0 or tsb_ref_value <= threshold_1:
        return 0.0
    band1 = max(0.0, min(tsb_ref_value, threshold_2) - threshold_1) / tsb_ref_value
    band2 = max(0.0, tsb_ref_value - threshold_2) / tsb_ref_value
    return net_earnings * (band1 * rate_tier1 + band2 * rate_tier2)


@dataclass(frozen=True)
class CgtInputs:
    gross_gains_over_12m: float       # discountable
    gross_gains_under_12m: float      # non-discountable
    capital_losses: float             # current-year + brought-forward


@dataclass(frozen=True)
class CgtResult:
    net_realised_cg: float
    unused_capital_loss: float        # carries forward to next year's losses


def net_capital_gain(cgt: CgtInputs, discount_rate: float) -> CgtResult:
    """s102-5 method: losses hit non-discount gains first, then the 1/3
    discount applies to the surviving long-held remainder."""
    loss = cgt.capital_losses
    nondisc_net = max(0.0, cgt.gross_gains_under_12m - loss)
    loss = max(0.0, loss - cgt.gross_gains_under_12m)
    disc_net = max(0.0, cgt.gross_gains_over_12m - loss)
    net = nondisc_net + disc_net * (1.0 - discount_rate)
    unused = max(0.0, loss - cgt.gross_gains_over_12m)
    return CgtResult(net_realised_cg=net, unused_capital_loss=unused)


@dataclass(frozen=True)
class PooledIncome:
    dividends_grossed: float          # incl. franking credits, per the fund return
    interest: float
    rent: float
    other: float
    net_realised_cg: float            # from net_capital_gain()
    deductible_expenses: float


def pooled_total(p: PooledIncome) -> float:
    """Realised income less deductible expenses. May be negative."""
    return (p.dividends_grossed + p.interest + p.rent + p.other
            + p.net_realised_cg - p.deductible_expenses)


def member_earnings(member: Member, pool_total: float) -> float:
    """Segregated members earn only their override; otherwise share × pool.

    `override is None` ⇒ pooled. An override of 0.0 is honoured (not treated
    as 'no override').
    """
    if member.override is not None:
        return member.override
    return member.share * pool_total


@dataclass(frozen=True)
class MemberResult:
    name: str
    earnings: float
    tsb_ref: float
    used_greater_of: bool
    net_earnings: float
    band1: float
    band2: float
    tier1_tax: float
    tier2_tax: float
    tax: float
    new_loss: float
    below_threshold: bool


def compute_member(
    member: Member,
    pool_total: float,
    yt,                               # YearThresholds
    rate_tier1: float,
    rate_tier2: float,
) -> MemberResult:
    earnings = member_earnings(member, pool_total)
    ref = tsb_ref(member, yt.use_greater_of)
    net = earnings - member.prior_loss
    below = ref <= yt.threshold_1

    if ref > 0:
        band1 = max(0.0, min(ref, yt.threshold_2) - yt.threshold_1) / ref
        band2 = max(0.0, ref - yt.threshold_2) / ref
    else:
        band1 = band2 = 0.0

    if net <= 0 or below:
        tier1 = tier2 = 0.0
        new_loss = max(0.0, -net)
    else:
        tier1 = net * band1 * rate_tier1
        tier2 = net * band2 * rate_tier2
        new_loss = 0.0

    return MemberResult(
        name=member.name, earnings=earnings, tsb_ref=ref,
        used_greater_of=yt.use_greater_of, net_earnings=net,
        band1=band1, band2=band2, tier1_tax=tier1, tier2_tax=tier2,
        tax=tier1 + tier2, new_loss=new_loss, below_threshold=below,
    )


def compute_fund(
    members: Sequence[Member],
    pool_total: float,
    yt,
    rate_tier1: float,
    rate_tier2: float,
) -> list[MemberResult]:
    return [compute_member(m, pool_total, yt, rate_tier1, rate_tier2)
            for m in members]


def fund_total_tax(results: Sequence[MemberResult]) -> float:
    return sum(r.tax for r in results)


@dataclass(frozen=True)
class ShareStatus:
    pooled_count: int
    share_sum: float
    all_segregated: bool
    ok: bool


def pooled_share_status(members: Sequence[Member]) -> ShareStatus:
    """Soft guard on pooled members' shares. Suppressed (ok=True) when every
    member is segregated (overridden), since there is then no pool to split."""
    pooled = [m for m in members if m.override is None]
    share_sum = sum(m.share for m in pooled)
    all_seg = len(pooled) == 0
    ok = all_seg or abs(share_sum - 1.0) < 1e-9
    return ShareStatus(pooled_count=len(pooled), share_sum=share_sum,
                       all_segregated=all_seg, ok=ok)
