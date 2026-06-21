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
