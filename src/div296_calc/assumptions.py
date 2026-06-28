"""Rates, member cap, and the year→thresholds table.

The table is SEEDED WITH ONLY CONFIRMED YEARS. At v0.1 that is just
2026-27 ($3M / $10M, un-indexed). Later years' thresholds are CPI-indexed
(in $150k / $500k steps) and not yet published — add them by hand to
YEAR_TABLE as the ATO announces them. Looking up an absent year raises
UnknownYearError with an actionable message (no silent default).

rate_tier1 / rate_tier2 deliberately mirror the frozen div296.assumptions
so the two tools never drift; a test pins the equality.
"""

from __future__ import annotations

from dataclasses import dataclass

# Enacted Division 296 additional rates (slice form: 15% on $3M–$10M,
# 25% = 15%+extra 10% on the slice above $10M). See the spec §5 footgun note.
RATE_TIER1: float = 0.15
RATE_TIER2: float = 0.25

# SMSF 1/3 CGT discount — used by the CGT netting helper.
DISCOUNT_RATE: float = 1.0 / 3.0

# Up to 4 members per SMSF.
MEMBER_COUNT: int = 4


@dataclass(frozen=True)
class YearThresholds:
    threshold_1: int          # lower-super-balance threshold (LSBT), e.g. $3M
    threshold_2: int          # very-large-super-balance threshold (VLSBT), e.g. $10M
    use_greater_of: bool      # False for 2026-27 (closing TSB); True from 2027-28


# Only confirmed years. ADD A ROW per ATO-published indexed thresholds.
YEAR_TABLE: dict[str, YearThresholds] = {
    "2026-27": YearThresholds(3_000_000, 10_000_000, use_greater_of=False),
}


class UnknownYearError(KeyError):
    """Raised when an income year is not in YEAR_TABLE."""


def thresholds_for(year: str) -> YearThresholds:
    """Return the thresholds row for `year`, or raise UnknownYearError."""
    try:
        return YEAR_TABLE[year]
    except KeyError:
        raise UnknownYearError(
            f"No threshold row for income year {year!r}. Add the ATO-published "
            f"indexed $3M/$10M thresholds for {year} to "
            f"div296_calc.assumptions.YEAR_TABLE before running."
        ) from None
