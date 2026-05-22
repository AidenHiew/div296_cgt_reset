"""Single source of truth for rates, thresholds, and labels.

Every number that appears in an Excel formula must originate here and be
written to a named cell on the Inputs tab. Formulas reference cells, never
literals. The same constants drive the Python test mirror in calcs.py.

Refer to docs/BUILD_PLAN.md §4 (Zone 3) for the spec.
"""

from dataclasses import dataclass


@dataclass(frozen=True)
class Assumptions:
    # Div 296 additional rates
    rate_tier1: float = 0.15          # $3m–$10m band
    rate_tier2: float = 0.25          # above $10m

    # Thresholds (un-indexed for 2026–27)
    threshold_1: int = 3_000_000
    threshold_2: int = 10_000_000

    # Indexation increments — informational only, not auto-applied
    indexation_increment_1: int = 150_000
    indexation_increment_2: int = 500_000

    # SMSF CGT discount (1/3 = 33.333%, i.e. taxable = gain × 2/3)
    discount_rate: float = 1.0 / 3.0

    # Standard SMSF accumulation-phase earnings tax rate applied to
    # ordinary capital gains. Pension phase is NOT modelled (see Notes).
    fund_cgt_rate: float = 0.15

    # Asset register capacity (rows on Inputs)
    asset_register_rows: int = 50

    # Members (single member is the default; rows 2–4 stay visible per spec)
    member_count: int = 4


ASSUMPTIONS = Assumptions()
