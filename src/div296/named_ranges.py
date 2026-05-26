"""Named-range registry.

Centralises every Excel defined name used by formulas. Keeps formula
text readable (e.g. `=rate_tier1` rather than `=Inputs!$C$42`) and
avoids fragile cell-address coupling.

Populated in v1.0 when Inputs/Analyser are built.

v3.0: control-panel toggles (`reset_on`, `tier10_on`, `discount_on`) are
removed. Reset is always elected for the per-asset detail view; both
scenarios remain computable in the side-by-side fund summary. Tier 2
($10m / +25%) is always applied. CGT discount applies iff `held_over_12mo`.
"""

# Assumption constants (spec §4 Zone 3)
RATE_TIER1 = "rate_tier1"
RATE_TIER2 = "rate_tier2"
THRESHOLD_1 = "threshold_1"
THRESHOLD_2 = "threshold_2"
DISCOUNT_RATE = "discount_rate"
FUND_CGT_RATE = "fund_cgt_rate"
INDEXATION_INCR_1 = "indexation_increment_1"
INDEXATION_INCR_2 = "indexation_increment_2"

ALL_NAMES = (
    RATE_TIER1, RATE_TIER2, THRESHOLD_1, THRESHOLD_2,
    DISCOUNT_RATE, FUND_CGT_RATE,
    INDEXATION_INCR_1, INDEXATION_INCR_2,
)
