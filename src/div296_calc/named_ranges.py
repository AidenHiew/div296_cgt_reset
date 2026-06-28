"""Defined-name registry for the ongoing calculator.

Rates + discount mirror the frozen tool's names. Threshold cells are the
RESOLVED per-year values (looked up from the year table), not raw constants.
"""

RATE_TIER1 = "rate_tier1"
RATE_TIER2 = "rate_tier2"
DISCOUNT_RATE = "discount_rate"
T1_SEL = "t1_sel"               # resolved threshold_1 for the selected year
T2_SEL = "t2_sel"               # resolved threshold_2 for the selected year
GREATER_OF_SEL = "greater_of_sel"   # resolved use_greater_of (1/0)

ALL_NAMES = (RATE_TIER1, RATE_TIER2, DISCOUNT_RATE, T1_SEL, T2_SEL, GREATER_OF_SEL)
