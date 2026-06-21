"""Tests for div296_calc.assumptions — year table + frozen-law parity."""

import pytest

from div296_calc.assumptions import (
    MEMBER_COUNT,
    RATE_TIER1,
    RATE_TIER2,
    UnknownYearError,
    YearThresholds,
    thresholds_for,
)


def test_package_version_present():
    import div296_calc
    assert div296_calc.__version__ == "0.1.0"


def test_2026_27_thresholds():
    yt = thresholds_for("2026-27")
    assert yt == YearThresholds(threshold_1=3_000_000, threshold_2=10_000_000,
                                use_greater_of=False)


def test_unknown_year_raises_with_actionable_message():
    with pytest.raises(UnknownYearError) as exc:
        thresholds_for("2027-28")
    assert "2027-28" in str(exc.value)
    assert "add" in str(exc.value).lower()


def test_rates_match_frozen_div296_assumptions():
    # The ongoing tool must not drift from the shipped CGT-reset tool's rates.
    from div296.assumptions import ASSUMPTIONS
    assert RATE_TIER1 == ASSUMPTIONS.rate_tier1 == 0.15
    assert RATE_TIER2 == ASSUMPTIONS.rate_tier2 == 0.25


def test_2026_27_thresholds_match_frozen_constants():
    from div296.assumptions import ASSUMPTIONS
    yt = thresholds_for("2026-27")
    assert yt.threshold_1 == ASSUMPTIONS.threshold_1
    assert yt.threshold_2 == ASSUMPTIONS.threshold_2


def test_member_cap_is_four():
    assert MEMBER_COUNT == 4
