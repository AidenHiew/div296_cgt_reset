"""Golden-string tests pinning the exact arithmetic of generated formulas.

The rest of the fast suite is structural (does the formula reference the
right cells?) and cannot catch a sign flip. These pin full formula text for
the highest-blast-radius builders. If one fails after an INTENTIONAL formula
change: re-verify the new string against the sign conventions in CONTEXT.md,
then update the golden string in the same commit as the change.
"""
import pytest

from div296._formulas import per_member_div296_tax_formula
from div296.build import build_workbook
from div296.tabs import comparison as C_TAB


@pytest.fixture(scope="module")
def wb():
    return build_workbook()


def test_per_member_tax_formula_golden():
    assert per_member_div296_tax_formula(7, "$L$6") == (
        "=IF(OR('Inputs'!B7=\"\",'Inputs'!C7=\"\",'Inputs'!B7<=0,"
        "'Inputs'!C7<=0,$L$6<=0),0,"
        "$L$6*'Inputs'!C7*'Inputs'!D7*rate_tier1"
        " + $L$6*'Inputs'!C7*'Inputs'!E7*rate_tier2)"
    )


def test_analyser_div296_postdisc_gain_row17_golden(wb):
    # Col J, first per-asset row: elected-reset Div 296 gain (post-discount).
    assert wb["Analyser"]["J17"].value == (
        "=IF('Inputs'!G16=\"\",\"\","
        "IF(('Inputs'!G16-'Inputs'!E16)<=0,('Inputs'!G16-'Inputs'!E16),"
        "IF('Inputs'!J16=\"Yes\","
        "('Inputs'!G16-'Inputs'!E16)*(1-discount_rate),"
        "('Inputs'!G16-'Inputs'!E16))))"
    )


def test_analyser_ord_gross_gain_row17_golden(wb):
    # Col E, first per-asset row: ordinary gross gain (proceeds − orig CB).
    assert wb["Analyser"]["E17"].value == (
        "=IF('Inputs'!G16=\"\",\"\","
        "'Inputs'!G16-'Inputs'!C16)"
    )


def test_comparison_noreset_gain_row16_golden(wb):
    cell = f"{C_TAB.PER_REG_GAIN_A_COL}16"
    assert wb["Comparison"][cell].value == (
        "=IF('Inputs'!G16=\"\",\"\","
        "IF(('Inputs'!G16-'Inputs'!C16)<=0,('Inputs'!G16-'Inputs'!C16),"
        "IF('Inputs'!J16=\"Yes\","
        "('Inputs'!G16-'Inputs'!C16)*(1-discount_rate),"
        "('Inputs'!G16-'Inputs'!C16))))"
    )
