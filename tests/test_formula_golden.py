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
    # v3.4 F2: blanks when EITHER proceeds or MV (col E) is blank.
    assert wb["Analyser"]["J17"].value == (
        "=IF(OR('Inputs'!G16=\"\",'Inputs'!E16=\"\"),\"\","
        "IF(('Inputs'!G16-'Inputs'!E16)<=0,('Inputs'!G16-'Inputs'!E16),"
        "IF('Inputs'!J16=\"Yes\","
        "('Inputs'!G16-'Inputs'!E16)*(1-discount_rate),"
        "('Inputs'!G16-'Inputs'!E16))))"
    )


def test_analyser_ord_gross_gain_row17_golden(wb):
    # Col E, first per-asset row: ordinary gross gain (proceeds − orig CB).
    # v3.4 F3: blanks when EITHER proceeds or orig CB is blank.
    assert wb["Analyser"]["E17"].value == (
        "=IF(OR('Inputs'!G16=\"\",'Inputs'!C16=\"\"),\"\","
        "'Inputs'!G16-'Inputs'!C16)"
    )


def test_comparison_noreset_gain_row16_golden(wb):
    cell = f"{C_TAB.PER_REG_GAIN_A_COL}16"
    # v3.4 F3: blanks when EITHER proceeds or orig CB is blank.
    assert wb["Comparison"][cell].value == (
        "=IF(OR('Inputs'!G16=\"\",'Inputs'!C16=\"\"),\"\","
        "IF(('Inputs'!G16-'Inputs'!C16)<=0,('Inputs'!G16-'Inputs'!C16),"
        "IF('Inputs'!J16=\"Yes\","
        "('Inputs'!G16-'Inputs'!C16)*(1-discount_rate),"
        "('Inputs'!G16-'Inputs'!C16))))"
    )


# --- s102-5 fund Ordinary CGT recon chain (Analyser O70/P70/Q70 -> B71/B73) ---
# These feed the headline Fund Ordinary CGT and the TOTAL TAX BURDEN, yet are
# excluded from the recalc gate (they false-positive in the pure-Python engine),
# so a sign/criteria flip here would otherwise ship green (v3.4 review).

def test_analyser_recon_disc_gains_helper_golden(wb):
    # O70: discountable gains = positive projected gains that are long-held
    # (Inputs!J = "Yes"). A "<>Yes" here would swap disc/non-disc baskets.
    assert wb["Analyser"]["O70"].value == (
        "=SUMIFS('Inputs'!H16:H65, 'Inputs'!H16:H65, \">0\", "
        "'Inputs'!J16:J65, \"Yes\")"
    )


def test_analyser_recon_nondisc_gains_helper_golden(wb):
    # P70: non-discountable gains = positive projected gains NOT long-held.
    assert wb["Analyser"]["P70"].value == (
        "=SUMIFS('Inputs'!H16:H65, 'Inputs'!H16:H65, \">0\", "
        "'Inputs'!J16:J65, \"<>Yes\")"
    )


def test_analyser_recon_gross_losses_helper_golden(wb):
    # Q70: gross losses as a POSITIVE number (negate the SUMIF of negatives).
    # A ">0" criterion here would capture gains and corrupt the netting.
    assert wb["Analyser"]["Q70"].value == "=-SUMIF('Inputs'!H16:H65, \"<0\")"


def test_analyser_fund_ord_cgt_recon_golden(wb):
    # B71: s102-5 fund Ordinary CGT. Losses (Q70) absorb non-discount gains
    # (P70) FIRST; the remainder eats discount gains (O70); only the surviving
    # long-held slice is discounted. Drives the Comparison TOTAL TAX BURDEN.
    assert wb["Analyser"]["B71"].value == (
        "=(MAX(0,P70-Q70)+MAX(0,O70-MAX(0,Q70-P70))*(1-discount_rate))"
        "*fund_cgt_rate"
    )


def test_analyser_carry_forward_losses_golden(wb):
    # B73: capital losses carried forward = losses in excess of all gains.
    assert wb["Analyser"]["B73"].value == "=MAX(0, Q70 - (O70 + P70))"


def test_analyser_per_asset_ord_cgt_row17_golden(wb):
    # G17: per-asset Ordinary CGT (info only). The 1/3 discount applies ONLY
    # when the eligibility flag F17 = "Yes"; a swapped branch would discount
    # short-held assets and full-tax long-held ones (the review's mutant).
    assert wb["Analyser"]["G17"].value == (
        '=IF(E17="","",IF(E17<=0,"—",'
        'E17*(1-IF(F17="Yes",discount_rate,0))*fund_cgt_rate))'
    )


def test_inputs_band1_proportion_row7_golden(wb):
    # Inputs D7: fraction of TSB in the $3m-$10m band (threshold_1..threshold_2).
    # Drives every member's Div 296 tier-1 tax; a threshold swap mis-taxes all.
    assert wb["Inputs"]["D7"].value == (
        "=IF(B7>0, MAX(0, MIN(B7, threshold_2) - threshold_1) / B7, 0)"
    )


def test_inputs_band2_proportion_row7_golden(wb):
    # Inputs E7: fraction of TSB above $10m (threshold_2). Must NOT reference
    # threshold_1 — only the >$10m slice.
    assert wb["Inputs"]["E7"].value == (
        "=IF(B7>0, MAX(0, (B7 - threshold_2)) / B7, 0)"
    )
