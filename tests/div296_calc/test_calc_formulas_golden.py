"""Golden-string tests pinning the exact arithmetic of generated formulas.

If one fails after an INTENTIONAL formula change, re-verify the new string
against div296_calc.calcs (the pure mirror) and the spec §5/§7, then update
the golden in the same commit.
"""

from div296_calc import _formulas as F


def test_pooled_total_formula_uses_sum_not_plus():
    assert F.pooled_total_formula("B6", "B7", "B8", "B9", "B10", "B11") == (
        "=SUM(B6,B7,B8,B9,B10)-B11"
    )


def test_cgt_net_formula_golden():
    assert F.cgt_net_formula("B15", "B16", "B17") == (
        "=MAX(0,B16-B17)+MAX(0,B15-MAX(0,B17-B16))*(1-discount_rate)"
    )


def test_cgt_unused_loss_formula_golden():
    assert F.cgt_unused_loss_formula("B15", "B16", "B17") == (
        "=MAX(0,MAX(0,B17-B16)-B15)"
    )


def test_tsb_ref_formula_golden():
    assert F.tsb_ref_formula("B23", "B24", "greater_of_sel") == (
        '=IF(B24="","",IF(greater_of_sel=1,MAX(B23,B24),B24))'
    )


def test_earnings_formula_override_then_pooled_with_name_guard():
    assert F.earnings_formula("B22", "B27", "B25", "$B$12") == (
        '=IF(B22="","",IF(B27<>"",B27,IF(OR(B25="",$B$12=""),"",B25*$B$12)))'
    )


def test_net_earnings_formula_golden():
    assert F.net_earnings_formula("B28", "B26") == (
        '=IF(B28="","",B28-IF(B26="",0,B26))'
    )


def test_band1_formula_guard_first():
    assert F.band1_formula("B29", "t1_sel", "t2_sel") == (
        '=IF(OR(B29="",B29<=0),"",MAX(0,MIN(B29,t2_sel)-t1_sel)/B29)'
    )


def test_band2_formula_only_references_t2():
    assert F.band2_formula("B29", "t2_sel") == (
        '=IF(OR(B29="",B29<=0),"",MAX(0,B29-t2_sel)/B29)'
    )


def test_tier1_tax_formula_golden():
    assert F.tier1_tax_formula("B30", "B31", "B29", "t1_sel") == (
        '=IF(OR(B30="",B30<=0,B29<=t1_sel,B31=""),0,B30*B31*rate_tier1)'
    )


def test_tier2_tax_formula_golden():
    assert F.tier2_tax_formula("B30", "B32", "B29", "t1_sel") == (
        '=IF(OR(B30="",B30<=0,B29<=t1_sel,B32=""),0,B30*B32*rate_tier2)'
    )


def test_new_loss_formula_golden():
    assert F.new_loss_formula("B30") == '=IF(B30="","",MAX(0,-B30))'


def test_status_formula_golden():
    assert F.status_formula("B29", "t1_sel") == (
        '=IF(B29="","",IF(B29<=t1_sel,"Below $3M — not liable","Liable"))'
    )


def test_threshold_lookup_uses_sumifs_and_fails_safe_on_unknown_year():
    # SUMIFS (never INDEX/MATCH — recalc-gate false positives), wrapped in a
    # COUNTIF=0 -> NA() guard so an unknown income year errors out instead of
    # resolving the threshold to a silent 0 (which would emit fictitious tax).
    assert F.threshold_lookup_formula("$P$2:$P$10", "$O$2:$O$10", "$B$3") == (
        "=IF(COUNTIF($O$2:$O$10,$B$3)=0,NA(),"
        "SUMIFS($P$2:$P$10,$O$2:$O$10,$B$3))"
    )


def test_year_known_guard_formula_actionable():
    out = F.year_known_guard_formula("$O$2:$O$10", "$B$3")
    assert out.startswith('=IF(COUNTIF($O$2:$O$10,$B$3)=0,')
    assert "add" in out.lower()


def test_pooled_share_contrib_formula_golden():
    assert F.pooled_share_contrib_formula("B22", "B27", "B25") == (
        '=IF(AND(B22<>"",B27=""),IF(B25="",0,B25),0)'
    )


def test_pooled_count_formula_uses_boolean_sumproduct():
    assert F.pooled_count_formula("B22:E22", "B27:E27") == (
        '=SUMPRODUCT((B22:E22<>"")*(B27:E27=""))'
    )


def test_share_guard_status_formula_handles_all_segregated():
    out = F.share_guard_status_formula("B38", "B39")
    assert "all members segregated" in out
    assert "100%" in out
