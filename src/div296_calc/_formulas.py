"""Excel formula-string builders for the Calculator tab.

Discipline (spec §7): aggregation via SUM (text coerces to 0, '+' propagates
#VALUE!); guard-first IF for every division/blank-render; threshold lookup via
exact-match SUMIFS (never INDEX/MATCH/VLOOKUP — recalc-gate false positives).
Boolean-only SUMPRODUCT (no text in the multiplied ranges) is allowed.
"""

from __future__ import annotations


def pooled_total_formula(div_c, int_c, rent_c, other_c, cg_c, exp_c) -> str:
    return f"=SUM({div_c},{int_c},{rent_c},{other_c},{cg_c})-{exp_c}"


def cgt_net_formula(over_c, under_c, loss_c) -> str:
    return (f"=MAX(0,{under_c}-{loss_c})"
            f"+MAX(0,{over_c}-MAX(0,{loss_c}-{under_c}))*(1-discount_rate)")


def cgt_unused_loss_formula(over_c, under_c, loss_c) -> str:
    return f"=MAX(0,MAX(0,{loss_c}-{under_c})-{over_c})"


def tsb_ref_formula(open_c, close_c, greater_of_c) -> str:
    return (f'=IF({close_c}="","",'
            f'IF({greater_of_c}=1,MAX({open_c},{close_c}),{close_c}))')


def earnings_formula(name_c, override_c, share_c, pool_c) -> str:
    return (f'=IF({name_c}="","",'
            f'IF({override_c}<>"",{override_c},'
            f'IF(OR({share_c}="",{pool_c}=""),"",{share_c}*{pool_c})))')


def net_earnings_formula(earnings_c, prior_loss_c) -> str:
    return (f'=IF({earnings_c}="","",'
            f'{earnings_c}-IF({prior_loss_c}="",0,{prior_loss_c}))')


def band1_formula(tsb_ref_c, t1_c, t2_c) -> str:
    return (f'=IF(OR({tsb_ref_c}="",{tsb_ref_c}<=0),"",'
            f'MAX(0,MIN({tsb_ref_c},{t2_c})-{t1_c})/{tsb_ref_c})')


def band2_formula(tsb_ref_c, t2_c) -> str:
    return (f'=IF(OR({tsb_ref_c}="",{tsb_ref_c}<=0),"",'
            f'MAX(0,{tsb_ref_c}-{t2_c})/{tsb_ref_c})')


def tier1_tax_formula(net_c, band1_c, tsb_ref_c, t1_c) -> str:
    return (f'=IF(OR({net_c}="",{net_c}<=0,{tsb_ref_c}<={t1_c},{band1_c}=""),0,'
            f'{net_c}*{band1_c}*rate_tier1)')


def tier2_tax_formula(net_c, band2_c, tsb_ref_c, t1_c) -> str:
    return (f'=IF(OR({net_c}="",{net_c}<=0,{tsb_ref_c}<={t1_c},{band2_c}=""),0,'
            f'{net_c}*{band2_c}*rate_tier2)')


def new_loss_formula(net_c) -> str:
    return f'=IF({net_c}="","",MAX(0,-{net_c}))'


def status_formula(tsb_ref_c, t1_c) -> str:
    return (f'=IF({tsb_ref_c}="","",'
            f'IF({tsb_ref_c}<={t1_c},"Below $3M — not liable","Liable"))')


def threshold_lookup_formula(value_range, year_range, year_c) -> str:
    # Fail SAFE on an unknown income year: COUNTIF=0 -> NA() so the missing
    # threshold propagates #N/A through every band/tier/total cell (visibly
    # broken) instead of SUMIFS returning 0, which would collapse t2 to 0,
    # drive band2 to 1.0 and emit a fictitious >$10M-tier liability. The
    # year_known_guard banner explains it; this stops the numbers computing.
    return (f"=IF(COUNTIF({year_range},{year_c})=0,NA(),"
            f"SUMIFS({value_range},{year_range},{year_c}))")


def year_known_guard_formula(year_range, year_c) -> str:
    return (f'=IF(COUNTIF({year_range},{year_c})=0,'
            f'"⚠ Income year not in the threshold table — add the '
            f'ATO-published $3M/$10M row before relying on this.",'
            f'"✓ thresholds loaded")')


def pooled_share_contrib_formula(name_c, override_c, share_c) -> str:
    return (f'=IF(AND({name_c}<>"",{override_c}=""),'
            f'IF({share_c}="",0,{share_c}),0)')


def pooled_count_formula(name_range, override_range) -> str:
    return f'=SUMPRODUCT(({name_range}<>"")*({override_range}=""))'


def share_guard_status_formula(pooled_count_c, share_sum_c) -> str:
    return (f'=IF({pooled_count_c}=0,"n/a — all members segregated",'
            f'IF(ROUND({share_sum_c},4)=1,"✓ pooled shares = 100%",'
            f'"⚠ pooled shares ≠ 100% — check member shares"))')
