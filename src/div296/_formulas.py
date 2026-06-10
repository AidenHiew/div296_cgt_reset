"""Shared Excel formula builders.

Extracted in v3.0 to avoid drift between Analyser and Comparison, which
both compute per-member Div 296 tax using identical formulas. Before v3.0
each tab carried its own copy of `_member_tax_formula`; after v3.0 they
both call into here.

The post-v3.0 per-member tax formula:
- Reads band1 (Inputs col D) and band2 (Inputs col E) directly — the same
  cells the user sees on Inputs, satisfying the single-source-of-truth
  manual-validation goal.
- No `tier10_on` IF branch (v3.0 removed the toggle).
- No `proportion_override` reference (v3.0 removed the field).
- Guards zero/blank TSB, split, or earnings so blank member rows render
  as 0 rather than #DIV/0!.
"""

from __future__ import annotations

INPUTS_SHEET = "'Inputs'"


def per_member_div296_tax_formula(member_inputs_row: int, earnings_cell: str) -> str:
    """Excel formula string for a member's Div 296 tax on `earnings_cell`.

    Parameters
    ----------
    member_inputs_row
        Row number on the Inputs sheet where the member's data lives
        (TSB in col B, split% in col C, band1 in col D, band2 in col E).
    earnings_cell
        Cell reference (already qualified with sheet if cross-sheet) for the
        fund-level Div 296 earnings figure to apply this member's share to.
        Examples: `"$B$8"` (same-sheet), `"'Analyser'!$B$8"` (cross-sheet).

    Returns
    -------
    Excel formula string beginning with `=`.

    Formula shape (v3.0):
        IF(any guard fails, 0,
           earnings × split × band1 × rate_tier1
           + earnings × split × band2 × rate_tier2)

    Guard conditions return 0 (not blank or #DIV/0!) so totals and signed
    differences continue to make sense across mixed populated/blank members.
    """
    tsb = f"{INPUTS_SHEET}!B{member_inputs_row}"
    split = f"{INPUTS_SHEET}!C{member_inputs_row}"
    band1 = f"{INPUTS_SHEET}!D{member_inputs_row}"
    band2 = f"{INPUTS_SHEET}!E{member_inputs_row}"
    earnings_m = f"{earnings_cell}*{split}"
    tax_expr = (
        f"{earnings_m}*{band1}*rate_tier1 + {earnings_m}*{band2}*rate_tier2"
    )
    guard = (
        f'OR({tsb}="",{split}="",{tsb}<=0,{split}<=0,{earnings_cell}<=0)'
    )
    return f"=IF({guard},0,{tax_expr})"


def div296_adj_gain_formula(proceeds: str, cost_base_expr: str, held: str) -> str:
    """Per-asset Div 296 adjusted gain: (proceeds − cost base), 1/3 CGT
    discount applied iff the normalised held>12m flag is "Yes", never to
    losses. Shared by Analyser (per-asset cols + hidden helpers) and
    Comparison (per-register grid) — v3.4 dedup of two byte-identical copies.
    """
    raw = f"({proceeds}-{cost_base_expr})"
    return (
        f'=IF({proceeds}="","",'
        f'IF({raw}<=0,{raw},'
        f'IF({held}="Yes",{raw}*(1-discount_rate),{raw})))'
    )
