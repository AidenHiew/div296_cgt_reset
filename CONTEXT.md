# Division 296 CGT Reset Model — Glossary

Single source of truth for domain terms used in this codebase, the workbook UI,
and conversations with users. When a term here conflicts with code or wording
in the workbook, the workbook follows this file.

Add terms here when they're actively negotiated in a session, not preemptively.
Implementation details (cell addresses, formulas, function names) belong in
code comments or ADRs, not here.

---

## Members

A **Member** is a single SMSF beneficiary whose Total Superannuation Balance
(TSB) is tracked separately. The model supports up to 4 members. Members with
TSB ≤ 0 render as empty data rows but their **placeholder labels** ("Member 1"
… "Member 4") always show on the Comparison tab so the structure is visible
even on a blank workbook.

## TSB — Total Superannuation Balance

Per-member balance. The threshold for triggering Div 296 is calculated against
each member's TSB independently, not against the fund total. "Total fund TSB"
(`SUM` across members) is shown in the Comparison strip only as a Total-row
summary; the individual member TSBs are the values that drive tax.

## Div 296 — Division 296

Treasury's *Better Targeted Super Concessions* additional tax. Applied to a
member's share of fund earnings above the TSB-proportional threshold. The
model implements Year 1 only (no multi-year projection).

**Fund earnings = `MAX(0, sum of Div 296 adjusted gains)`** — capital
gains and losses are netted within the income year at the fund level,
and the net is floored at zero (Div 296 earnings cannot be negative).
This replaced the v3.0 per-asset floor in v3.1; see "Capital losses /
netting" below.

## Threshold 1 / Tier 1

**$3,000,000**. Additional 15% (`rate_tier1`) applied to the portion of
earnings attributable to the slice of TSB between $3m and $10m.

## Threshold 2 / Tier 2 ("$10m tier" / "+25% tier")

**$10,000,000**. Additional 25% (`rate_tier2`) applied to the portion
attributable to TSB above $10m. **Always applied — v3.0 removed the toggle.**
The $10m tier is enacted law; the v2.x `tier10_on` toggle had no Bill-correct
use case and was a footgun (default OFF before v2.5 produced wrong numbers
for $12m+ members). v3.0 removes the option entirely.

## band1 / band2 (per-member proportion bands)

The two-band decomposition of a member's TSB used by the Div 296 tax formula:

- **`band1`** = proportion of TSB in the `$3m–$10m` slice, taxed at `rate_tier1`
  (15%). Excel: `MAX(0, MIN(tsb, threshold_2) - threshold_1) / tsb`.
- **`band2`** = proportion of TSB above `$10m`, taxed at `rate_tier2` (25%).
  Excel: `MAX(0, tsb - threshold_2) / tsb`.

`band1 + band2` = total proportion of TSB above `$3m`. Both bands are
displayed as auto-derived columns on Inputs Section 1 (cols D and E) from
v3.0 onwards, so reviewers can manually reconcile:

> `member tax = earnings × split × (band1 × 15% + band2 × 25%)`

Single source of truth: the per-member tax formula in Analyser and Comparison
reads `band1`/`band2` directly from the Inputs cells, not by recomputing
inline.

## v3.0 cut-over (breaking change)

v3.0 removed the three control-panel toggles (`reset_on`, `tier10_on`,
`discount_on`) and the `proportion_override` cell. Effect on default-config
numbers: **zero** (the toggles' default-ON state was the only Bill-correct
configuration). Breaking changes are in the Python API for `calcs.py`
consumers; see README "What's new in v3.0.0" and the breaking-change manifest.

## Reset (cost-base reset election)

A member's pre-30-June-2026 election to step the Div 296 cost base up to
**market value at 30 Jun 2026** instead of original cost base. Affects Div 296
only — ordinary CGT continues to use original cost base regardless. The
election is irreversible and applies to the whole TSB, not per-asset.

The two scenarios the model compares are:
- **If no reset (default)** — Div 296 cost base = original cost base
- **If elected to reset** — Div 296 cost base = market value at 30 Jun 2026

The "reset trap": for assets currently in an unrealised-loss position, the
reset *creates* a Div 296 gain that didn't exist before.

## Difference

The **signed** delta between the two scenarios:

> `Difference = (If elected to reset) − (If no reset)`

- **Negative** difference (rendered as bracketed-and-coloured negative) means
  the reset *reduces* tax — a saving.
- **Positive** difference means the reset *increases* tax — a cost.

Used consistently across the Comparison tab (headline card, scenario
subtotals, per-member breakdown, per-asset detail). Replaced the v2.4 label
"Change" in v2.5 because "Difference" is unambiguous about being a
two-scenario comparison; "Change" could be read as "year-on-year change".

**v2.5 step 13 — headline-only verbose variant.** The big headline cards on
Comparison use longer, fully-spelled labels to give the client maximum
context on the tile they see first:

| Position | Headline card (verbose)                              | Tables (short — glossary) |
|----------|------------------------------------------------------|---------------------------|
| Left     | "If no Div 296 CostBase Reset (default)"             | "If no reset (default)"   |
| Middle   | "If elected to reset Div 296 CostBase Reset"         | "If elected to reset"     |
| Right    | "Difference (Net Div 296 Tax)"                       | "Difference"              |

Subtotal and per-member tables keep the short form so they stay scannable.
The glossary terms in this file are the short form — they're the canonical
names. The verbose variants are presentation-layer only.

## Headline / per-member / per-asset

Three levels of granularity for the Div 296 tax calculation, all reconciling
to the same total:

- **Headline** — fund-level total Div 296 tax for the year.
- **Per-member** — each member's share, allocated by their proportion of fund
  earnings (auto-derived from TSB split).
- **Per-asset** — each asset's share of the headline, allocated by
  `MAX(0, that asset's adjusted gain) / SUM(positive adjusted gains)`.
  Loss assets contribute $0 to the per-asset allocation — **even though
  they reduce the headline at fund level under v3.1 intra-year netting.**
  Per-asset shares still sum exactly to the headline.

The "Top 10 most affected" panel on Comparison ranks assets by **absolute tax
difference** (v2.5+; previously ranked by absolute gain difference).

## Discount (CGT discount)

The 1/3 super-fund CGT discount on assets held > 12 months. Applies to
ordinary CGT and to Div 296 *gains* (never to losses). Toggle defaults ON.

## Capital losses / netting (v3.1)

Capital gains and losses are netted **within the income year** at the
fund level — for both ordinary CGT and Div 296. Three rules:

1. **Ordinary CGT** uses the s102-5 ITAA 1997 method statement:
   - Sum gross capital gains (split by holding period — long-held gains
     are eligible for the 1/3 discount; short-held are not).
   - Sum gross capital losses.
   - Apply losses to non-discount gains first (taxpayer-favourable;
     preserves the discount on long-held gains where possible — common
     SMSF practice, but not the only legal allocation; the taxpayer may
     elect otherwise per s102-5).
   - Apply the 1/3 discount to the remaining long-held portion.
   - Multiply by the fund CGT rate (15%).
   - Result: the Analyser "Fund Ordinary CGT (after intra-year netting)"
     cell. The per-asset Ord CGT column on the same tab is a STANDALONE
     DIAGNOSTIC VIEW only — greyed out, not summed.

2. **Div 296 fund earnings** = `MAX(0, sum of Div 296 adjusted gains)`.
   Gains and losses net within the year; the net is floored at zero
   (Div 296 earnings cannot be negative per the Treasury Bill).

3. **Per-asset Div 296 tax** is still attributed only to positive-gain
   assets via `MAX(0, my_gain) / SUM(positive_gains) × headline_tax`.
   Loss assets bear $0 Div 296 tax, even when they reduce the headline
   at fund level. Per-asset shares still reconcile to the headline.

4. **Carry-forward losses** = `MAX(0, gross_losses − gross_gains)` at
   the fund level on an ordinary-CGT basis (uses `original_cost_base`,
   not affected by reset election). Display-only — the workbook does
   not consume this against any subsequent year's calc; carry it
   manually if material.

v3.0 used a per-asset silo (ordinary CGT) and per-asset floor (Div 296)
that contradicted both s102-5 and the Div 296 earnings concept. v3.1
reverses both — see calcs.py "v3.1 capital-loss netting" docstring
section for the breaking-change rationale.

---

## What's NOT in this glossary

- Cell addresses (row/col positions on each tab) — see code constants in
  `src/div296/tabs/*.py`.
- Formulas — see `src/div296/calcs.py` (Python mirror) and the workbook itself.
- Versioning, release roadmap, build process — see `README.md`.
- Implementation decisions / trade-offs — would live in `docs/adr/` if any
  were significant enough to record.
