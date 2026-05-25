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

## Threshold 1 / Tier 1

**$3,000,000**. Additional 15% (`rate_tier1`) applied to the portion of
earnings attributable to the slice of TSB between $3m and $10m.

## Threshold 2 / Tier 2 ("$10m tier" / "+25% tier")

**$10,000,000**. Additional 25% (`rate_tier2`) applied to the portion
attributable to TSB above $10m. **v2.5 default: ON.** When the toggle is OFF
the formula collapses to single-tier (15% on everything above $3m, ignoring
the $10m threshold entirely — useful for comparison but not Bill-correct for
high-TSB members).

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
  Loss assets contribute $0 to the per-asset allocation.

The "Top 10 most affected" panel on Comparison ranks assets by **absolute tax
difference** (v2.5+; previously ranked by absolute gain difference).

## Discount (CGT discount)

The 1/3 super-fund CGT discount on assets held > 12 months. Applies to
ordinary CGT and to Div 296 *gains* (never to losses). Toggle defaults ON.

---

## What's NOT in this glossary

- Cell addresses (row/col positions on each tab) — see code constants in
  `src/div296/tabs/*.py`.
- Formulas — see `src/div296/calcs.py` (Python mirror) and the workbook itself.
- Versioning, release roadmap, build process — see `README.md`.
- Implementation decisions / trade-offs — would live in `docs/adr/` if any
  were significant enough to record.
