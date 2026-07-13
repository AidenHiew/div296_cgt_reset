# Fable Review — Domain & Law Primer

> Read this before the code. The two Excel tools model **different things** under
> **different framings of the law**; conflating them is the most common reviewer
> error. The canonical glossary is the repo's `CONTEXT.md` (reset tool); this file
> distils what you need and adds the ongoing-tool + enacted-law context that
> `CONTEXT.md` predates.

---

## Division 296 in one paragraph

Division 296 ("Better Targeted Super Concessions") is an **additional tax on
superannuation earnings** for members whose **Total Superannuation Balance (TSB)**
is large. It is **two-tier and applied to the slice of earnings attributable to
the TSB above each threshold**:

- **Threshold 1 = $3,000,000** → extra **15%** on the earnings attributable to the
  **$3m–$10m** slice of TSB.
- **Threshold 2 = $10,000,000** → extra **25%** on the earnings attributable to the
  TSB **above $10m**. (Often described in the press as "+10% on top of the 15%";
  in the code's **slice form** the >$10m slice carries the full **25%** — see the
  footgun below.)

Both tools decompose a member's TSB into two proportion **bands** and tax earnings
through them:

```
member tax = earnings × split × (band1 × 15% + band2 × 25%)
band1 = MAX(0, MIN(tsb, 10m) − 3m) / tsb      # proportion in the $3m–$10m slice
band2 = MAX(0, tsb − 10m) / tsb                # proportion above $10m
```

---

## ⚠ The two-tier slice-form footgun (critical — applies to every kernel)

The slice form pairs **band2 with 25%, never 10%**. The two algebraic forms are
equal:

- **Slice form** (what the code uses): `15% × band1  +  25% × band2`
- **Cumulative form**: `15% × (band1+band2)  +  10% × band2`

They give the identical answer. The trap: pairing the **slice-form band2** with
**10%** (instead of 25%) silently **undercounts** the >$10m tier. The canonical
regression anchor:

> **Emma — TSB > $10m — total Div 296 tax = `$115,581.40`.**
> Pairing band2 with 10% wrongly yields `$87,256`. Two prior reviewers made exactly
> this error. If you touch or reason about the kernel, reproduce Emma.

---

## The TWO things this repo models

### (1) The reset / transition **decision** — `src/div296/`
A **one-off, year-one** question: should an SMSF member **elect to reset their
Div 296 cost base to market value at 30 Jun 2026** (instead of original cost base)
before the rules bite? The workbook compares two scenarios side by side —
**"if no reset"** vs **"if elected to reset"** — and surfaces the **"reset trap"**:
for assets sitting in an *unrealised-loss* position, the reset *creates* a Div 296
gain that didn't exist before. It is an **illustrative decision-support** tool
(realised-vs-realised comparison), explicitly **not advice**. Year-1 only; no
multi-year projection.

### (2) The **ongoing** annual tax — `src/div296_calc/`
"Plug in the figures, read off each member's Div 296 tax **for an income year**."
This implements the **enacted law** (see below): realised-earnings basis, indexed
two-tier thresholds via a user-maintained year table, up to 4 members
(members-as-columns), pooled realised income less expenses, s102-5 CGT netting,
greater-of-TSB rule, and negative-earnings **carry-forward** (carried, not
refunded).

### (3) The web explainer + live calculator — `web/`
Public-facing explanation of Div 296 + an interactive calculator. Audience is
broader (general public / SMSF trustees / advisers). Its calc logic is JS —
**watch for whether it duplicates the Python kernels and can drift** (a key
cross-cutting question for your review).

---

## The law-version nuance (don't miss this)

- **Enacted law (Royal Assent 13 Mar 2026):** Div 296 is now **law**. Final shape:
  **realised-earnings basis** (the politically contentious *unrealised* earnings
  were dropped), **indexed** $3m/$10m thresholds, two-tier **15% / +10%** (= slice
  15%/25%), **commences 2026-27** (first test date **30 Jun 2027**), negative
  earnings **carried forward** (no refund, no expiry), **individually assessed**.
- **The ongoing calculator (`src/div296_calc/`) targets this enacted law.** Its
  year table is seeded with **2026-27 only** — future indexed thresholds are
  unpublished, so an unknown income year **hard-fails** (this was a real bug once:
  an unknown year silently produced fictitious tax; now guarded with
  `COUNTIF→NA()`).
- **The reset tool (`src/div296/`) predates the final enacted framing** and models
  the *transition decision* on a realised-vs-realised basis. Its `CONTEXT.md` and
  Notes tab carry the caveats (e.g. "Reset OFF scenario is realised-only";
  pension phase not modelled; prior-year losses not modelled). Treat its framing as
  a **year-one transition lens**, not the ongoing-tax lens.

When you critique "presentation," keep straight which tool a user would reach for:
**one-time transition decision (2026)** vs **annual ongoing compliance (2027+)**.

---

## Essential glossary (condensed — full text in `CONTEXT.md`)

- **Member / TSB** — up to 4 members; each member's TSB independently drives the
  threshold test (not the fund total).
- **band1 / band2** — the $3m–$10m and >$10m proportion bands (formulas above).
- **Reset (cost-base reset election)** — pre-30-Jun-2026 election to step the
  Div 296 cost base up to market value; irreversible; whole-of-TSB; affects Div 296
  only (ordinary CGT keeps original cost base).
- **Difference** — signed delta `= (if reset) − (if no reset)`. **Negative =
  saving** (rendered muted-green/bracketed); **positive = the reset costs money**
  (muted-red — the "trap"). This sign convention is load-bearing across the UI.
- **1/3 CGT discount** — super-fund discount on assets held > 12 months; applies to
  gains (never losses); to ordinary CGT and to Div 296 gains.
- **s102-5 netting (v3.1+)** — capital gains/losses net **within the income year at
  fund level**; losses applied to non-discount gains first (taxpayer-favourable);
  Div 296 earnings = `MAX(0, sum of adjusted gains)` (floored at zero).
- **Carry-forward** — reset tool: display-only unused capital loss. Ongoing tool:
  a member's negative Div 296 earnings carried to future years (an input *and*
  output; not auto-rolled).
- **Headline / per-member / per-asset** — three reconciling granularities; per-asset
  Div 296 tax is pro-rata of the headline by positive-gain share.

---

## Audience & framing constraints

- Everything is **"ILLUSTRATIVE — NOT ADVICE"** (watermarks, headers, disclaimers).
  No "you should…" framing anywhere — neutral net-effect only. Any redesign must
  preserve this posture.
- The Excel tools are handed to / used by **SMSF advisers and accountants** (and
  increasingly general members since the repo went public + MIT). The web surface
  targets a broader lay audience. Consider whether "one product for everyone" or
  "distinct products per audience" is the right call — that's a fair thing to
  opine on.
