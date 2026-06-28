# Ongoing Division 296 Calculator — v0.1 Design Spec

- **Date:** 2026-06-20
- **Status:** Draft v4 (v2 = four-lens expert review; v3 = second-pass gaps: income-concept fixes, override×pool rule, CI integration, usability extras; v4 = structural review: sample-vs-goldens split, user-maintained threshold table, all-overridden edge, distinct carry-forwards, pre-fund-tax basis, discount-phase caveat)
- **Branch:** `div296-ongoing-calc/v0.1`
- **Law basis:** Division 296 as enacted — *Treasury Laws Amendment (Building a Stronger and Fairer Super System) Act 2026* (Royal Assent 13 Mar 2026), first income year 2026-27. Validated 2026-06-20 (memory `div296-enacted-law-2026`).
- **Build discipline:** test-driven (TDD), see §10.

## 1. Purpose

A single-year, multi-member Division 296 tax calculator for SMSFs. Plug in a fund's figures for **one income year** and read off **each member's Division 296 tax** plus a fund total. Re-run each year as real numbers arrive.

Separate from the existing year-one **CGT-reset** workbook (which decides *whether to reset cost bases* before 1 July 2026). This tool computes *ongoing-year liabilities* once that decision is behind you.

## 2. Scope

**In scope (v0.1):**
- One income year per run; up to 4 members; per-member calculation + fund roll-up.
- Pooled fund earnings (realised income **less deductible expenses**) allocated by member share %, with an optional per-member earnings override for segregated members.
- A **CGT netting helper** that derives net realised capital gains from raw report figures (s102-5 loss-ordering + 1/3 SMSF discount).
- Enacted-law mechanics: two tiers ($3M / $10M), thresholds via a year lookup table, the greater-of TSB rule, negative-earnings carry-forward (entered as input, produced as output — not auto-rolled).
- Ships with a worked example pre-filled and a documented way to clear inputs; prints cleanly for a client file.

**Out of scope (future):**
- Multi-year rollover / projection (the per-member "new carried-forward loss" output is the seam).
- Return-rate / balance-growth estimators.
- Defined-benefit interest valuation, non-arm's-length income (NALI), and aggregating multiple funds beyond a single TSB figure per member.
- Statutory derivation of earnings from the fund tax return (taxable income − contributions + ECPI − NALI). v0.1 takes realised income components (less expenses) directly; the Notes tab documents the full derivation.

## 3. Architecture

- **New package `src/div296_calc/`** — a sibling to the frozen `div296/` package, which is **not touched**. Shares the repo harness (`ruff`, `pytest`, CI, recalc gate).
- **CI:** the workflow must build **and** test `div296_calc` — `ruff check`, the fast `pytest` suite, and the **strict build** (`python -m div296_calc.build`, recalc gate on). The slow workbook-recalc integration test (§9) runs in the `slow` suite.
- **Reuse decision ("import now, extract `div296_core` later"):**
  - **Reuse `div296.styles`** for fonts/fills/number formats.
  - **Reuse / mirror the s102-5 capital-loss netting** logic (`div296.calcs`, v3.1) for the CGT helper.
  - **Mirror the per-member Excel formula builder shape** (`div296._formulas.per_member_div296_tax_formula`) and the blank-guard discipline (`div296.tabs.inputs`).
  - **Reuse the recalc-gate pattern** (`div296.build`), started **strict** (fail on any error cell; no skip-list copied over).
  - **Do NOT import `div296.calcs.div296_tax_for_member`.** It internally multiplies by `member.split_pct` (would double-apply our share allocation) and bundles guards we compute ourselves. The two-tier tax is 4 lines — re-implement it natively (§5).
- **Own `assumptions.py`:** a `year → (threshold_1, threshold_2, use_greater_of)` table; rates `rate_tier1 = 0.15`, `rate_tier2 = 0.25`; member cap = 4. **The table is seeded with only confirmed years — at v0.1 that is just 2026-27 ($3M / $10M, un-indexed).** Future years' thresholds are CPI-indexed (in $150k / $500k steps) and not yet published, so they are added by hand as the ATO announces them (see §4 unknown-year behaviour and §8 verified-date stamp). A test pins the 2026-27 row to the frozen `div296.assumptions` constants.
- **Own `calcs.py` (first-class domain logic, fully tested — not "glue"):** CGT netting, pooled allocation (incl. expense deduction and override handling), TSB_REF rule, net-earnings + carry-forward, per-member tax, fund roll-up. **No openpyxl import anywhere in the calc modules.**
- **Own `build.py`:** emits `dist/Division_296_Calculator_v0.1.xlsx` (pre-filled sample, print setup). Tabs: `Calculator`, `Notes`.

## 4. Inputs (Calculator tab)

- **Income year** — dropdown; drives `(T1, T2, use_greater_of)` via exact-match lookup. **Hard-fails on a year not in the table** (no silent default). At v0.1 the table holds only **2026-27**; the failure message must tell the user *to add the ATO-published threshold row for that year* (not just raise) — the indexed $3M/$10M figures for 2027-28+ aren't known yet.
- **Fund pooled realised income** → `POOLED_TOTAL`. Labelled "all realised income, **both pension and accumulation phase**; **measured before the fund's own 15% tax** (taxable-income basis — Div 296 sits *on top of* fund tax); **exclude contributions**":
  - Dividends / distributions — **grossed-up (incl. franking credits), per the fund return**
  - Interest · Rent · Other realised income
  - **Net realised capital gain** — from the CGT helper below (not hand-entered)
  - **less: Deductible expenses** — admin/accounting/ASIC fees, insurance premiums, investment expenses, LRBA interest, etc.
- **CGT netting helper** (raw figures off a CLASS/BGL report):
  - Gross capital gains on assets held **>12 months** (discountable)
  - Gross capital gains on assets held **<12 months** (non-discountable)
  - Capital losses (current year **+ brought-forward**)
  - → tool applies s102-5 ordering + 1/3 discount → **net realised capital gain**; any unused capital loss is surfaced as a carry-forward output.
- **Per member (1–4), shown as columns:**
  - Name · Opening TSB (1 Jul) · Closing TSB (30 Jun) · Share % of pooled earnings · Prior-year carried-forward Division 296 loss (default 0) · *Optional* earnings override.
  - TSB cells labelled **"Total Super Balance — ALL funds, from the member's ATO/myGov record (not just this SMSF)."**
- **Override semantics:** a member **with** an override is treated as **segregated** — they earn *only* their override and are **excluded from the pool split** and from the share guard. Common case (no overrides): all members are pooled.
- **Share guard:** the shares of the **pooled** members (those without an override) should total 100% — a **visible soft warning** (red/green conditional-format check cell), **not a hard block** (mid-year exits / reserves legitimately break 100%). **When there are no pooled members (every member overridden), the guard is suppressed** (shows "n/a — all members segregated", not a false ≠100% warning) and the `POOLED_TOTAL` inputs are inert/greyed.
- **Sample + reset:** ships pre-filled with **one coherent worked-example fund** — a single pool, 2–3 pooled members whose shares sum to 100%, and realistic figures — visible on open so the headline use case is self-documenting. (This is *not* the same as the four kernel goldens in §9, which are per-member earnings→tax unit tests and cannot all live in one pool — Jamal's negative earnings can't come from a positive pool. See §9.) A clearly-marked input range / "clear inputs" note lets the user reset to blank.

## 5. Calculation logic

For income year `Y` with `(T1, T2, use_greater_of)`:

**CGT helper (reuse s102-5 netting):**
```
loss          = capital_losses_incl_brought_forward
nondisc_net   = MAX(0, gross_under12m − loss)
loss          = MAX(0, loss − gross_under12m)          # losses hit non-discount gains first
disc_net      = MAX(0, gross_over12m − loss)
net_realised_cg     = nondisc_net + disc_net × (2/3)     # 1/3 discount on the long-held remainder
unused_capital_loss = MAX(0, loss − gross_over12m)       # carries forward (output)
```

**Pooled total (realised income less expenses) and per-member tax:**
```
POOLED_TOTAL = dividends_grossed + interest + rent + other + net_realised_cg − deductible_expenses
               # may be negative (a net realised loss after expenses) → flows to carry-forward below

pooled_members = members with NO override        # their shares should sum to 100%

for each member i:
    earnings_i = override_i              if override_i provided   (segregated; not from pool)
               = share_i × POOLED_TOTAL  otherwise

    TSB_REF_i  = close_i                  if use_greater_of is FALSE (2026-27)
               = MAX(open_i, close_i)     if use_greater_of is TRUE  (2027-28+)

    net_i = earnings_i − prior_loss_i

    if net_i <= 0  or  TSB_REF_i <= T1:
        tax_i      = 0
        new_loss_i = MAX(0, −net_i)
    else:
        band1_i = MAX(0, MIN(TSB_REF_i, T2) − T1) / TSB_REF_i   # the $3M–$10M slice
        band2_i = MAX(0, TSB_REF_i − T2)          / TSB_REF_i   # the slice above $10M
        tax_i   = net_i × (band1_i × 0.15 + band2_i × 0.25)
        new_loss_i = 0

fund_total_tax = SUM of tax_i
```

> **Formula form — read this.** We use the **disjoint-slice** decomposition: 15% on the $3M–$10M slice, 25% (15% + the extra 10%) on the slice above $10M. This is **algebraically identical** to the law's wording "15% on the proportion above $3M, plus an extra 10% above $10M" — both give Emma (TSB 12.9M, earnings 840k) exactly **$115,581.40**. During review, two separate experts mis-paired the slice-form `band2` with a 10% rate and undercounted to $87,256; the mandatory >$10M golden test (§9) catches exactly that. `rate_tier2 = 0.25` is the **total** Div 296 rate on the >$10M slice, **not** 0.10.

**Rounding:** the engine runs at full floating-point precision; **tests assert the exact computed value**. Only *display* cells round (whole dollars). Golden numbers are this engine's output, not sources' rounded figures.

## 6. Outputs (Calculator tab, locked cells)

Per member: earnings used · TSB used (with which rule applied) · net Division 296 earnings · band1 / band2 proportions · Tier-1 tax (15%) · Tier-2 tax (extra) · **Total Division 296 tax** · new carried-forward loss. A member with `TSB_REF ≤ T1` shows an affirmative **"Below $3M — not liable"** status, not a bare $0.

Fund roll-up: **total Division 296 tax across all members** (via `SUM`); plus the CGT helper's **unused capital-loss carry-forward**.

**Two distinct carry-forwards — must be labelled so they are never conflated** (they roll into different inputs next year):
- *Member Division 296 loss* (per-member `new_loss_i`, member-level) → next year's **"prior-year carried-forward Division 296 loss"** for that member.
- *Unused capital loss* (fund-level, s102-5, from the CGT helper) → next year's **"Capital losses (current + brought-forward)"** in the CGT helper.

The output labels and the §8 Notes state this mapping explicitly.

Prior-year loss echo: each member's entered prior-year loss is shown back prominently, stamped "carried forward — verify against last year's output."

Liability block (2–3 plain lines): the tax is assessed to the **individual** (on top of the fund's own 15%); payable **personally** or by **election to release** from a super fund (release-authority deadline applies).

## 7. Excel formula-safety mandates

- **Aggregation uses `SUM`, never `+`-chaining and never `SUMPRODUCT(flag*range)`** (the repo's text-coercion scar). Unused member columns hold `""`; `SUM` coerces text→0, `+` would propagate `#VALUE!`.
- **Guard-first `IF` for any division or blank-render:** e.g. `band1 = IF(OR(TSB_REF="",TSB_REF<=0,TSB_REF<=T1),"",(MIN(TSB_REF,T2)-T1)/TSB_REF)`. `IF` short-circuits → no `#DIV/0!`.
- **Year → (T1,T2,use_greater_of) via exact-match `SUMIFS`** against the small year table — **not** `INDEX/MATCH`/`VLOOKUP` (recalc-gate false-positive list).
- **One column-keyed Python builder** generates all four members' result formulas (no hand-written near-duplicates).
- **Locking:** explicitly unlock input (yellow) cells (`Protection(locked=False)`), then `ws.protection.sheet = True`; result cells stay locked.
- **Number formats:** reuse `styles.FMT_CURRENCY` / `FMT_PERCENT`. The carry-forward loss is a **positive** number → plain currency, **not** `FMT_CURRENCY_DELTA`.
- **Print/page setup:** A4, fit-to-width, repeating header rows, so `Calculator` and `Notes` print cleanly for a client file.
- **Recalc gate strict:** `python -m div296_calc.build` must produce **zero** error cells, no skip-list.

## 8. Notes / Disclaimer tab

- Law basis + Royal Assent date; thresholds-by-year table with a **"thresholds verified [date]"** stamp. State the table is **user-maintained**: at v0.1 only 2026-27 is confirmed; add later years' indexed $3M/$10M from the ATO as published.
- The full statutory earnings derivation (taxable income − assessable contributions + net ECPI − NALI), and a plain statement that v0.1 approximates it via realised income components less expenses.
- **Franking note:** dividends should be the grossed-up amount (cash + franking credits) per the fund return; exact treatment under the final regulations to be confirmed.
- **Reset cost-base link:** for 2026-27 onward, realised capital gains should be measured from the **30-Jun-2026 reset cost base if the fund made that election** — point to the year-one CGT-reset tool.
- Plain statements that, under the **enacted realised-earnings** method: contributions/withdrawals are **not** added back to earnings, and there is **no $3M loss-floor** (both were 2023-draft features).
- **CGT-discount caveat:** the CGT helper applies the 1/3 discount to **all** realised gains held >12 months regardless of pension/accumulation phase — an approximation of the realised-earnings base pending final regs (the statutory base nets ECPI, which may treat pension-phase gains differently).
- **Two carry-forwards — don't conflate:** the *member Division 296 loss* rolls into next year's per-member prior-loss input; the *unused capital loss* rolls into next year's CGT-helper capital-losses input (mirrors §6).
- Caveats: realised-earnings attribution, the negative-earnings floor, and DB valuation factors are **subject to final regulations** — outputs are estimates until confirmed against the enacted Act + final regs.
- Scope note: **accumulation and account-based pensions only**; one calc per member across **all** their super (not per fund). General disclaimer (not personal advice); version watermark.

## 9. Tests (acceptance gate)

Engine output at full precision. The four golden worked examples below are **pure-layer kernel unit tests** (a member's earnings → that member's tax); they are deliberately *not* one fund (Jamal's negative earnings can't share a positive pool). Recomputed values — sources' rounded display values in brackets:

| Case | Inputs | Expected tax |
|---|---|---|
| Jess | TSB 4.5M, earnings 476,625 | **$23,831.25** (source displays $23,829) |
| Emma (two-tier, >$10M) | TSB_REF 12.9M, earnings 840,000 | **$115,581.40** (source displays $115,577) — **mandatory** |
| Jill | TSB 3.1M, earnings 100,000 | **$483.87** (source displays $484) |
| Jamal (loss) | earnings −200,000 | **$0** tax, **$200,000** carry-forward |

Plus the boundary/path tests that drive TDD:
- Boundaries: `TSB = 3,000,000` → $0; `TSB = 10,000,000` → tier-1 only; `TSB < 3M` with earnings → $0; `TSB = 0` → $0.
- Earnings/loss: zero earnings → $0 + $0 loss; prior loss partially absorbs; prior loss == earnings (net 0 → $0 tax, $0 loss); prior loss exceeds earnings → carry-forward.
- **Pooled total:** deductible expenses reduce earnings; `POOLED_TOTAL` negative (expenses/losses exceed income) → members get negative allocated earnings → carry-forward.
- Paths: override ignores share/pool; **overridden member excluded from the pool split** (pool goes to remaining pooled members per their shares); pooled path = share × pool; **override = 0 honoured, not blank**; mixed members.
- TSB rule: 2026-27 closing only; 2027-28 greater-of; greater-of when opening > closing; unknown year raises.
- CGT helper: losses hit non-discount gains first; 1/3 discount on long-held remainder; losses exceeding all gains → net CG 0 + unused-loss carry-forward.
- Guards/roll-up: pooled shares = 100% passes; ≠ 100% warns (soft); overridden + blank members excluded from the share sum; **all members overridden → guard suppressed (no false warning), pooled inputs inert**; fund total = SUM incl. a below-threshold $0 member.
- Regression: **no double-application of share**; **2026-27 thresholds match frozen `div296.assumptions`**.
- **Integration (slow):** build the workbook, recalc via the `formulas` engine, and assert **(a)** the shipped worked-example fund's pinned output cells (total tax + each member's tax) equal the pure-layer values, and **(b)** the four kernel goldens reproduce through the Excel layer when entered as *override* members (segregated, earnings flow directly) — Emma's >$10M result is the mandatory anchor. The recalc gate must yield **zero** error cells.

## 10. TDD build order (red → green → refactor)

1. **`assumptions.py`** — year lookup + `unknown_year_raises` + parity test vs frozen constants.
2. **Tax kernel** — `member_div296_tax(net_earnings, tsb_ref, T1, T2, r1, r2)` against the boundary tests + Emma. Native slice formula.
3. **TSB_REF rule** — the three switch tests.
4. **CGT netting** — s102-5 ordering + discount tests.
5. **Pooled total + allocation** — expenses reduce earnings, negative pool, override-vs-pooled, override excluded from pool, override-zero trap, prior-loss absorb cases.
6. **Fund roll-up + share guard** — roll-up sum; pooled-share guard incl. the all-overridden suppression case; the four kernel goldens driven as override member rows; the two distinct carry-forward outputs.
7. **Excel layer last** — `build.py` + `Calculator`/`Notes` tabs as a projection of the proven pure layer. Wire the threshold `SUMIFS` lookup first and run the strict recalc gate before adding member logic; finish with the slow workbook-recalc integration test.

## 11. Open items

- Year-to-year rollover (auto-carry losses) — deferred; carry-forward output is the seam.
- Confirm realised-earnings attribution (incl. franking gross-up), the negative-earnings floor, and DB factors against final regulations.
- Whether to keep the optional override and the unused-capital-loss output after real-world use.
