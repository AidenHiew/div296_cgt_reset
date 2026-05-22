# Division 296 Cost Base Reset Model — Build Plan

**Deliverable:** a single Microsoft Excel workbook (`.xlsx`) — the "Division 296 Cost Base Reset Model".
**Audience for the finished file:** partners, managers and internal staff; also printed/shown to clients.
**Purpose:** illustrate Division 296 tax outcomes and make the case for pre–30 June 2026 action on assets sitting in an unrealised-loss position (the "reset trap").

This document is the single source of truth. Build the workbook to this spec. It is a **functional spec** — exact calculation logic, column order, layout zones and worked acceptance numbers are prescribed; you choose the exact cell addresses, named ranges and openpyxl mechanics.

---

## 0. Context an implementer needs (read first)

Division 296 is enacted Australian law (Royal Assent 13 March 2026), applying from 1 July 2026. It adds tax on the **realised** earnings attributable to the portion of a member's Total Superannuation Balance (TSB) above set thresholds. Key facts this model relies on:

- **Two parallel cost bases.** When an SMSF elects the one-off **CGT cost base reset** at 30 June 2026, each asset gets a *second* cost base used **for Division 296 purposes only**. Ordinary income-tax CGT continues to use the **original** cost base. So one sale produces **two different gains** and **two different tax answers**. This duality is the entire point of the tool.
- **The reset is all-or-nothing and irrevocable.** It applies to *every* CGT asset held at 30 June 2026 — you cannot cherry-pick. Assets in a loss position are also reset down to their (lower) market value. This is why a loss-making asset can become a Division 296 *gain* on later sale → "the reset trap".
- **Tax sits on a proportion.** Div 296 taxes only the proportion of earnings attributable to the balance above $3m (e.g. a $12m TSB → (12m−3m)/12m = 75%). Additional rate is **+15%** for the $3m–$10m band and **+25%** above $10m.
- **Discount.** Inside super, capital gains on assets held >12 months get a one-third (33⅓%) CGT discount. It applies to both the ordinary and the Div 296 taxable gain. Capital **losses are not discounted**.
- **Assessed to the individual**, not the fund; multi-member funds split pooled earnings (in reality via an actuarial certificate — this model uses a simpler user-entered split, see §6).

> The model is **illustrative**, not advice. Build the caveats in §10 into the file verbatim.

---

## 1. Locked design decisions (do not re-litigate)

1. **Basis:** enacted law. Reset is a toggle. $10m/+25% tier is switchable, **off by default**.
2. **Display model ("Option D"):** one fund-wide master switch drives the headline; the analyser always shows the *with-reset* and *without-reset* gains side by side; each asset carries a **reset-impact flag** (green = reset reduces Div 296 gain; red = reset *creates* a Div 296 gain it didn't have = the trap).
3. **Mechanics:** proportion above $3m **auto-calculated** from TSB, with a manual override cell. 33⅓% discount is a **toggle (default ON)** applied to *both* gain calcs. $10m tier off by default.
4. **Rows + reconciliation:** lean per-asset rows on the print view; a portfolio **reconciliation panel** (ordinary CGT payable, Div 296 tax payable, capital losses carried forward). Full 8-column audit trail on the Analyser tab.
5. **Earnings source:** Part A "Division 296 earnings" is **auto-fed** from the asset analyser total **by default**, with a switch to enter a manual figure instead.
6. **Scale:** asset register holds **50+ rows** and auto-expands; each asset has a **valuation-source/date** field. Up to **4 members**, attributed by a user-entered **split %** (defaults to a single member at 100%).
7. **Decision view:** a focused **"hold vs sell before 30 June 2026"** panel that triggers only for the red-flagged loss assets.

---

## 2. Workbook structure — 5 tabs

| # | Tab name | Role | Editable? |
|---|----------|------|-----------|
| 1 | `Inputs` | The only data-entry sheet (3 zones: control panel, asset register, members & advanced assumptions) | Yes (green cells only) |
| 2 | `Analyser` | Live full audit-trail view + mirrored lever strip + totals + reconciliation panel | No (reads Inputs) |
| 3 | `Comparison` | Print-ready side-by-side (reset OFF vs reset ON) on one landscape A4, with header block | No (reads Inputs) |
| 4 | `Decision` | "Hold vs sell before 30 June 2026" panel for loss-flagged assets | No (reads Inputs) |
| 5 | `Notes` | Terminology, caveats, valuation log, "not advice" statement | Header fields editable |

Tabs 2–5 must read **everything** from `Inputs` via formulas — no duplicated input cells anywhere. Enter data once.

---

## 3. Naming convention (use these exact labels everywhere)

- Part A headline figure → **Division 296 earnings**
- Gain columns → **Ordinary taxable capital gain** and **Div 296 adjusted taxable capital gain**
- Tax columns → **Ordinary CGT** and **Div 296 tax**
- Cost base columns → **Original cost base** and **Div 296 cost base**
- Abbreviate "Division" → "Div" only where column width forces it (e.g. headers). Never abbreviate in the Notes tab.

---

## 4. Tab 1 — `Inputs` (layout map)

Three visually separated zones, top to bottom. Use a title band at top.

### Zone 1 — Control panel (the demo levers)
Five labelled controls, each a single editable cell with a dropdown (data validation) where applicable:

| Control | Type | Default | Drives |
|---------|------|---------|--------|
| Reset election | dropdown `ON`/`OFF` | `ON` | which cost base every asset uses for Div 296 |
| $10m / +25% tier | dropdown `ON`/`OFF` | `OFF` | whether the above-$10m slice is taxed at 25% |
| CGT discount | dropdown `ON`/`OFF` | `ON` | whether the 33⅓% discount applies to both gains |
| Div 296 earnings source | dropdown `Auto`/`Manual` | `Auto` | Part A earnings = analyser total (Auto) or the manual cell (Manual) |
| Manual earnings (used only if Manual) | currency input | blank | the override figure |

These five controls are the canonical inputs. The same five values are **mirrored (display-only, via formula)** as a thin strip near the top of the Analyser, Comparison and Decision tabs so the user never has to switch tabs to read the current scenario. (Mirrors are read-only links, not editable copies.)

### Zone 2 — Asset register
One row per asset; build **50 rows** of capacity and make every downstream formula/range cover the full block (and ideally use a structured Excel **Table** or full-column-bounded ranges so adding rows 51+ extends automatically). Columns, in order:

1. **Asset code** (text, input)
2. **Asset name** (text, input)
3. **Quantity** (number, input)
4. **Original cost base** ($, input) — total, not per-unit
5. **Total value** ($, input) — current/most-recent value, informational
6. **Market value at 30 Jun 2026** ($, input) — *the load-bearing input; this becomes the Div 296 cost base when reset = ON*
7. **Valuation source / date** (text, input) — e.g. "Independent val, 30/06/26" or "ASX close 30/06/26"
8. **Projected sale proceeds** ($, input) — assumed future disposal price (timing not modelled)
9. **Held > 12 months?** (dropdown `Yes`/`No`, input) — drives discount eligibility for that asset

All nine are **editable inputs** (green fill, blue text). Everything computed about the asset lives on the `Analyser` tab, not here — keep the register clean.

### Zone 3 — Members & advanced assumptions

**Members block** (4 rows; row 1 defaults filled, rows 2–4 blank):

| Member | TSB ($, input) | Split % of fund earnings ($, input) | Proportion above $3m (auto) | Proportion override (input, optional) |
|--------|----------------|--------------------------------------|------------------------------|----------------------------------------|

- Split % across the 4 members must sum to 100%; add a visible check cell that flags if it doesn't.
- "Proportion above $3m (auto)" = `MAX(0,(TSB − threshold_3m)/TSB)`.
- If the override cell is non-blank, use it instead of the auto figure (`IF(override="",auto,override)`).
- Single-member use: member 1 split = 100%, members 2–4 blank.

**Advanced assumptions block** (editable, but set-once — place at the bottom, lightly de-emphasised):

| Assumption | Default |
|------------|---------|
| Div 296 additional rate — tier 1 ($3m–$10m) | 15.0% |
| Div 296 additional rate — tier 2 (above $10m) | 25.0% |
| Threshold 1 | $3,000,000 |
| Threshold 2 | $10,000,000 |
| CGT discount rate | 33.333% (i.e. taxable = gain × 2/3) |
| Indexation increment — threshold 1 | $150,000 |
| Indexation increment — threshold 2 | $500,000 |

(Indexation increments are informational for now — thresholds for 2026–27 are the un-indexed $3m / $10m. Hold the increments as labelled cells for future use; do not auto-index.)

> **Formula rule:** every rate/threshold/percentage above must live in its own assumption cell and be referenced by cell address in formulas — never hardcode `0.15` etc. inside a formula. (Skill requirement.)

---

## 5. Tab 2 — `Analyser` (the full audit trail)

A read-only mirror-lever strip at the top, then one row per asset (covering all 50 register rows; blank register rows produce blank/zero output, not errors — guard divisions). Column order, left → right:

| Col | Header | Logic |
|-----|--------|-------|
| 1 | Asset (name + code) | from register |
| 2 | Proceeds | = projected sale proceeds |
| 3 | **Original cost base** | from register |
| 4 | **Div 296 cost base** | `IF(reset_on, MV_30Jun2026, original_cost_base)` |
| 5 | **Ordinary taxable capital gain** | raw gain = proceeds − original cost base; if gain>0 and held>12m and discount ON → ×2/3; if gain<0 → keep full loss (no discount) |
| 6 | **Ordinary CGT** | `MAX(0, ordinary_taxable_gain) × tier1_normal_rate` where the normal fund CGT rate = 15%. Losses → $0 (tracked separately as carry-forward, see reconciliation). |
| 7 | **Div 296 adjusted taxable capital gain** | raw = proceeds − Div 296 cost base; discount applied same way as col 5 (gains ×2/3 if held>12m & discount ON; losses kept full) |
| 8 | **Div 296 tax** | see §7 (proportion × rate, with optional $10m tier). Negative adjusted gains contribute $0 tax but still flow into the netting of earnings — see §7 note. |
| 9 | **Reset impact** | `(col7 with reset) − (col7 without reset)`; render flag: ≤0 = green "saves"; >0 = red "trap". (Implement by computing the with/without figures; you may add two helper columns, hidden or to the right, to hold both.) |

**Totals row** under the block: sum of proceeds, ordinary CGT, Div 296 adjusted gain, Div 296 tax.

**Reconciliation panel** (below totals, 3 metric cards or a small block):
- **Ordinary CGT payable** = sum of col 6.
- **Div 296 tax payable** = the member-attributed total (see §7) — this is the headline Div 296 tax.
- **Capital losses carried forward** = sum of `MAX(0, −ordinary raw gain)` across assets sold at a loss (i.e. total ordinary capital losses crystallised).

**Visual flag:** any asset row where ordinary raw gain < 0 **and** Div 296 adjusted gain > 0 is the trap → shade that row's background red (`#FBE9E9`) and the Div 296 figures red text (`#A32D2D`).

---

## 6. Member attribution

- Fund-level **Division 296 earnings** = total of all assets' **Div 296 adjusted taxable capital gain** (col 7), floored at 0 for tax purposes but show the raw net too. (Per the locked scope, earnings = realised Div 296 gains only; dividends/interest/other earnings are deliberately excluded.)
- If `earnings source = Manual`, use the manual earnings cell instead of the asset total.
- Each member's earnings = fund earnings × that member's **split %**.
- Each member's Div 296 tax computed per §7 on *their* share, using *their* proportion above $3m.
- Fund Div 296 tax (the headline) = sum across members.
- **Footnote (required):** "Member split is a user-entered assumption. A real multi-member fund determines each member's share of Division 296 earnings via an actuarial certificate based on time-weighted average balances; this model does not replicate that calculation."

---

## 7. The Division 296 tax formula (precise)

For a member with attributed earnings `E`, TSB, and proportion `p`:

**Tier off (default):**
`Div296 tax = E × p × tier1_rate` where `tier1_rate` = 15% (assumption cell), and `p = proportion above $3m` (auto or override).

**Tier on ($10m/+25%):** split earnings across the two bands by balance proportion:
- proportion in band 1 = `MAX(0, MIN(TSB, threshold2) − threshold1) / TSB`
- proportion in band 2 = `MAX(0, TSB − threshold2) / TSB`
- `Div296 tax = E × band1_prop × tier1_rate + E × band2_prop × tier2_rate`
(When TSB ≤ threshold2, band2_prop = 0 and this collapses to the tier-off result.)

**Edge guards:**
- If TSB ≤ threshold1 → proportion = 0 → tax = 0.
- Guard all `/TSB` divisions against blank/zero TSB (`IF(TSB>0, …, 0)`) to avoid `#DIV/0!`.
- Negative net earnings → Div 296 tax = 0 (no negative tax).

> **Per-asset Div 296 tax (col 8 on Analyser):** for display, attribute the headline tax back to assets pro-rata on their positive Div 296 adjusted gain, OR (simpler and acceptable) show per-asset `adjusted gain × p × tier1_rate` as an illustrative per-asset figure and reconcile the total in the panel. Pick one approach and note it. The **authoritative** Div 296 tax number is the member-attributed total in §6.

---

## 8. Tab 3 — `Comparison` (print-ready, one landscape A4)

Header block at top (editable cells):
- Firm name
- Logo placeholder (a merged cell labelled "[ logo ]" the user can drop an image into)
- "Prepared for: __________" (client)
- "Prepared by: __________"
- "Date: __________"
- One-line disclaimer: "Illustrative only — not financial, tax or legal advice."

Body: two panels side by side, both reading the same register but one forcing reset OFF and the other reset ON (independent of the master switch, so the comparison always shows both):
- **Scenario A — No reset:** per-asset Div 296 cost base = original cost base; show asset, Div 296 adjusted gain, Div 296 tax; subtotal Div 296 earnings + Div 296 tax.
- **Scenario B — Reset elected:** Div 296 cost base = MV 30 Jun 2026; same columns.
- Keep these panels **lean** (cost base + the two key figures), not the full 8-column trail.

Footer strip: **Net effect of electing the reset = Scenario A tax − Scenario B tax** → show "$X saved" (green) or "$X created" (red), plus the two totals (e.g. "$202,500 → $42,750"). Include the one-line reminder that the loss asset adds Div 296 gain it didn't previously have.

**Print setup:** landscape, fit to 1 page wide, A4, sensible margins, repeat header block; set print area to the panels + footer.

---

## 9. Tab 4 — `Decision` ("hold vs sell before 30 June 2026")

Triggers **only** for assets flagged as the trap (ordinary raw gain < 0 AND Div 296 adjusted gain > 0 under reset). For each such asset, two columns side by side:

- **Path 1 — Hold into the reset:** the reset locks in the low Div 296 cost base (= MV 30 Jun 2026), so a later sale produces a Div 296 adjusted gain and Div 296 tax on a gain that is **not economically real**. Show the Div 296 tax cost.
- **Path 2 — Sell before 30 June 2026:** the asset is disposed before the reset date, so (a) it never enters the reset, and (b) it **crystallises an ordinary capital loss** that carries forward to offset future ordinary capital gains. Show the ordinary capital loss banked and the Div 296 tax avoided ($0).

Per-asset rows + a **portfolio total** ("total Div 296 tax avoided" and "total carry-forward losses crystallised by selling before 30 June"). End with the plain-language takeaway line.

If no assets are flagged, show a neutral message ("No loss-position assets detected in the register — the reset trap does not currently apply.").

---

## 10. Tab 5 — `Notes` (build verbatim)

Include:
- **Terminology** key (the §3 names + one-line definitions).
- **Caveats:**
  - Illustrative tool only — not financial, tax or legal advice. Confirm against final ATO method, regulations and a registered tax agent / licensed financial adviser.
  - Division 296 is enacted law (Royal Assent 13 Mar 2026); some operational detail sits in regulations — check current ATO guidance before relying on any figure.
  - The reset is **all-or-nothing and irrevocable**; this model lets you toggle it freely **for comparison only**.
  - 30 June 2026 **valuations are the load-bearing input** — garbage in, garbage out.
  - Multi-member split is a user assumption, not an actuarial certificate (per §6).
  - Scope deliberately excludes dividends, interest and other earnings — Div 296 CGT only.
- **Valuation log:** a small table mirroring each asset's valuation source/date for the audit trail.
- **Version/date** and "model logic per build plan dated [today]".

---

## 11. Formatting & visual standards

- **Font:** Arial throughout (skill default). Title 16pt bold, section bands 10–11pt bold white on dark teal (`#1D3B34`), body 10pt.
- **Colour conventions (industry standard — follow exactly):**
  - **Blue text (`#0000FF`)** = hardcoded inputs the user changes.
  - **Black text** = formulas/calculations.
  - **Green text (`#008000`)** = cells that pull from another sheet in the same workbook (all of tabs 2–5).
  - **Green fill (`#E1F5EE`)** = editable input cells (so "type here" is unmistakable).
  - Trap highlighting: row fill `#FBE9E9`, figures `#A32D2D`. Winner flags green `#0F6E56`.
  - Cost-base columns accented blue (`#185FA5`) headers; Div 296 gain accented teal (`#0F6E56`).
- **Number formats:** currency `$#,##0;($#,##0);"-"` (negatives in parentheses, zeros as "-"); percentages `0.0%`; quantities plain integers; dates as text where used as labels.
- **Protection:** lock all non-input cells; leave green input cells unlocked; protect sheets (no password, or a documented one) so users can't accidentally overwrite formulas. Leave the Notes header fields unlocked.
- **Data validation:** dropdowns for all ON/OFF, Auto/Manual, Yes/No cells; a sum-to-100% check on member splits; optional input-range validation on currency cells (≥0 where appropriate, but allow negative proceeds? no — proceeds ≥ 0).
- **No gridlines** on presentation tabs (Comparison, Decision); freeze header rows on Inputs and Analyser.

---

## 12. Sample data + acceptance tests (ship pre-loaded; verify against these)

Ship the workbook with these **three illustrative assets** pre-entered in the register (clearly flagged as examples in a cell note), single member, TSB **$12,000,000**, all assets **held > 12 months = Yes**, discount **ON**, $10m tier **OFF**.

| Asset | Qty | Original cost base | MV 30 Jun 2026 | Projected proceeds |
|-------|-----|--------------------|-----------------|--------------------|
| Commercial property | 1 | $800,000 | $2,400,000 | $2,600,000 |
| Listed shares parcel | 5,000 | $300,000 | $520,000 | $600,000 |
| Loss-making holding | 2,000 | $500,000 | $100,000 | $200,000 |

Member: TSB $12,000,000 → proportion above $3m = (12m−3m)/12m = **75%**.

### Expected outputs with **reset ON**, discount ON, tier OFF (use as self-check; round to whole dollars):

| Asset | Ordinary taxable gain | Ordinary CGT | Div 296 adjusted taxable gain | Div 296 tax (per-asset illustrative) |
|-------|----------------------|--------------|-------------------------------|--------------------------------------|
| Commercial property | $1,200,000 | $180,000 | $133,333 | $15,000 |
| Listed shares parcel | $200,000 | $30,000 | $53,333 | $6,000 |
| Loss-making holding | ($300,000) loss | $0 | $66,667 | $7,500 |
| **Totals** | — | **$210,000** | **$253,333** | **$28,500** |

Reconciliation: Ordinary CGT payable **$210,000**; Div 296 tax payable **$28,500**; Capital losses carried forward **$300,000**.

> Derivations to verify the engine:
> - Property ordinary: (2,600,000−800,000)=1,800,000 × 2/3 = 1,200,000 × 15% = 180,000.
> - Property Div296: (2,600,000−2,400,000)=200,000 × 2/3 = 133,333 × (75%×15%=11.25%) = 15,000.
> - Loss asset ordinary: (200,000−500,000) = −300,000 loss → CGT 0, carry-forward 300,000.
> - Loss asset Div296: (200,000−100,000)=100,000 × 2/3 = 66,667 × 11.25% = 7,500.  ← **the trap**

### Expected with **reset OFF** (Div 296 cost base = original):

| Asset | Div 296 adjusted taxable gain | 
|-------|-------------------------------|
| Commercial property | (2,600,000−800,000)=1,800,000 × 2/3 = $1,200,000 |
| Listed shares parcel | (600,000−300,000)=300,000 × 2/3 = $200,000 |
| Loss-making holding | (200,000−500,000) = ($300,000) loss → $0 taxable |
| Div 296 earnings (net) | 1,200,000 + 200,000 + 0 = **$1,400,000** |
| Div 296 tax | 1,400,000 × 11.25% = **$157,500** |

### Comparison footer self-check:
- Scenario A (no reset) Div 296 tax = **$157,500**.
- Scenario B (reset) Div 296 tax = **$28,500**.
- **Net effect = $129,000 saved** by electing the reset — *but* the loss asset contributes $7,500 of Div 296 tax under reset that it wouldn't under "no reset", which is the Decision-tab argument for selling it before 30 June 2026.

> Note for implementer: if you net losses against gains within Div 296 earnings differently (e.g. allow the loss asset's negative adjusted gain to reduce total earnings rather than flooring at 0 per asset), the totals shift. **Decision taken: floor each asset's taxable gain at $0 for the per-asset display, and compute Div 296 earnings as the sum of positive adjusted gains** (capital losses are quarantined, not netted against the Div 296 gain in-year — consistent with CGT loss treatment). Document this choice in the Notes tab so it's transparent and easily changed if the firm's view differs.

---

## 13. Tech notes for the build

- **Format:** `.xlsx`. Opens and recalculates normally in Microsoft Excel and Google Sheets — no special software required by the end user.
- **Build library:** `openpyxl` (formulas + formatting). Use Excel **formulas** for every calculation — never compute in Python and paste static values (the file must stay live when inputs change).
- **Recalc/verify step:** openpyxl writes formula *text* but not cached results. To verify zero formula errors before delivery, recalculate the file. Preferred: a headless **LibreOffice** recalc (e.g. the `recalc.py` approach from the xlsx skill) which opens the file, recalculates all sheets, and reports any `#REF!/#DIV/0!/#VALUE!/#NAME?`. **If LibreOffice is not installed in your environment:** either install it, or skip the auto-recalc and instead (a) statically verify the formula logic against the §12 acceptance numbers, and (b) note that Excel will calculate all values automatically the first time the user opens the file. Either path yields a valid workbook.
- **Quality bar:** zero formula errors; all assumptions in referenced cells (no magic numbers in formulas); ranges cover all 50 register rows and degrade gracefully on blank rows (guard divisions); cross-sheet links use green text.
- **Suggested named ranges** (optional but recommended for readable formulas): `reset_on`, `tier10_on`, `discount_on`, `earnings_source`, `manual_earnings`, `rate_tier1`, `rate_tier2`, `threshold_1`, `threshold_2`, `discount_rate`, and the member TSB/split cells.

---

## 14. Definition of done

- [ ] 5 tabs built per §2, reading from `Inputs` only.
- [ ] Control panel toggles all work: reset ON/OFF, $10m tier, discount, earnings source.
- [ ] Analyser shows the 8-column audit trail, totals, reconciliation panel, and red trap highlighting.
- [ ] Comparison prints on one landscape A4 with header block + net-effect footer; both scenarios shown regardless of master switch.
- [ ] Decision tab triggers correctly for the loss asset and shows hold-vs-sell with portfolio totals.
- [ ] Notes tab carries terminology, caveats (§10), valuation log.
- [ ] Sample data pre-loaded; outputs match §12 acceptance numbers exactly (whole-dollar rounding).
- [ ] Zero formula errors on recalc; input cells green/blue, formulas black, cross-sheet links green; inputs unlocked, rest protected.
- [ ] Reads/recalcs cleanly when opened in Microsoft Excel.
