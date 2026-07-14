# Fable Review — FINDINGS

*Reviewer: Fable (fresh eyes). Scope per `00_MANDATE_AND_SCOPE.md`: deep on the ongoing calculator, then the cross-cutting product-shape question, light pass on reset + web. Reviewed on `main` @ `e184725` via the `fable-brief` worktree. Date: 2026-07-13.*

**Verification scope note:** I inspected the built `.xlsx` artifacts cell-by-cell via openpyxl (values, formulas, formats, protection, hidden columns, DV, merges, print setup), ran the Python kernels directly, and exercised the web calculator in a live browser via DOM/JS (screenshot rendering was unavailable in this environment, so I make no pixel-level visual claims — layout/behaviour claims below are grounded in the built cells and live DOM behaviour I could actually verify).

> **Opus verification (post-review):** spot-checked Fable's cheapest concrete claims against the code — the web "1 July 2025" date (`web/index.html:55`), the stale "Up to 4 members" comment (`src/div296_calc/assumptions.py:25-26`), and the inverted protection (rate cells built with `_input()` = unlocked, `tabs/calculator.py:231`) all confirmed. The year-table partial-row silent-error hole is reasoned (not executed against a partial-row workbook) but matches the known SUMIFS-no-match-zero footgun family and the same failure mode as the ongoing-calc cold-review MAJOR; treat as high-plausibility, confirm the exact formula before/while fixing.

---

## 1. Executive summary

The engineering discipline here is better than most commercial tax tooling: typed, frozen, openpyxl-free calc kernels that double as test oracles; fail-loud threshold lookups; golden formula-string tripwires; a strict recalc gate on the new tool. The correctness story is genuinely done — Emma reproduces to the cent, and I found **no correctness bug in any shipped artifact**.

The problems are all product-shape, and they cluster around one theme: **the portfolio is pointed backwards in time.** The mature, polished surfaces (reset workbook, reset web calculator) serve a one-off decision whose valuation date — 30 June 2026 — has *already passed*, while the recurring annual obligation (every affected member, every year, from 2027) is served by a v0.1 workbook whose three load-bearing annual workflows are all manual, and one of which is actively obstructed by the workbook's own sheet protection. The web calculator — the widest-reach surface — still counts down to a date in the past and doesn't compute the ongoing tax at all.

The ongoing workbook itself is a good *calculator* but not yet a good *annual tool*. Its kernel and guard idioms are excellent. Its year-table maintenance path (the one thing a user must do every single year) requires unprotecting the sheet and unhiding locked, hidden columns — steps the Notes tab never mentions — and a partially-filled year row produces **silently wrong numbers** that slip past the fail-loud guard. Its protection polarity is inverted: the statutory rates (fixed by law) are unlocked editable inputs, while the year thresholds (the user-maintained part) are locked and hidden.

**Top 5 recommendations (ranked):**

1. **[T1]** Build the **ongoing** tax into the web surface and make it the headline; demote the reset calculator to an "archived 2026 transition decision" page. A hosted page also dissolves the year-table maintenance problem — thresholds update centrally.
2. **[T0]** Close the ongoing workbook's year-table trap: row-completeness guard (a partially-filled year row currently miscomputes silently), unhide/unlock the table, document the maintenance steps in-tool.
3. **[T0]** Invert the protection asymmetry: lock the statutory rate cells (B46:B48), expose the year table. Add the income year to the printed output.
4. **[T0]** Fix the web calculator's verified focus-loss-per-keystroke bug and the stale "From 1 July 2025" commencement copy (enacted law commences 2026-27).
5. **[T0]** Wire `web/tests/parity.test.js` into CI and generate the (currently triplicated) §12 fixtures from one source. Do **not** build a shared cross-language engine — at ~300 lines per kernel it isn't warranted (argued in §3).

Also worth stating up front: **the 4-member cap is a real market gap** — SMSFs have allowed six members since 1 July 2021, and `src/div296_calc/assumptions.py:26`'s comment "Up to 4 members per SMSF" is factually stale. Cheap fix, real coverage.

---

## 2. PRIMARY — Ongoing calculator (`src/div296_calc/`), in depth

### What it is & who it's for

A two-tab Excel workbook (Calculator + Notes), built by Python, that an SMSF accountant or adviser runs once per income year from 2026-27: enter the fund's pooled realised income, CGT components, and up to four members' TSBs/shares; read off each member's Div 296 tax, the fund total, and two carry-forwards (member Div 296 loss, unused capital loss) to roll into next year's run. It implements the enacted law (realised basis, indexed two-tier thresholds via a year table, greater-of TSB from 2027-28, loss carry-forward). Confirmed understanding.

### Engine sanity (pass)

Ran the kernel directly (`src/div296_calc/calcs.py`):

- **Emma = $115,581.40** ✓ (tier 1 $68,372.09 + tier 2 $47,209.30). The wrong slice pairing (band2 × 10%) yields $87,255.81 — matches the documented footgun value, so the kernel is on the correct side of it.
- Seeded sample fund reproduces every anchor in `03`: net CG $160,000, pooled total $350,000, Alice $7,875, Bob $19,263.57, fund $27,138.57 ✓.
- The built workbook's formulas (inspected in `dist/ongoing_calculator/Div_296_Ongoing_Calculator_v0.1.0.xlsx`) mirror the kernel faithfully; the `COUNTIF→NA()` guard is present in S2:S4 and `#N/A` propagates through the tier formulas (the `<=t1_sel` comparisons carry the error). Good.
- Kernel code quality is high: frozen dataclasses, the override-of-0.0 semantics documented and honoured in both Python and Excel (`IF(B27<>"",…)` treats numeric 0 as a real override), negative-net carry-forward accrues even below-threshold with a comment explaining why. No correctness flags on the shipped artifact.

> ### ⚠ CORRECTNESS-ADJACENT — the maintenance path can produce silent wrong numbers
>
> Not a bug in the shipped file — a robustness hole in the tool's **designed annual workflow**. The fail-loud guard (`IF(COUNTIF(year_range,year)=0,NA(),SUMIFS(...))`, `src/div296_calc/_formulas.py:70-77`) only covers *year absent*. The Notes tab instructs the user to add each year's row to the hidden year table (cols O–R). If they add the year and T1 but leave **T2 blank**, `SUMIFS` returns 0 → `t2_sel = 0` → band1 clamps to 0 and band2 becomes `ref/ref = 1.0` → the member's **entire** balance is taxed at the 25% tier — silently, with "✓ thresholds loaded" showing. If they leave **GreaterOf blank** in a 2027-28+ row, `greater_of_sel = 0` → closing-TSB basis instead of greater-of → potentially *under*-stated tax, silently. The Notes text (`src/div296_calc/tabs/notes.py:22-24`) tells users to add "the ATO-published $3M/$10M row" and never mentions the GreaterOf column or its semantics at all.
>
> **Fix (T0):** make each selector fail loud on an incomplete row — e.g. `IF(COUNTIF(year_range,year)=0, NA(), IF(COUNT(matched row P:R)<3, NA(), SUMIFS(...)))` — or a single `year_row_complete` guard cell feeding the banner, plus a Notes sentence explaining GreaterOf (`0` for 2026-27, `1` from 2027-28).

### Presentation critique

**The medium question first.** Is a single-sheet Excel workbook the right medium for a tool a fund runs every year from 2027? As the *audit-grade artifact an accountant attaches to a workpaper file* — yes, and the execution suits that role: every intermediate (earnings used, TSB ref, band1/band2, tier1/tier2) is a visible row, which is exactly what an auditor wants. As the *primary delivery vehicle* — no, and the reasons are specific, not aesthetic:

1. **The annual-maintenance contradiction (the biggest single flaw).** The one action the tool requires every year — adding the new ATO-published threshold row — is behind three undocumented barriers. The year table lives in **hidden** columns O–R; its cells are **locked** (verified: `O2.protection.locked = True`); and the sheet is **protected**. To follow the Notes tab's instruction, a user must know to unprotect the sheet (passwordless, but nothing says so), unhide four columns, and fill four cells whose fourth column ("GreaterOf") is explained nowhere. Combine this with the silent-zero hole above and the tool's most safety-critical workflow is its least designed. The v0.1 year table being seeded with only 2026-27 makes this *guaranteed* to bite in year 2 — this isn't an edge case, it's the plan of record.

2. **Protection polarity is inverted.** The statutory rates — 15%, 25%, 1/3 discount, fixed by enacted law — are **unlocked, input-styled cells** (B46:B48, verified unlocked in the built file). The year thresholds — the part the user is *supposed* to maintain — are locked and hidden. A stray keystroke in B46 silently rescales every output, there's no guard or conditional format watching it, and the assumptions strip sits *outside* the print area (`A1:E44`), so a printed/PDF output doesn't even disclose the rates it used. Lock the rates; expose the table; extend the print area (or echo the rates and income year in the printed header).

3. **It's a single-year scratchpad wearing an annual tool's job description.** The two carry-forwards are the workbook's only stateful outputs, and the roll-forward is a hand transcription described in a 9-pt italic footnote (row 44) and a Notes paragraph: row 36 → next year's row 26, row 19 → next year's row 17. There is no "close of year" affordance, no income-year in the filename convention, the print header, or the footer (the footer carries the *tool* version, not the *year*), and the sample badge urges users to "overwrite" — implying one file reused per year, which destroys last year's record unless the user invents their own convention. Nothing needs a database here; it needs the workflow to be *designed*: a boxed "Year-end → next year" block that stages the two numbers side by side with copy instructions, an income-year cell echoed in the print header, and one Notes line on file-per-year practice.

4. **Members-as-columns: right call, wrong cap.** The layout itself is good — four parallel audit columns beat four stacked blocks, and the transparency rows are a feature, not clutter. But `MEMBER_COUNT = 4` (`assumptions.py:26`, with the stale "Up to 4 members per SMSF" comment) excludes 5–6 member funds, legal since July 2021 and exactly the large-balance family funds Div 296 targets. The layout constant mostly propagates (`MEMBER_COLS` derives from it), but the fix isn't a pure constant bump: `_band()` hardcodes `A{row}:E{row}` merges and `range(1, 6)` fills (`tabs/calculator.py:139-143`), as do the badge/liability merges and the `A1:E…` print area. Half a day, not five minutes — but do it before this tool gets real users.

5. **Output framing is thinner than the audience needs.** The tool computes per-member tax, then summarises liability in three generic italic lines (rows 42–44). The people who owe the money are individuals; the natural deliverable is a per-member statement ("Bob — Division 296 assessed: $19,264; payable personally or by release authority"). At minimum, enrich the Status row from "Liable" to "Liable — $19,264", and colour it (the only conditional formatting on the sheet today is on the share-guard cell; the actual *results* get none).

6. **Notes tab likely clips its own caveats.** Every paragraph row is fixed at 46 pt over a merged A:H range (~79 chars/line ≈ 3 lines ≈ ~240 chars), but the earnings-basis paragraph is 319 chars and the scope/disclaimer 256 (verified in the built file). Merged cells don't auto-fit height in Excel, so the two most legally-important paragraphs are probably truncated on screen and in print. I couldn't render pixels to prove it, so verify in Excel — but the arithmetic says clip. Cheap fix: taller heights or per-paragraph computed heights.

7. **Smaller frictions, all real:** the B3 year dropdown's DV list is `$O$2:$O$11` — one real year plus nine blank rows in the dropdown; the sample-data badge fingerprints only B6 + B24, so clearing those two leaves the sample CGT figures (B15:B17) un-badged; "Share % of pooled earnings" doesn't warn at point of entry that TSB is all-super while the pool is this-fund-only (Notes covers it, but the mismatch bites at the input cell); the long ⚠ year-warning text in D3 overflows columns D–E (width 18 each) — fine while F is empty, but fragile.

### Concrete improvements (prioritised)

1. **[T0]** Year-table row-completeness guard (`_formulas.py`) + GreaterOf documentation in Notes + unhide/unlock the year table + in-tool maintenance instructions. *(Closes the ⚠ callout. Do this first.)*
2. **[T0]** Lock B46:B48 (rates); include income year + rates in printed output; income-year echo in the header.
3. **[T0]** Six-member support (B–G) + fix the stale comment.
4. **[T0]** "Year-end → next year" carry-forward staging block; Status row shows the amount; conditional formatting on totals.
5. **[T0]** Notes tab row heights; broaden the sample fingerprint; trim the DV list to non-blank years.
6. **[T1]** Web front-end for this engine (see §3 — the strategic move; the workbook remains the audit artifact).
7. **[T2]** Not warranted. The workbook does not need a greenfield rebuild — it needs the T0 list and a web sibling.

---

## 2b. LIGHT — Reset workbook + Web

### Reset workbook (`src/div296/`) — brief pass

Strengths a fresh eye notices immediately: the layered information architecture (Inputs → Analyser for the audit trail → Comparison for the client → Notes for the caveats) is genuinely good product design in Excel; the paste-guard tripwires and sample-data banners on Inputs are defensive UX most spreadsheet tools never get; the s102-5 reconciliation panel is the right answer to "why is this number what it is." Two small things worth recording, then leave it alone:

- **[T0]** The Comparison headline scenario labels read "If no Div 296 CostBase Reset (default)" and "If elected to reset Div 296 CostBase Reset" (Comparison!A21/E21 in the built v3.4.0 file) — "CostBase" unspaced, and the second label says "reset … Reset" twice. This is the client-facing tearsheet's *headline*; it should read like the web copy ("If no reset (default)" / "If elected to reset"), settled in `CONTEXT.md` first per house rule.
- **[T0-plan]** This tool's audience ends with the 2026-27 lodgment season. Not a criticism — a sunset to *plan*: freeze it after the last plausible election date, mark it archived, and redirect polish budget to the ongoing line.

The recalc-gate blind spot (`_recalc_limitations.py`) is honestly documented and compensated by tests; at this maturity I wouldn't invest further there.

### Web (`web/`) — brief pass

The content is strong — the three-card explainer, the "for the accountant" deep-dive, the applicability gate, and the trap callout are better plain-English Div 296 material than most professional publishers have shipped. But:

> ### ⚠ CORRECTNESS (content) — stale commencement date on the public page
>
> `web/index.html:55`: "From **1 July 2025**, Division 296 adds tax on super balances above $3 million." Under the enacted Act (Royal Assent 13 Mar 2026), Div 296 commences **2026-27** — from 1 July 2026, first test 30 June 2027. "1 July 2025" is the abandoned 2023-draft start date. This is the most public sentence in the whole project and it states the law's commencement wrong. One-line fix.

- **[T0 — verified bug]** Typing in any numeric input loses focus after **every keystroke**. `numInput`'s `input` handler calls `render()` (`web/app.js:70-74`), which rebuilds the members/assets tables via `innerHTML`, destroying the input mid-typing. Verified in a live browser: after one simulated keystroke, `document.activeElement` falls to `<body>`. Entering an 8-digit TSB means re-clicking the field 8 times — the calculator is effectively unusable for real data entry. Fix: on `input`, update state and re-render *outputs only* (`renderResults` + banners); rebuild the input tables only on add/remove row.
- **[T0]** The hero countdown now permanently reads "30 June 2026 … has passed" — the page's top-of-fold element is a dead clock. Pivot the framing: the valuation date is fixed/past; the live deadline is the election with the 2026-27 return. (This also reinforces §3: this surface's relevance is expiring in real time.)
- Known-deferred items in `PROGRESS.md` (hosting, a11y toggle, mobile tables) — not re-reported; the toggle-`<span>` a11y gap is worth bundling into the focus fix since both touch the same render loop.

---

## 3. Cross-cutting findings

### Triple-maintained math — keep the hand-port, fix the process around it

Recommendation: **do not build a shared engine.** The "three engines" frame mildly overstates the avoidable duplication. The Excel formula layer is *inherently* a second implementation whatever you do — that's the price of the medium, and `calcs.py`-as-oracle is already the correct mitigation. The only avoidable duplication is `web/calcs.js` (310 lines) mirroring `src/div296/calcs.py` (264 lines). At that size, codegen, WASM-compiled Python, or a JSON formula spec are all more code and more failure modes than the thing they'd replace. Hand-port + parity tests is the right architecture at this scale. But two process holes make it riskier than it needs to be:

1. **The parity test is manually run.** A test CI doesn't run is a test that doesn't run. Wiring `node web/tests/parity.test.js` into the existing Actions matrix is an afternoon.
2. **The fixtures are triplicated** (§12 sample lives in Python, in `calcs.js`, and in the workbook seed). Generate one `fixtures.json` from the Python suite and have the JS test consume it — drift in *expected values* becomes impossible, and the hand-port only has to keep the *logic* in step.

If (and only if) the ongoing tool gets a web front-end, apply the same pattern there: Python canonical, small hand-ported JS kernel, generated fixtures, parity in CI. Revisit a shared engine only if a fourth surface ever appears.

### Product shape — the web should headline the ongoing tax. Make the call: yes.

The audience arithmetic is lopsided. The reset is one decision per fund, decided in the 2026-27 lodgment window; its addressable audience goes to zero during 2027. The ongoing tax touches every affected member, every year, indefinitely — and its calculator is *simpler* than the reset one (a dozen scalar inputs, no 50-row asset register). Today the widest-reach surface computes only the expiring decision, and the evergreen obligation has no web presence. That's the wrong way around.

The web version of the ongoing calculator also **dissolves this review's biggest T0 finding**: a hosted page ships new ATO thresholds centrally the day they're published. No hidden columns, no unprotect-and-unhide ritual, no partially-filled year row — the year-table maintenance problem simply isn't a problem on the web. That single fact is most of the argument.

Convergence model: **one site, two tools** (ongoing = headline; reset kept but reframed as "the 2026 transition decision" and visibly archived once the election window closes). **Two workbooks stay** as the adviser/audit-grade artifacts, downloadable from the site — don't merge them; they model different things under different law framings, and a merged workbook would blur the law-version boundary the repo is currently careful about. **Don't build one product for everyone:** the lay member wants a number and an explanation (web); the accountant wants an audit trail they can file (workbook). Different jobs — the two-medium split is right; only the *coverage* (which tax gets which medium) is wrong.

### Excel-as-medium — inherent for the reset, partly accidental for the ongoing

For the reset tool, most of the Excel tax was inherent: advisers audit in Excel, the CLASS import is a spreadsheet workflow, the tearsheet prints. The hidden helpers and layout-constant fragility were the cost of doing that job properly, and it's done. For the ongoing tool, the ledger is different. The visible audit-trail rows: inherent, and good. The hidden year table, hidden selector cells, protection ceremony, manual roll-forward, and manual threshold upkeep: **accidental** — year-one-model DNA applied to a tool whose job (12 inputs, 4 outputs, annually) doesn't need them. A web front-end removes exactly the accidental parts and keeps the workbook for what Excel is actually good at here: being filed.

### Docs coherence

- The root `README.md` still describes only the reset tool. A newcomer to the now-public repo can't discover the ongoing calculator or the website exist. Cheap, high-value: a three-deliverable table at the top (the one in `02_ARCHITECTURE_MAP.md` is already written — move it).
- The "1 July 2025" web error is a symptom: **law facts have no single owner across surfaces.** The reset Notes, ongoing Notes, and web copy each restate the law independently. A short `LAW_BASIS.md` (dates, thresholds, rates, basis, citations) that all three surfaces copy from — with a change-sweep checklist — costs an hour and prevents the next drift.

---

## 4. Optimisation / redesign plan

Ordered by recommended sequence. Items 1–7 are independent of the strategic decision in item 8; do them regardless.

> **[T0] 1. Close the year-table trap (ongoing)**
> **Problem:** Partially-filled year row → silent wrong tax (T2 blank → whole balance at 25%; GreaterOf blank → wrong TSB basis). Table hidden + locked + undocumented; the guard only catches *absent* years.
> **Proposal:** Row-completeness check feeding `NA()` + the banner; unhide/unlock O–R (or relocate the table to a visible block); Notes explains all four columns incl. GreaterOf semantics; DV list trimmed to non-blank years.
> **Payoff:** The tool's mandatory annual workflow stops being its biggest hazard.
> **Cost/risk:** Small; golden-string tests need updating for the new selector formulas.
> **Lands in:** `src/div296_calc/_formulas.py`, `tabs/calculator.py`, `tabs/notes.py`.

> **[T0] 2. Fix protection polarity + print disclosure (ongoing)**
> **Problem:** Statutory rates unlocked (stray edit silently rescales everything); year thresholds locked; printed output shows neither rates nor income year.
> **Proposal:** Lock B46:B48; echo income year (+ optionally rates) in the print header; extend print area or add a footer line.
> **Payoff:** Tamper surface matches what should actually change. **Cost/risk:** Trivial.
> **Lands in:** `src/div296_calc/tabs/calculator.py`.

> **[T0] 3. Fix web input focus loss + stale law date (web)**
> **Problem:** Verified: every keystroke in a numeric field destroys the input (full-table `innerHTML` rebuild inside the field's own `input` handler) — data entry effectively unusable. And `index.html:55` states the wrong commencement year on the project's most public sentence.
> **Proposal:** On `input`: update state, re-render outputs only; rebuild tables only on add/remove. Correct the copy to the enacted 2026-27 commencement; pivot the dead countdown to election-window framing. Regenerate `standalone/`.
> **Payoff:** The public calculator becomes usable and legally current. **Cost/risk:** Low; keep parity numbers pinned.
> **Lands in:** `web/app.js`, `web/index.html`.

> **[T0] 4. Wire web parity into CI + single-source the fixtures (cross-cutting)**
> **Problem:** The only drift guard on a hand-ported engine is a manually-run test; the §12 fixture exists in three copies.
> **Proposal:** Node step in the existing Actions workflow; generate `fixtures.json` from the Python suite, consume it in `parity.test.js`.
> **Payoff:** The declared biggest structural risk becomes continuously enforced. **Cost/risk:** Trivial-to-small.
> **Lands in:** `.github/workflows/*`, `web/tests/parity.test.js`, new generated `web/tests/fixtures.json`.

> **[T0] 5. Six-member support + annual-roll affordance + output polish (ongoing)**
> **Problem:** 4-member cap excludes legal 5–6 member SMSFs (stale comment at `assumptions.py:26`); carry-forward roll is a footnote, not a workflow; results carry no conditional formatting; Notes paragraphs likely clip at fixed 46-pt heights.
> **Proposal:** `MEMBER_COUNT = 6` + fix hardcoded `A:E` merges/fills/print-area; "Year-end → next year" block staging the two carry-forwards; Status row shows the amount; Notes rows sized to content; broaden the sample fingerprint.
> **Payoff:** Covers the actual market and behaves like the annual tool it claims to be. **Cost/risk:** Half a day + test updates.
> **Lands in:** `src/div296_calc/assumptions.py`, `tabs/calculator.py`, `tabs/notes.py`.

> **[T0] 6. Monorepo README + LAW_BASIS.md (docs)**
> **Problem:** Root README describes one of three deliverables; law facts restated independently per surface (that's how "1 July 2025" survived).
> **Proposal:** Three-deliverable table up top (lift from `02`); one law-basis file all surfaces copy from, with a change-sweep checklist.
> **Payoff:** Discoverability for a public repo; no next drift. **Cost/risk:** An hour.
> **Lands in:** `README.md`, new `LAW_BASIS.md`.

> **[T0] 7. Comparison headline label copy (reset — optional)**
> **Problem:** Client tearsheet headline reads "If no Div 296 CostBase Reset (default)" / "If elected to reset Div 296 CostBase Reset" — unspaced "CostBase", doubled "reset".
> **Proposal:** Match the web wording; settle in `CONTEXT.md` first.
> **Payoff:** The most client-visible strings read professionally. **Cost/risk:** Trivial + golden-test updates.
> **Lands in:** `src/div296/tabs/comparison.py`, `CONTEXT.md`.

> **[T1] 8. Web front-end for the ongoing calculator — the strategic item**
> **Problem:** The evergreen annual tax has no web presence; the web serves only the expiring reset decision; the workbook's year-table burden is structural to the offline medium.
> **Proposal:** A second page on the existing site: hand-ported JS kernel of `div296_calc/calcs.py` (~200 lines), same fixture-generation + CI-parity pattern as item 4, thresholds served with the page so they're always current. Ongoing becomes the site headline; reset reframed as the archived 2026 transition tool. Workbook stays as the downloadable audit artifact.
> **Payoff:** Highest-leverage move in the portfolio: right tax, right audience, right medium — and it deletes the year-table maintenance class of problems for web users.
> **Cost/risk:** Days, not weeks. Requires the parked hosting decision (GitHub Pages now viable; repo public/MIT) — **maintainer decision needed first**. Depends on item 4's fixture pattern.
> **Lands in:** `web/` (new page + `calcs-ongoing.js` + tests).

> **[T1] 9. Plan the reset tool's sunset**
> **Problem:** A mature tool whose audience ends with the 2026-27 lodgment season will otherwise keep absorbing polish budget.
> **Proposal:** Freeze after the election window; mark archived in README/site; redirect maintenance to the ongoing line.
> **Payoff:** Focus. **Cost/risk:** None — a decision, not code. **Maintainer decision needed** on the date.
> **Lands in:** `README.md`, site copy.

> **[T2] 10. Greenfield rebuild — not recommended**
> I looked for a case for a from-scratch product (single web app, shared engine, unified workbook generator) and the numbers don't support it: the kernels are ~200–300 lines each, the correctness scaffolding is the project's crown jewel and would have to be rebuilt, and every problem in this review is addressable in place (T0) or by one new page (T1). No T2 item is warranted.

---

## 5. What's already good (don't throw these away)

- **The calc-kernel-as-oracle pattern.** Frozen dataclasses, openpyxl-free, docstrings that *teach the footgun they guard* (`div296_calc/calcs.py:7-11`). The best thing in the repo; any redesign should keep Python canonical.
- **Fail-loud discipline.** The `COUNTIF→NA()` year guard is genuinely superior engineering — it survives DV bypass-by-paste, and `#N/A` propagates through the tier formulas so a bad year visibly breaks everything instead of quietly zeroing. Extend it (item 1); never dilute it.
- **The guard-first formula idioms** (`SUM` not `+`, blank renders as `""`, no INDEX/MATCH) and the strict no-skip recalc gate on the ongoing tool.
- **Golden formula-string tests + acceptance anchors** as a regression contract — this is what makes every T0 above safe to do.
- **The transparency rows** in the ongoing Calculator (band1/band2, tier splits, earnings-used) — resist any urge to hide them; they're the audit story.
- **Sample-data tripwires** on every surface (workbook badges, web badge, paste guards) — a class of error most tools never think about.
- **The sign-colour convention** (negative = saving = muted green), consistently load-bearing across workbook, web, and docs.
- **The web content itself** — the three-card explainer, the accountant deep-dive, and the applicability gate are excellent plain-English tax communication. The bugs are in the input loop, not the substance.
- **Honest documentation of limits** — `_recalc_limitations.py`, the Notes tab's approximation caveats, `PROGRESS.md`'s own open-items list. Rare and valuable.
