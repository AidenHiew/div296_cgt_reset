# Fable Review — Architecture Map

> A synthesis of all three surfaces and — most importantly — **how they relate**.
> The per-surface detail is in the appendix at the bottom; read the relationships
> first, because the cross-cutting story is where your review has the most leverage.

---

## The monorepo at a glance

One repository, **three deliverables on independent version lines**:

```
div296_cgt_reset/
├── src/div296/          Reset/transition workbook   (Python→xlsx)  v3.4  · reset-v3.x line
├── src/div296_calc/     Ongoing calculator          (Python→xlsx)  v0.1  · ongoing-v0.x line
├── web/                 Web explainer + calculator   (static JS)    evolving
├── tests/               reset-tool suite + tests/div296_calc/ (ongoing suite)
├── scripts/             export_pdf.py, recalc.py, build_user_guide_v3.py
├── docs/                specs, plans, BUILD_PLAN.md, user guides, THIS folder
├── CONTEXT.md           canonical glossary (reset tool)
└── README.md            reset-tool README (does NOT yet frame the 3 deliverables)
```

Rough sizes: reset tool ≈ 3,880 LOC src + 2,301 LOC tests; ongoing ≈ 780 LOC src +
610 LOC tests; web ≈ 2,184 LOC (excl. the 1,992-line generated standalone bundle).

---

## How the three engines relate — THE central diagram

There are **three separate implementations of overlapping Div 296 math**, in two
languages, kept in sync **by hand + tests**, not by a shared module:

```
        ┌─────────────────────────────┐
        │  src/div296/calcs.py         │   RESET engine (Python, 264 ln)
        │  + assumptions.py            │   canonical for the reset math
        └──────────────┬──────────────┘
                       │  hand-ported 1:1 (NOT imported, NOT generated)
                       │  guarded only by web/tests/parity.test.js (manual run)
                       ▼
        ┌─────────────────────────────┐
        │  web/calcs.js (310 ln)       │   RESET engine, JS copy
        │  + duplicated ASSUMPTIONS    │   → the website calculator is a RESET calc
        │  + duplicated SAMPLE fixtures│
        └─────────────────────────────┘

        ┌─────────────────────────────┐
        │  src/div296_calc/calcs.py    │   ONGOING engine (Python, 197 ln)
        │  + assumptions.py            │   standalone; does NOT import div296
        └──────────────┬──────────────┘
                       │  rates/thresholds parity-pinned to div296.assumptions
                       │  by tests/div296_calc/test_calc_assumptions.py
                       ▼
                 (no JS/web counterpart — the ONGOING tool has no web surface)

   Shared code between the two Excel tools: ONLY div296.styles (look & feel).
```

**Read this carefully — it drives most of the cross-cutting findings:**

1. **The website calculator models the RESET tool**, not the ongoing tax. There is
   **no web front-end for the ongoing calculator** at all. (Is that the right gap
   to have left open? Fair question for you.)
2. **`web/calcs.js` is a hand-copied 1:1 port** of `src/div296/calcs.py`, with its
   own duplicated `ASSUMPTIONS` and a *third* copy of the §12 sample fixture. Drift
   is caught only by a **manually-run** Node parity test — no CI wires it. This is
   the single biggest structural risk in the repo.
3. **The two Python engines don't share a kernel** — deliberately (the ongoing tool
   re-implements the slice-form kernel natively to avoid a `split_pct` double-apply
   trap in the reset tool's `div296_tax_for_member`). Their *assumptions* are
   parity-pinned; their *kernels* are independent.
4. **The Excel tools share only `styles.py`.** Everything else — formula-builder
   patterns, guard idioms, layout-constant discipline, recalc gate — is
   independently re-implemented in each package (similar patterns, separate code).

---

## Common design DNA (both Excel tools)

Worth knowing because a redesign proposal should respect (or consciously replace)
these:

- **Source of truth = the Python builder.** The `.xlsx` is a gitignored build
  artifact; never hand-edited. `python -m <pkg>.build` regenerates it.
- **Pure-Python `calcs.py` mirror** of every Excel formula = the test oracle
  (`test_calcs`). Excel and Python are two independent implementations kept
  numerically identical by golden-string + acceptance tests.
- **Formula-building layer** (`_formulas.py`) emits Excel formula strings with
  **guard-first / SUM-safe idioms** (blank renders as `""` not `0`; aggregation via
  `SUM` never `+`; threshold lookups via `SUMIFS` wrapped in a `COUNTIF→NA()` fail-
  loud guard; INDEX/MATCH/VLOOKUP avoided because the recalc engine false-positives
  on them).
- **Build-time recalc gate** via the pure-Python `formulas` package — recalculates
  every cell, fails on any Excel error sentinel. **Caveat:** on the reset tool this
  gate has a real blind spot — the s102-5 netting chain and the whole top-10 detail
  panel are *excluded* (the engine can't evaluate them / OOMs), so the core tax
  logic there is guarded by tests, not the gate. The ongoing tool's gate is strict
  (small workbook, no skip-list).
- **Layout-constants-as-source-of-truth** — sheet geometry is module-level integer
  constants threaded through builders + tests. No magic numbers, but inserting a row
  ripples through several files and the golden tests; a spreadsheet-native reviewer
  can't trace the dependency graph without reading Python.
- **Structural caps baked into layout:** **4 members**, **50 asset rows** (reset) /
  4 members (ongoing). Raising them means reworking every tab builder, not bumping a
  constant.
- **Tamper-evident, not tamper-proof:** passwordless sheet protection.
- **"ILLUSTRATIVE — NOT ADVICE"** posture enforced (watermarks, no-recommendation-
  language test on Notes).

---

## Cross-cutting observations to seed your review (not a full critique — your job)

- **Triple-maintained math / drift risk** (the diagram above). Is there a defensible
  single-source-of-truth architecture (shared engine, codegen, WASM-compiled Python,
  a JSON spec both consume)? Or is hand-porting + parity tests actually fine at this
  scale? Argue it.
- **The ongoing tool has no web presence; the web tool only does the reset.**
  Product-shape question: should the public web calculator expose the *ongoing* tax
  (the thing every large-balance member will face annually from 2027), not just the
  one-off 2026 reset decision?
- **Manual multi-step workflows** as load-bearing UX: CLASS Import's 5-step
  copy/paste-special transfer (reset); the ongoing tool's manual year-table upkeep
  and manual carry-forward roll. Candidates for "why not push-button."
- **Excel-as-medium tax:** hidden helper columns, layout-constant fragility, recalc-
  gate blind spots, single-A4 print density (Comparison = 1,032 lines of layout).
  How much of this complexity is inherent to "must be an Excel workbook the adviser
  can open and audit," vs. accidental?
- **README/docs don't yet describe the monorepo** — the root README is still the
  reset tool's alone; a newcomer wouldn't learn the ongoing tool or the website
  exist. Cheap, high-value fix regardless of the bigger redesign.
- **Web specifics:** `calcs.js` duplication (above); the committed `standalone/`
  bundle can silently drift; the "Held >12m?" toggle is a non-semantic `<span>`
  (keyboard/a11y gap); no live hosting yet; content is calculator-first / explainer-
  light.

---

# Appendix — per-surface detail

### A. Reset / transition workbook — `src/div296/` (v3.4, 14 files)

- **Output:** `dist/Division_296_Model_v3.4.0.xlsx` (5 tabs) or `_Adviser_Edition`
  (4 tabs, `--edition adviser` drops CLASS Import + rewrites 2 CLASS-naming strings).
- **Tabs:** Inputs (only data-entry sheet — members, 50-row asset register, advanced
  assumptions) → CLASS Import (staging; full edition only) → Analyser (audit: fund
  summary both scenarios, per-asset table, s102-5 reconciliation panel) → Comparison
  (print-ready A4 client tearsheet: 3 hero cards, subtotals, per-member, top-10-by-
  |Δtax|) → Notes (glossary, 13 caveats, valuation log, provenance).
- **Engine:** `calcs.py` (264 ln) — per-asset ordinary gain + 1/3 discount, scenario-
  dependent Div 296 cost base, s102-5 fund netting, `MAX(0, Σ adjusted gains)`
  earnings, two-band member tax, per-asset pro-rata of headline.
- **File index:** `build.py` 189 · `calcs.py` 264 · `_formulas.py` 75 ·
  `_recalc_limitations.py` 53 · `assumptions.py` 41 · `named_ranges.py` 29 ·
  `styles.py` 75 · `tabs/inputs.py` 547 · `tabs/class_import.py` 378 ·
  `tabs/analyser.py` 900 · `tabs/comparison.py` 1032 · `tabs/notes.py` 293.
- **Tests:** `test_calcs.py` 529 · `test_workbook.py` 1220 · `test_integration.py`
  252 (slow) · `test_formula_golden.py` 129 · `test_adviser_edition.py` 103 ·
  `test_export_pdf.py` 68.

### B. Ongoing calculator — `src/div296_calc/` (v0.1, 9 files)

- **Output:** `dist/ongoing_calculator/Div_296_Ongoing_Calculator_v0.1.0.xlsx` —
  **Calculator** + **Notes** tabs.
- **Layout:** members-as-columns (B–E); labels in col A; sections = pooled income,
  CGT netting helper, per-member block (17 stacked fields), assumptions strip; hidden
  year→thresholds table (cols O–R) + resolved selectors (col S) exposed as named
  ranges; sheet-protected with only inputs unlocked.
- **Engine:** `calcs.py` (197 ln) — slice-form `member_div296_tax`, `tsb_ref`
  (greater-of from 2027-28), `net_capital_gain` (s102-5), pooled allocation with
  segregated-member override, `compute_member/compute_fund`, negative-net → carry-
  forward even below threshold.
- **Enacted-law specifics:** year table seeded **2026-27 only**; unknown year
  **fails loud** (`COUNTIF→NA()`); rates/thresholds parity-pinned to
  `div296.assumptions`.
- **File index:** `calcs.py` 197 · `tabs/calculator.py` 410 · `_formulas.py` 99 ·
  `build.py` 98 · `notes.py` 79 · `assumptions.py` 55 · `named_ranges.py` 14 ·
  `__init__.py` 8. **Tests:** `test_calc_engine.py` 240 · `test_calc_integration.py`
  114 (slow) · `test_calc_formulas_golden.py` 112 · `test_calc_build.py` 98 ·
  `test_calc_assumptions.py` 48.

### C. Web explainer + calculator — `web/` (static, 9 files)

- **Shape:** single-page static site — Div 296 explainer + interactive **reset**
  calculator. Vanilla ES-module JS, no framework/bundler. Optional `build-
  standalone.mjs` inlines everything into one `file://`-openable HTML (committed at
  `standalone/div296-reset-calculator.html`).
- **Calc:** `calcs.js` (310 ln) = hand-maintained 1:1 port of `src/div296/calcs.py`;
  duplicated `ASSUMPTIONS` + §12 sample fixture. `computeComparison()` → both
  scenarios, per-member, per-asset, trap flag, carry-forward.
- **UX:** editable member + asset tables, live recalc on every keystroke (no
  debounce), loss-row flagging, `@media print` A4 tearsheet, countdown to 30 Jun
  2026, sample-data badge. A11y gaps: non-semantic toggle span, inputs labelled only
  by column headers.
- **File index:** `index.html` 380 · `app.js` 363 · `calcs.js` 310 · `styles.css`
  938 · `build-standalone.mjs` 43 · `standalone/…html` 1992 (generated) ·
  `tests/parity.test.js` 81 · `README.md` 54 · `PROGRESS.md` 58.
- **Status:** no live hosting yet (GitHub Pages deferred pending go-ahead).
