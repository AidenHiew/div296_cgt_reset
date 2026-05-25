# Division 296 CGT Cost Base Reset Model

A Microsoft Excel workbook (`.xlsx`) that illustrates Division 296 tax outcomes for SMSFs and makes the case for pre–30 June 2026 action on assets sitting in an unrealised-loss position (the "reset trap").

**Status:** v2.5.0 — stable. Client-feedback pass #3 over v2.4.0 (8 items: calc-default flip, Inputs member 2, full Comparison rewrite). Workbook, builder, tests, and PDF export all green.

**Previous releases:**
- v2.4.0 (frozen as reference; artifact `dist/Division_296_Model_v2.5.0.xlsx`). Client-feedback pass over v2.3.0 (4 items).
- v2.3.0 (frozen as reference; artifact `dist/Division_296_Model_v2.3.0.xlsx`). Client-feedback pass over v2.2.0 (16 items).
- v2.2.0 (frozen as reference; artifact `dist/Division_296_Model_v2.2.0.xlsx`). Client-readability pass over v2.0.0.
- v2.0.0 (frozen as reference; tag `v2.0.0`, artifact `dist/Division_296_Model_v2.0.0.xlsx`). UX pass over v1.0.0.
- v1.0.0 (frozen as reference; tag `v1.0.0`, artifact `dist/Division_296_Model_v1.0.0.xlsx`).

**What's new in v2.5.0:**
- **Calc default** — `$10m / +25% tier` toggle on Inputs B6 now defaults to ON (was OFF). When OFF the formula taxed all >$3m earnings at only 15%, ignoring the legislated $10m threshold; ON applies 15% to the $3m-$10m slice and 25% additional to the >$10m slice as per the Bill. The toggle remains in the UI for comparison. §12 acceptance numbers updated: Member 1 ($12m TSB) headline tax 28,500→32,722 (reset on) and 157,500→180,833 (reset off); net effect of electing reset 129,000→148,111.
- **Inputs** — Member 2 seeded with TSB $3.5m alongside Member 1's $12m, so the workbook ships with two populated members and the split-balance behaviour is visible out of the box.
- **Comparison — side-by-side strip** — top tile rewritten from a single "Total fund TSB" cell to a per-member mini-table (Member 1..4 + Total). Other three tiles (proportion, discount, tier) vertically merge across the same rows so the strip reads as one panel.
- **Comparison — headline cards, subtotals, per-member, per-asset** — all Change columns flipped to a **signed** convention: Change = If elected − If no reset. A reset that reduces tax now shows as a NEGATIVE figure in red brackets (`($129,000)`). Old positive-savings convention was confusing because the sign meaning differed by section.
- **Comparison — wording standardised** — every section now uses "If no reset (default)" and "If elected to reset" headers (was a mix of "Default outcome", "If you elect", "If you elect the reset by 30 Jun 2026").
- **Comparison — per-member breakdown** — labels for Member 1..4 always shown as placeholders, even when Inputs has blank rows; numeric cells stay blank.
- **Comparison — per-asset detail** — Change column moved from between the panels (old col F) to the far right (col K). Metric switched from gain-delta to **tax-delta** so the visible Change matches the sort metric ("top 10 most affected by tax" instead of "by gain"). Panel titles reworded.
- **Comparison — chart removed** — the v2.3 horizontal bar chart didn't add narrative on top of the per-asset detail table; cut entirely.
- **Layout re-tune** — col widths bumped for the new longer headers; single-page A4 landscape print fit.

**Audience:** internal use by partners, managers and staff; also printed or shared with clients.

> **Illustrative only — not financial, tax, or legal advice.** Confirm against final ATO method, regulations and a registered tax agent / licensed financial adviser before relying on any figure.

---

## Quickstart

```bash
# 1. Install (dev mode + dev tools)
python -m pip install -e .[dev]

# 2. Build the workbook (runs a live recalc check at the end)
python -m div296.build
# -> dist/Division_296_Model_v2.5.0.xlsx
# -> Recalc validation: OK (no Excel error cells).

# 3. Run the test suite (fast dev loop — skips live-recalc)
pytest -m "not slow"

# Full suite incl. live-recalc against §12 numbers (~30s+):
pytest
```

Pass `--no-validate` to skip the post-build recalc check (faster, not recommended). The check uses the pure-Python `formulas` package to recalculate every cell and fails the build if any cell resolves to `#REF!`, `#DIV/0!`, `#VALUE!`, `#NAME?`, `#NULL!`, `#NUM!`, or `#N/A`. If `formulas` isn't installed, the build prints a notice and continues.

> **Known validator false positives (v1.7+).** The pure-Python `formulas` package reports `#VALUE!` errors on the Comparison tab's per-asset detail rows (the sort-by-impact `LARGE` / `MATCH` / `INDEX` chain introduced in v1.7). Excel and LibreOffice both evaluate these cells correctly — the rendered PDFs are proof — and `tests/test_calcs.py` mirrors every formula in pure Python and asserts the §12 acceptance numbers, so calc correctness is verified by the test suite rather than by the recalc gate. When the validator reports errors only on Comparison rows 30–39 of the per-asset detail panel, treat them as a known limitation of the `formulas` library against this specific dependency graph, not as a real defect in the workbook. (Same failure pattern is present in v2.0.0; see the v2.2.0 roadmap notes.)

### Exporting to PDF (client-shareable Comparison page)

```bash
python scripts/export_pdf.py dist/Division_296_Model_v2.5.0.xlsx
# -> dist/Division_296_Model_v2.0.0_Comparison.pdf

# Other tabs / whole workbook:
python scripts/export_pdf.py dist/Division_296_Model_v2.5.0.xlsx --tab Analyser
python scripts/export_pdf.py dist/Division_296_Model_v2.5.0.xlsx --all-tabs
# -> dist/Division_296_Model_v2.5.0_Comparison.pdf  (default tab)
# -> dist/Division_296_Model_v2.5.0.pdf             (--all-tabs)
```

Requires [LibreOffice](https://www.libreoffice.org/) installed (`soffice` on PATH, or the default `C:\Program Files\LibreOffice\program\soffice.exe` on Windows).

The build script is the source of truth. The `.xlsx` is a build artifact and is **not** checked in (see `.gitignore`).

---

## Project layout

```
div296_cgt_reset/
├── README.md                  this file
├── pyproject.toml             openpyxl, pytest, ruff
├── .gitignore                 dist/, __pycache__, .venv, .pytest_cache
├── .gitattributes             *.xlsx -> binary
├── src/div296/
│   ├── build.py               entrypoint: python -m div296.build
│   ├── assumptions.py         rates, thresholds, labels (single source of truth)
│   ├── named_ranges.py        named-range registry
│   ├── styles.py              colours, fills, fonts per spec §11
│   ├── calcs.py               pure-Python mirror of every Excel formula
│   └── tabs/
│       ├── inputs.py
│       ├── analyser.py
│       ├── comparison.py
│       └── notes.py
├── tests/
│   ├── test_calcs.py          asserts spec §12 acceptance numbers
│   └── test_workbook.py       build -> recalc -> read back totals
├── scripts/
│   └── recalc.py              LibreOffice headless recalc helper
├── docs/
│   └── BUILD_PLAN.md          frozen reference copy of the functional spec
└── dist/                      gitignored: Division_296_Model_vX.xlsx
```

---

## Locked design decisions

These supplement [docs/BUILD_PLAN.md](docs/BUILD_PLAN.md) §1.

1. **Source of truth = Python builder.** The `.xlsx` is regenerated by `python -m div296.build`. Never hand-edit the output.
2. **Calculation testing in Python.** `tests/test_calcs.py` mirrors every Excel formula in Python and asserts the §12 acceptance numbers. Catches drift between spec, Python mirror, and Excel formulas.
3. **Single assumptions module.** All rates, thresholds, and labels live in `src/div296/assumptions.py`. The Python builder writes them into named cells; Excel formulas reference those cells (never inline literals). Same constants drive the Python tests.
4. **Per-asset Div 296 column = pro-rata of headline.** Spec §7's fork is resolved: each asset's tax = (its positive Div 296 adjusted gain ÷ sum of positive adjusted gains) × the member-attributed headline. Per-asset sum always ties to the headline.
5. **Hidden provenance cell** on the Notes tab carries `build_version`, `build_date`, and the git short-SHA.
6. **Watermark "ILLUSTRATIVE — NOT ADVICE"** appears diagonally across the Comparison print area, in addition to the small disclaimer line in the header block.
7. **Tamper-evident, not tamper-proof.** Sheet protection is passwordless and documented as such in Notes.

### Deviations from the spec as written

| Spec § | Original | This build |
|---|---|---|
| §2 / §9 | 5 tabs including `Decision` | **4 tabs** (`Inputs`, `Analyser`, `Comparison`, `Notes`). Decision tab removed to keep the model purely illustrative. |
| §4 | Inputs zones: control panel → asset register → members & advanced assumptions (top to bottom) | **v1.5 reorder:** control panel → members → asset register → advanced assumptions. Fund-level setup is now compact and near the top; bulk asset data sits in the middle; set-once constants demoted to the bottom. |
| §5 | Analyser column order: Asset / Proceeds / Original CB / **Div 296 CB** / Ord taxable gain / Ord CGT / Div 296 adj gain / Div 296 tax / Reset impact | **v1.5 reorder:** Div 296 cost base moved from col 4 to col 6, so each cost base sits directly next to the gain it computes. |
| §8 | Lean Comparison panels (3 cols: Asset / Div 296 gain / Div 296 tax), 50 data rows reserved, no chart, no headline summary | **v1.5 redesign:** subtotals moved to the top; three hero metric cards (Scenario A / B / Net effect); fund-context strip; 5-col panels (Asset / Proceeds / Div 296 CB / Div 296 adj gain / Div 296 tax) plus a Δ column between panels; **TOTAL TAX BURDEN row** (Ord CGT + Div 296 tax) added to subtotals; 15 visible rows with overflow note to Analyser; bar chart kept at bottom. |
| §8 footer | "$X saved (green) / $X created (red)" recommendation | **Neutral net-effect calculation only.** Net-effect metric card reads "Net effect (A − B) = $X". No "saves/created" framing, no "you should..." text anywhere in the workbook. |

---

## Caveats documented in the workbook (Notes tab)

Factual disclosures, not recommendations:

1. **Loss-offset divergence from ATO method.** The per-asset Ordinary CGT silo overstates tax in years with offsetting gains and losses, and overstates carry-forward losses. Reconcile with the firm's tax practitioner.
2. **Pension phase not modelled.** Assumes 100% accumulation phase, fund earnings tax = 15%. Retirement-phase assets are 0% — model overstates Ordinary CGT for funds with pension members.
3. **Reset OFF scenario is realised-only.** In reality, a fund that does not elect the reset is taxed under Div 296 on TSB movement (unrealised + realised). The Comparison tab compares realised vs realised.
4. **Wash sale / Part IVA risk.** Selling pre–30 June 2026 purely for the tax outcome and reacquiring sits in anti-avoidance territory (TR 2008/1).
5. **Transaction costs, liquidity, market-timing risk** apply to any pre-reset disposal strategy.
6. **Alternative levers not modelled:** TSB recontribution split, pension commencement, estate timing.
7. **Royal Assent date** is held as an editable cell, not asserted as model truth.

---

## Roadmap

- **v0.1:** scaffolding, empty 4-tab workbook with title banners, build entrypoint, test placeholders.
- **v1.0:** Inputs zones 1–3, Analyser 9-column audit trail + reconciliation panel, Comparison side-by-side panels with neutral footer + watermark, Notes content, pytest suite asserting §12 numbers, live `formulas`-based recalc integration test.
- **v1.1:** response to independent code review (relative-row CF guarded by test, member-prop edge cases, quantity formatting, terminology consistency).
- **v1.2:** post-build recalc validation gate — `python -m div296.build` fails non-zero if any cell resolves to an Excel error sentinel.
- **v1.3:** GitHub Actions CI (lint + tests + build, Python 3.11/3.12/3.13); bar chart on Comparison; `scripts/export_pdf.py` LibreOffice headless PDF export.
- **v1.4:** PDF render fixes — chart cache injection, footnote wrap, recalc.py source-collision fix.
- **v1.5:** Inputs zones reordered (members up, advanced down); sample-data warning badge; Analyser column reorder (Div 296 cost base next to its gain); Comparison redesign (subtotals at top, metric cards, fund-context strip, 5-col panels + Δ column, total-tax-burden subtotal).
- **v1.6:** Comparison PDF fits on one A4 landscape page — chart anchored inline next to subtotals (was below data, on page 2), 5 visible data rows + `fitToHeight=1` removes mid-page whitespace.
- **v1.7:** Manual earnings input removed (Comparison-vs-Analyser divergence eliminated, control panel down to 3 levers); Analyser column label "Proceeds" renamed to "Projected sale proceeds" to match Inputs; Comparison per-asset detail now shows **top 10 assets sorted by |Δ (B − A)|** descending, via LARGE/MATCH/INDEX over a hidden per-register helper grid (cols N/O/P/R).
- **v2.0.0 — UX pass.** No calculation changes; v1.0.0 numbers reproduce bit-for-bit. Highlights:
  - **Comparison:** dead 2-bar headline chart replaced with a per-asset Δ (B − A) horizontal bar chart, sorted by impact.
  - **Analyser:** fund-state strip across the top, row-number column for printed reference, print titles repeat headers across pages, freeze panes removed for free scroll, column-group tinting (sand / slate / sage-teal / gold palette) to separate inputs / ordinary CGT / Div 296 / reset-impact, distinct totals row, gold accent on the Reset impact column.
  - **Inputs:** "Total value" renamed to "Current market value (as at today)" for clarity, column A widened, column resize allowed under sheet protection.
  - **Holistic:** dark-teal section titles unified across all tabs; "ILLUSTRATIVE — NOT ADVICE" print-header watermark now appears on all 4 tabs (was Comparison only).
- **v2.2.0 (this release) — client-readability pass.** Same calc engine; presentation changes only — except for one stale-reference bugfix that affects displayed totals (see below). Highlights:
  - **Plain-English re-labelling** across Comparison. "Scenario A — No reset" / "Scenario B — Reset elected" / "Net effect (A − B)" → **"Default outcome (no election lodged)" / "If you elect the reset by 30 Jun 2026" / "Change if you elect"**. All Greek `Δ` symbols dropped in favour of the word "Change". The "do nothing" framing is gone — not lodging an election by 30 Jun 2026 IS a permanent choice, and the new labels make that explicit.
  - **TSB traffic-light on Inputs.** New row 3 auto-checks whether any member's TSB sits above the $3m threshold and shows a green / red diagnostic so a client can see "does Div 296 even apply to me?" without reading any tax numbers.
  - **Sample-data warning propagated.** Comparison and Analyser now both show the amber "⚠ Sample data detected" badge whenever the asset register still contains the demo P1/S1/L1 codes — prevents an accountant accidentally PDF-exporting a polished-looking client tearsheet built on Mr Sample's commercial property.
  - **Loss-position highlight on Inputs.** Asset register rows where current MV < original cost base are tinted pink, so the unrealised-loss positions (which drive the entire reset-trap thesis) are visible on the data-entry sheet itself.
  - **Year-1 footnote** under the metric cards reminds readers that Div 296 is annually assessed and the same hit recurs each year.
  - **Trap-row legend on Analyser** explains the red row CF ("reset would create Div 296 tax on an asset in unrealised loss") rather than leaving a client to guess what the colour means.
  - **Capital losses CF caption** explains how the CF loss balance can be used (against future realised capital gains within the fund).
  - **Σ → "Total"; humanised state strip.** Totals row no longer leads with a Greek sigma; the state strip reads as a sentence ("Right now you are viewing: reset on, CGT discount on...") rather than a log line.
  - **Reset-toggle scope comment.** Cell comment on the Reset election input explains that the toggle affects Analyser only and that Comparison always shows both paths regardless.
  - **Bugfix (carry-over from v2.0.0):** Comparison's Ordinary CGT subtotal was referencing `Analyser!B73` (a band row) instead of `B74` (the value row) — stale reference from before v2.0.0's state-strip shift. Result: v2.0.0's PDF tearsheets showed `TOTAL TAX BURDEN = Div 296 tax` only. v2.2.0 fixes the reference, so the displayed total now correctly includes Ord CGT. **§12 acceptance numbers are unaffected** (they're computed in `calcs.py` and tested against the Analyser, not the Comparison subtotal).
  - **Recalc validator known limitation documented.** The pure-Python `formulas` package reports 33 spurious `#VALUE!` errors against the Comparison per-asset detail rows (the v1.7 `LARGE`/`MATCH`/`INDEX` sort-by-impact chain). v2.0.0 produces the identical errors — confirmed by running `validate_recalc` against the frozen `dist/Division_296_Model_v2.0.0.xlsx`. Excel and LibreOffice handle the formulas correctly; calc correctness is enforced by `tests/test_calcs.py` (which mirrors every formula in pure Python and asserts §12 acceptance numbers). See Quickstart for the full caveat.

---

## Build provenance

Generated by [Division 296 CGT Cost Base Reset Model build script](src/div296/build.py). See [docs/BUILD_PLAN.md](docs/BUILD_PLAN.md) for the functional specification.
