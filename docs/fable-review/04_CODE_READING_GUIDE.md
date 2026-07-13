# Fable Review — Curated Code-Reading Guide

> You do not need to read the whole tree. This is a **ranked** list per surface —
> read the ★★★ files in full, skim ★★, and only open ★ if a specific finding needs
> it. Paths are repo-relative from the worktree root. Total ★★★ reading is ~2,000
> lines — deliberately small.

Legend: ★★★ read fully · ★★ skim for shape · ★ reference only

---

## Read first (context — you're already partway through)

- `docs/fable-review/00`–`03` (this folder) — mandate, domain/law, architecture,
  invariants.
- `CONTEXT.md` ★★ — canonical glossary (reset tool). Confirms label wording + sign
  conventions.

---

## ⚠ Seeing the rendered output (you're judging *presentation*)

The two Excel surfaces are built by openpyxl code. Reading `comparison.py` to judge
the client tearsheet is like reviewing a website by reading its DOM builder — you'll
misjudge how it actually looks. Judge each surface at the altitude you can verify:

- **Build the workbook and inspect the produced `.xlsx` directly** (you can load cell
  values, number formats, fills, fonts, column widths, merged ranges via
  openpyxl/pandas — that's real evidence for information architecture & layout):
  - Reset: `python -m div296.build --no-validate` → `dist/Division_296_Model_v3.4.0.xlsx`
  - Ongoing: `python -m div296_calc.build --no-validate` → `dist/ongoing_calculator/…xlsx`
  - (Run from repo root; `--no-validate` skips the slow recalc gate.)
- **For a true visual/print verdict** on the reset tearsheet, the honest artifact is
  a PDF: `python scripts/export_pdf.py <xlsx>` (needs LibreOffice on PATH). If you
  can't render it, **say so** and scope your Excel critique to information
  architecture + workflow (visible in the builder and the built cells), not
  pixel-level visual design.
- **Web** is the easy one: open `web/index.html` in a browser, or read
  `index.html` + `styles.css` — HTML/CSS closely reflect the rendered result.

Don't assert a visual verdict you couldn't actually verify.

---

## Surface 1 — Reset workbook (`src/div296/`)

| File | | Why |
|------|---|-----|
| `calcs.py` | ★★★ | The reset engine. All the tax math in 264 lines. Reproduce **Emma** mentally against the two-band formula here. |
| `_formulas.py` | ★★★ | The shared Excel formula-string builders — the guard idioms live here. Small (75 ln). |
| `tabs/comparison.py` | ★★ | The client-facing tearsheet (1032 ln). **Skim for the presentation model**, not every cell — this is your main "is Excel the right medium / is this layout good" evidence. |
| `tabs/analyser.py` | ★★ | The audit sheet + s102-5 reconciliation panel (900 ln). Skim the fund-summary + reconciliation structure. |
| `tabs/inputs.py` | ★★ | The only data-entry surface (547 ln). Skim the zones + the paste-guard idioms (trip-banner, hidden normalised flag col J). |
| `_recalc_limitations.py` | ★★ | 53 ln — reveals the recalc-gate blind spot (core netting + top-10 panel excluded). Relevant to "how well is this really verified." |
| `build.py` | ★ | CLI + recalc gate wiring + `--edition adviser`. |
| `tabs/class_import.py` | ★ | The manual CLASS Super staging workflow (378 ln) — open only if you critique that flow. |
| `tabs/notes.py`, `assumptions.py`, `named_ranges.py`, `styles.py` | ★ | Reference. |

## Surface 2 — Ongoing calculator (`src/div296_calc/`)

| File | | Why |
|------|---|-----|
| `calcs.py` | ★★★ | The ongoing engine (197 ln). Slice kernel, tsb_ref greater-of, s102-5 netting, carry-forward. |
| `_formulas.py` | ★★★ | 99 ln — the guard-first builders incl. the **`COUNTIF→NA()` fail-loud** year guard (invariant #2). |
| `tabs/calculator.py` | ★★★ | 410 ln — the entire single-sheet UX: members-as-columns layout, year table, protection, sample seed. This IS the ongoing product's presentation. Read it. |
| `tabs/notes.py` | ★★ | The law-basis + approximation caveats (contributions add-back / $3M loss-floor omitted). Relevant to scope framing. |
| `assumptions.py` | ★★ | 55 ln — the year table (2026-27 only) + `UnknownYearError`. The manual-maintenance pain point. |
| `build.py`, `named_ranges.py` | ★ | Reference. |

## Surface 3 — Web (`web/`)

| File | | Why |
|------|---|-----|
| `index.html` | ★★★ | 380 ln — the entire page structure + content. Your main UX/content evidence. |
| `app.js` | ★★★ | 363 ln — the interactivity, state model, render loop, a11y touches (and gaps). |
| `calcs.js` | ★★ | 310 ln — confirm it's a 1:1 port of `src/div296/calcs.py` (the drift-risk finding). Don't re-verify the math cell-by-cell; the invariant is that it mirrors the reset engine. |
| `styles.css` | ★★ | 938 ln — skim the design system (palette, responsive breakpoints, `@media print`). |
| `PROGRESS.md` | ★★ | 58 ln — the team's own open/deferred list (hosting, mobile tables, shareable links). Don't re-report what they already know is deferred; build on it. |
| `build-standalone.mjs`, `README.md` | ★ | The regex-based bundler + deploy notes. |
| `standalone/…html` | ✗ | Generated bundle — **do not read** (1992 ln, redundant). |

---

## Don't spend reads on

- `dist/` artifacts (gitignored / binary), frozen `.xlsx` reference builds.
- The full test suites — the acceptance numbers you need are in `03`. Open a test
  only to confirm a specific invariant.
- `docs/superpowers/`, `docs/handoffs/`, generated user guides — background, not
  needed for the review.
