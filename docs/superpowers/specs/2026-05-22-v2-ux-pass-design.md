# Division 296 Model — v2.0.0 UX / Layout / Formatting Pass

**Status:** Design — pending implementation plan
**Date:** 2026-05-22
**Author:** brainstorming session, user-driven
**Target version:** v2.0.0 (v1.0.0 frozen as reference; see user's local project memory `project_v2_versioning.md`)
**Branch:** `v2/ux-pass` (TBD at plan stage)

## Goal

Address concrete UX issues found in the v1.0.0 test-drive of `dist/Division_296_Model_v1.0.0.xlsx`, plus a holistic visual polish pass, shipping as a v2.0.0 line. v1 stays untouched as a reference build.

## Scope

Four tabs, five issue areas. **No** logic / formula / calculation changes. **No** test changes beyond what's required to keep the suite green after layout shifts. **No** dependency changes.

---

## 1. Comparison tab — Replace 2-bar chart with per-asset Δ chart

**Problem (observed):** The current Scenario A vs Scenario B bar chart renders as a blank frame overlapping the per-asset detail table (rows 27+, cols F:K). Root cause is twofold — `plotVisOnly` defaults to `True` so the hidden L:M source columns produce an empty plot, and the chart anchor at `F20` overlaps Panel B.

**Decision:** Drop the Scenario A vs B 2-bar chart entirely. Replace with a **horizontal bar chart of per-asset Δ values** (column F of the per-asset detail block), sorted descending — same sort order as the table.

**Why:**

- The 2-bar chart was redundant with the existing metric cards (which already show the two scenario totals in 22-pt bold).
- A per-asset Δ chart shows **the concentration of the benefit** — which assets actually move the dial under the reset election. This is information the cards and table don't surface visually.
- Scales with register size — more useful at 30 assets, not less.

**Placement:** Anchored below or alongside the per-asset detail table, full panel width (A:K). Print area extended to include it. Chart source data must not be in hidden columns (move the per-asset Δ values to the visible per-asset detail block, or set `plotVisOnly=False` on the chart).

**Out of scope:** Animating, interactivity, chart styling beyond the standard openpyxl `BarChart` with type `"bar"` (horizontal), no legend, no axis titles. Data labels showing $ values per bar.

---

## 2. Analyser tab — Free scroll + visual polish

### 2.1 Remove freeze panes
Delete the `ws.freeze_panes = f"A{PERASSET_HEADER_ROW + 1}"` line at `analyser.py:315`. User wants free scroll.

### 2.2 Polish package (locked design)

Apply all of the following:

**(B) Reinforced section bands** — already present at rows 3, 10, 18, 72; ensure they use the same dark-teal `SECTION_BAND_FILL` / `SECTION_BAND_FONT` styling Comparison uses (cross-tab consistency).

**(C-revised) Reset impact column accent** — col I header `#C7A752` (gold edge), data cells `#FFF8E6` (existing soft-gold fill from Comparison). Signals "this is the column you're here to read."

**(1) State strip** — new row inserted between the title and the lever band, showing the current scenario state at a glance:

> **Current scenario:** Reset `[ON]` · CGT discount `[ON]` · $10m / +25% tier `[OFF]` · **Headline Div 296 tax: $XX,XXX**

Tinted background `#EFF5F3`, left-border accent `#1D3B34`. Pulls live from named ranges + Analyser headline cell.

**(2) Row-number column** — new leftmost column showing `#1, #2, #3…` matching the Inputs register row. Existing data shifts right by one column (per-asset block goes A-I → B-J; helpers K-L stay hidden). Width ~5. Light grey fill `#F2F2F2`, muted text colour `#888`.

**(3) Print titles** — `ws.print_title_rows = '<header_row>:<header_row>'` so the per-asset header row repeats on every printed page. Concrete row number resolved at plan stage (state-strip insertion shifts existing row constants by 1).

**(4) Distinct totals row** — row 70 (after column shift: still totals row). Double top border, bold font, fill `#E6ECEA`. Clear visual stop.

### 2.3 Column-group tinting (replaces zebra striping)

Per-asset table column zones, header row + data cells (5% wash):

| Column zone | Header fill | Data fill |
|---|---|---|
| Asset / # | `#1D3B34` (existing dark teal) | white |
| Projected sale proceeds | `#D7CBB8` (sand) | `#F7F4EE` |
| Original CB · Ord taxable gain · Ord CGT | `#C9D5DA` (slate) | `#F4F6F7` |
| Div 296 CB · Div 296 adj gain · Div 296 tax | `#B0C9BD` (sage-teal) | `#F2F6F4` |
| Reset impact | `#C7A752` (gold) | `#FFF8E6` |

**Header text colour:** the three tinted headers use dark-teal `#1D3B34` text; the gold Reset impact header uses white text.

**Zebra striping NOT applied.** Column tints provide visual structure; zebra + column-tint produces muddy cells and conveys less information.

**Trap-row CF (existing v1 conditional formatting on rows where ord raw gain < 0 AND col 7 > 0)** stays, with red fill winning over the column tints.

**Mockup:** [`docs/mockups/2026-05-22-analyser-v2.html`](../../mockups/2026-05-22-analyser-v2.html) — user-approved 2026-05-22.

---

## 3. Inputs tab — Protection scope + column width

### 3.1 Allow column/row resize under protection
On all 4 tabs (Inputs, Analyser, Comparison, Notes):

```python
ws.protection.formatColumns = False
ws.protection.formatRows = False
```

Cells stay locked. Formulas stay safe. User can drag column borders. Zero data risk.

### 3.2 Widen Inputs column A
Change `widths[0]` in `inputs.py:238` from `10` to `32`, so longest label (e.g. `"Split % sum (must equal 100%)"`) fits without truncation by default. Other column widths unchanged.

---

## 4. Asset register — Rename "Total value" column

`REGISTER_HEADERS[4]` in `inputs.py:74`: rename `"Total value"` → `"Current market value (as at today)"`.

Column stays in place (Inputs col E). No formula changes — the field is display-only audit evidence (no downstream calculation reads it), paired with the existing "Market value at 30 Jun 2026" column as a today-vs-30Jun valuation record.

`Asset.total_value` field name in `calcs.py:27` may optionally be renamed to `current_market_value` for clarity — open for plan-stage decision (cost: rename across `comparison.py:171` and the `SAMPLE_REGISTER_ROWS` tuple unpacking).

---

## 5. Holistic polish — Cross-tab consistency

Apply across all 4 tabs:

### 5.1 Title colour
All tab titles use `TITLE_FONT` with colour `#1D3B34` (dark teal). Update `TITLE_FONT` in `styles.py` if it doesn't already carry this colour; Analyser and Notes currently use the default and need to inherit.

### 5.2 Print header watermark
All 4 tabs get:
```python
ws.oddHeader.center.text = "ILLUSTRATIVE — NOT ADVICE"
ws.oddHeader.center.size = 28
ws.oddHeader.center.color = "CCCCCC"
```

Currently only Comparison has this. Critical compliance signal — must travel with any printed page.

### 5.3 Section bands
Already consistent on all 4 tabs (dark-teal `SECTION_BAND_FILL`); no change needed beyond what's covered by §2.

### 5.4 Sample-data badge
Stays as-is on Inputs row 2.

---

## Out of scope (explicitly)

- Any logic / formula / calculation change
- Row-height harmonisation across tabs
- Body-font harmonisation (already consistent)
- Tab colour coding
- Notes tab dedicated polish pass
- Sparklines, status glyphs, reconciliation tick marks, last-updated footer (rejected during brainstorming)
- v1-to-v2 migration / data-import tooling
- README / BUILD_PLAN updates (will be needed at PR stage but are documentation work, not part of this design)

---

## Reference: file impact map

| Change | File(s) touched |
|---|---|
| §1 Per-asset Δ chart | `src/div296/tabs/comparison.py` |
| §2.1 Remove freeze panes | `src/div296/tabs/analyser.py:315` |
| §2.2 State strip, row-num col, print titles, distinct totals | `src/div296/tabs/analyser.py` (significant — column shift) |
| §2.3 Column-group tinting | `src/div296/tabs/analyser.py`, `src/div296/styles.py` (new palette tokens) |
| §3.1 Protection scope | All 4 `src/div296/tabs/*.py` |
| §3.2 Col A width | `src/div296/tabs/inputs.py:238` |
| §4 Rename column header | `src/div296/tabs/inputs.py:74`; optionally `src/div296/calcs.py:27` + downstream |
| §5.1 Title colour | `src/div296/styles.py` (TITLE_FONT) |
| §5.2 Print header watermark | All 4 `src/div296/tabs/*.py` |
| Tests | `tests/` — likely need row-number adjustments after column shift in Analyser |
| Version bump | `src/div296/__init__.py` (__version__ → "2.0.0") |
| Build output filename | `src/div296/build.py` if filename is hardcoded |

---

## Versioning & workflow

Per user's local project memory `project_v2_versioning.md`:

- New branch `v2/ux-pass` (or per-issue branches merged into a v2 integration branch) off `main`
- v1.0.0 tag stays as reference; nothing on the v1 release branch is touched
- PR + green CI required to merge to main
- Per global CLAUDE.md: no `git push` without sentinel `push this branch now`
- v2.0.0 tag applied after final merge

## Verification gates

- `pytest -m "not slow"` green (fast unit suite)
- Build artifact opens cleanly in Excel without warnings
- PDF export via `scripts/export_pdf.py` succeeds and renders all 4 tabs faithfully
- Manual eyeball pass — confirm each spec item visible in the built workbook
- v1.0.0 build still reproducible from tag `v1.0.0` (sanity check that v2 work didn't accidentally touch v1 reference)

## Open questions deferred to plan stage

1. Rename `Asset.total_value` field in dataclass, or keep field name and only change header? Lower-risk to keep; cleaner to rename.
2. Comparison per-asset Δ chart — anchor below the detail table, or to the right of it inline? Mockup pending.
3. Analyser column shift — single-PR or split into two (column shift first, then visual polish)? Single is simpler; split is safer for review.
