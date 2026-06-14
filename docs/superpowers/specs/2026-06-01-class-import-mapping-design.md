# Design — CLASS Super → Div 296 Asset Register import (v3.3.0)

**Date:** 2026-06-01
**Status:** Awaiting user review
**Branch (intended):** v3.3 feature branch off the v3.2 line

## Problem

The firm exports an **Investment Summary Report** (CSV) from CLASS Super and
currently hand-keys each holding into the Div 296 model's **Inputs → Asset
register** (rows 16–65). This is slow and error-prone. v3.3 adds a way to map
a CLASS export into the register with minimal manual work, while keeping the
columns CLASS cannot supply under deliberate human control.

## Decisions (locked with user, 2026-06-01)

1. **Integration shape:** an in-workbook **"CLASS Import" tab**. The model is a
   protected, Python-built `.xlsx` that accountants open in Excel — they do not
   run Python — so the mapping must live in the workbook itself.
2. **Data flow:** **copy-paste staging block**. The Import tab filters and maps;
   the user Paste-Special-Values the result into the register. The existing
   register (unlocked inputs, col H formula, sheet protection, sample data) is
   **not changed**.
3. **Columns CLASS cannot supply** (Projected sale proceeds, MV at 30 Jun 2026,
   Held > 12 months?): **left blank** for deliberate user entry — no hidden
   assumptions. The model will not compute CGT for an asset until the user fills
   these, which is the intended, conservative behaviour.
4. **Row filtering:** **blacklist**. Drop cash and the realised-CGT line; let
   everything else through (so a new/unrecognised asset class is never silently
   lost).
5. **Drop handling:** **blank in place** (no row compaction). The mapped block is
   1:1 with the paste zone; dropped rows emit empty cells. Numerically harmless
   (totals SUM ignore blanks; existing CF already guards on non-blank). Avoids
   array/INDEX compaction plumbing.
6. **Cost-base basis (input requirement):** the model's "Original cost base"
   feeds the CGT math, so it must be the **CGT tax cost base** — not accounting
   book cost. CLASS can generate the Investment Summary Report on either an
   *Accounting Cost Base* or a *Tax Cost Base* basis; the layout is identical
   either way, so **the CSV does not self-identify which basis it was run on.**
   The import therefore **requires the Tax Cost Base version**, enforced by
   workflow instruction + an on-tab acknowledgement banner (we cannot detect or
   convert the basis from the data alone). If accounting-cost-base data is pasted
   by mistake, cost bases for trust/ETF/managed-fund holdings (tax-deferred &
   AMIT adjustments) will be overstated → gains/tax understated.
7. **Negative tax cost base:** **pass through + flag.** A negative tax cost base
   (cumulative tax-deferred / return-of-capital distributions have pushed the
   base below nil — CGT-event-E4 territory) is mapped verbatim into column C and
   **visibly flagged red** in the mapped block for manual review. The import
   stays a faithful mirror of CLASS; the preparer decides any E4 / cost-base-reset
   treatment. No silent flooring or omission. (Observed in the sample: GOOGL =
   −$1,772.96.)

## The CLASS export (observed shape)

18 columns, header in row 1. Positions used by the import (1-indexed in the
source CSV; the letter is where they land if the data is pasted starting at
column A of the paste zone):

| # | CLASS column | Used? | Paste-zone col |
|---|---|---|---|
| 2 | Security Code | yes → Asset code | B |
| 3 | Holding Account Name | yes → Asset name | C |
| 8 | G/L Class | yes → filter | H |
| 12 | Total Cost | yes → Original cost base | L |
| 13 | Market Value | yes → Current market value | M |

All other columns are ignored. Foreign holdings are **already AUD-converted** in
Total Cost / Market Value (verified against the sample: e.g. Bank of China
Total Cost = $18,415.46 AUD), so no FX handling is required.

## Column mapping (CLASS → register)

| Register col | Field | Source |
|---|---|---|
| A | Asset code | CLASS Security Code |
| B | Asset name | CLASS Holding Account Name |
| C | Original cost base | CLASS Total Cost **(Tax Cost Base basis export — see decision 6)** |
| D | Current market value (today) | CLASS Market Value |
| E | MV at 30 Jun 2026 | blank (user) |
| F | Valuation source / date | blank (user) |
| G | Projected sale proceeds | blank (user) |
| H | Projected gain/loss | register's own formula (= G − C) |
| I | Held > 12 months? | blank (user) |

## Exclusion rule (blacklist)

A paste-zone row is **dropped** from the mapped block when **any** of:

- `G/L Class` contains the text "Cash" — catches *Cash At Bank* and
  *Foreign Cash At Bank*;
- `Security Code = "REASEDCGT"` — the Realised-CGT carried line (negative cost,
  no units, not a holding);
- the row is blank (Security Code empty).

Everything else passes. Unrecognised asset classes therefore come through rather
than being silently excluded (the chosen safety property).

In the sample export this drops 3 of 36 data rows (Macquarie CMA, USD Account,
Realised CGT), leaving 33 assets.

## Tab layout — "CLASS Import"

A 5th worksheet, inserted **after Inputs** (tab order: Inputs, CLASS Import,
Analyser, Comparison, Notes). Staff staging tab, not client-facing.

1. **Title + instructions band** — the workflow, plus a prominent
   **basis-acknowledgement banner**: *"This import requires the Investment
   Summary Report generated on a **TAX COST BASE** basis. An accounting-cost-base
   export will silently overstate cost bases for trusts/ETFs/managed funds."*
   1. In CLASS, generate the *Investment Summary Report* on the **Tax Cost Base**
      basis (not Accounting Cost Base) and export to CSV.
   2. Select the data rows (everything **below** the header) and copy.
   3. Paste into the green paste zone at the marked top-left cell.
   4. Copy the mapped block **columns A:G only** (NOT col H), then
      Paste-Special → **Values** into **Inputs!A16**. Stopping at G protects the
      register's Projected gain/loss **formula in col H** (`= G − C`); col H is
      also a locked cell, so an accidental A:I paste is blocked by sheet
      protection rather than silently wiping the formula. Pasting Values (not a
      normal paste) preserves the register's formatting and the Held dropdown.
   5. Fill the blank columns by hand, per asset: MV at 30 Jun (E), Projected sale
      proceeds (G), Held > 12 months (I) — and resolve any negative-cost-base flag.

   Mechanics note: the mapped block populates only A–D (code / name / tax cost
   base / current MV); E/F/G are emitted blank so the A:G paste also *clears* any
   stale sample data in those columns of the overwritten rows. Col H (formula) and
   col I (held) are left to the register.
2. **Paste zone** — 50 rows (matches register capacity), columns in CLASS's
   native order so a straight paste lands correctly. Styled as input/green.
3. **Mapped block** — formula-driven, register-shaped (cols A..I), 1:1 with the
   paste zone. Each row: if the source row is dropped by the blacklist, emit
   blanks; otherwise emit the mapped fields (E/F/G/I left blank by design). A
   **negative cost base** (col C < 0) is tinted red with a review note via
   conditional formatting (decision 7).
4. **Capacity guard** — a visible warning if the paste zone has more than 50
   non-blank rows (register only holds 50).
5. **Sample data** — ships pre-loaded with the DEMO_SMSF_2020 **tax-cost-base**
   export (the 2026-06-09 file), matching the workbook's existing "sample data
   preloaded" convention, so staff see a worked example — including the GOOGL
   negative-cost-base edge case.

## Files & wiring

- **New:** `src/div296/tabs/class_import.py` — `build(wb)`, layout constants as
  the single source of truth (mirroring the `inputs.py` constant pattern).
- **Edit:** `src/div296/build.py` — `from div296.tabs import ... class_import`
  and call `class_import.build(wb)` after `inputs.build(wb)`. Footer loop already
  covers all sheets. Update the module docstring's "4 tabs" → "5 tabs".
- **Edit:** `src/div296/__init__.py` and `pyproject.toml` — version → `3.3.0`.
- **Edit:** `tests/test_calcs.py` — version-sanity assertion → `3.3.0`.
- **Edit:** `CONTEXT.md` — glossary entries for the CLASS Import tab, the
  column-mapping table, and the exclusion rule (CONTEXT.md is canonical for
  labels/terminology).
- **Edit:** `README.md` — "What's new in v3.3.0" section; previous release moves
  to "Previous releases".

## Testing

- **Structural** (`tests/test_workbook.py`, fast): "CLASS Import" sheet exists;
  it is positioned after Inputs; instruction band present; paste zone is 50 rows;
  mapped block headers match the 9 register columns; mapped-block formulas
  reference the correct paste-zone columns (B/C/L/M); exclusion formula text
  present; sample data present in the paste zone.
- **Behavioural spot-check:** with the sample (tax-cost-base) data loaded, assert
  the mapped block drops exactly the 3 known non-CGT rows, maps the cost/value of
  at least one known asset, and that the negative-cost-base conditional format is
  present on the mapped-block C column (GOOGL = −$1,772.96 is the fixture)
  (formula-text assertion, since cells aren't live without recalc).
- **Recalc gate:** `python -m div296.build` (the existing `formulas` recalc) must
  stay clean — no new `#VALUE!`/`#REF!` from the import formulas. Known caveat:
  recalc can hit the documented `MemoryError` fallback; an open-in-Excel check of
  the new tab is advised before merge.

## Versioning / acceptance

- v3.3.0 is **additive**: no existing calc, layout constant, or downstream tab
  changes. For a fund that does not use the import, output is **byte-equivalent**
  to v3.2.1 (the §12 acceptance numbers are unchanged).
- Per project convention: PR + green CI + manual user merge; no push without the
  `push this branch now` sentinel.

## Out of scope (explicit YAGNI)

- No row compaction / FILTER (decision 5).
- No auto-derivation of held > 12 months, proceeds, or 30-Jun value (decision 3).
- No live formula links into the register (decision 2).
- No whitelist of asset classes (decision 4).
- No multi-fund / multi-pool handling — one fund's export at a time.
- Prior-year carry-forward loss support remains a separate, already-filed v3.x
  feature, untouched here.
