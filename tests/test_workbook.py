"""Integration tests: build the workbook, optionally recalc, read it back.

v3.0 — control panel removed (toggles always-on internally per Bill-correct
calc); Inputs Section 2 now shows two-band proportion display (cols D/E);
Analyser fund summary restructured side-by-side with per-member breakdown
and signed Difference column.

Row numbers shifted:
- Inputs: Members section moved up by 4 (now rows 6-12; was 10-16); Asset
  register starts at row 16 (was 20).
- Analyser: Lever mirror rows 4-7 deleted; everything shifted up by ~4.
  Fund summary band at row 6 (was 11); per-asset first row at 17 (was 21).
"""

from pathlib import Path

import pytest
from openpyxl import load_workbook

from div296 import named_ranges as nr
from div296.build import build_workbook
from div296.tabs.comparison import (
    DATA_FIRST_ROW as CMP_DATA_FIRST_ROW,
    DATA_LAST_ROW as CMP_DATA_LAST_ROW,
    DATA_OVERFLOW_NOTE_ROW as CMP_OVERFLOW_ROW,
)


EXPECTED_TABS = ["Inputs", "Analyser", "Comparison", "Notes"]


def test_workbook_has_four_tabs_in_spec_order(tmp_path: Path):
    out = tmp_path / "out.xlsx"
    wb = build_workbook()
    wb.save(out)
    reopened = load_workbook(out)
    assert reopened.sheetnames == EXPECTED_TABS


def test_every_named_range_resolves(tmp_path: Path):
    """v3.0: ALL_NAMES no longer includes the 3 toggle names."""
    out = tmp_path / "out.xlsx"
    wb = build_workbook()
    wb.save(out)
    reopened = load_workbook(out)

    missing = [name for name in nr.ALL_NAMES if name not in reopened.defined_names]
    assert not missing, f"Named ranges not found in workbook: {missing}"

    for name in nr.ALL_NAMES:
        dn = reopened.defined_names[name]
        destinations = list(dn.destinations)
        assert destinations, f"Named range {name!r} has no destinations"
        for sheet_name, ref in destinations:
            assert sheet_name in reopened.sheetnames, f"{name} points at missing sheet {sheet_name}"
            ws = reopened[sheet_name]
            cell = ws[ref.replace("$", "")]
            assert cell is not None


def test_v3_toggle_named_ranges_removed(tmp_path: Path):
    """v3.0: reset_on, tier10_on, discount_on must NOT exist."""
    out = tmp_path / "out.xlsx"
    wb = build_workbook()
    wb.save(out)
    reopened = load_workbook(out)
    for removed in ("reset_on", "tier10_on", "discount_on"):
        assert removed not in reopened.defined_names, (
            f"v3.0 should have removed named range {removed!r}"
        )


# --- Inputs tab -----------------------------------------------------------

def test_inputs_sample_data_preloaded(tmp_path: Path):
    """v3.0: Asset register first row shifted from row 20 to row 16
    (control panel rows 4-8 removed)."""
    out = tmp_path / "out.xlsx"
    wb = build_workbook()
    wb.save(out)
    ws = load_workbook(out)["Inputs"]

    # Row 16: Commercial property (was row 20 in v2.6)
    assert ws["A16"].value == "P1"
    assert ws["B16"].value == "Commercial property"
    assert ws["C16"].value == 800_000             # Original cost base
    assert ws["E16"].value == 2_400_000           # MV at 30 Jun 2026
    assert ws["G16"].value == 2_600_000           # Projected sale proceeds
    assert ws["I16"].value == "Yes"               # Held>12m

    # Row 17: Listed shares parcel
    assert ws["A17"].value == "S1"
    assert ws["C17"].value == 300_000
    assert ws["E17"].value == 520_000

    # Row 18: Loss-making holding
    assert ws["A18"].value == "L1"
    assert ws["C18"].value == 500_000
    assert ws["E18"].value == 100_000

    # Projected gain/loss formula in col H references new row numbers.
    pg_formula = ws["H16"].value
    assert pg_formula and "G16" in pg_formula and "C16" in pg_formula

    # v3.0: Sole member on row 7 (was row 11): TSB $12m.
    assert ws["B7"].value == 12_000_000
    c7 = ws["C7"].value
    assert isinstance(c7, str) and c7.startswith("="), (
        f"C7 (Split %) should be a formula in v3.0, got {c7!r}"
    )
    assert "SUM($B$7:$B$10)" in c7 and "B7/" in c7


def test_inputs_section_2_band_columns(tmp_path: Path):
    """v3.0 NEW: cols D and E in Members section are auto-derived band1 / band2."""
    out = tmp_path / "out.xlsx"
    wb = build_workbook()
    wb.save(out)
    ws = load_workbook(out)["Inputs"]

    # Header row (row 6 in v3.0)
    assert "$3m" in str(ws["D6"].value) and "$10m" in str(ws["D6"].value), (
        f"Col D header should mention $3m-$10m band, got {ws['D6'].value!r}"
    )
    assert "$10m" in str(ws["E6"].value), (
        f"Col E header should mention above $10m, got {ws['E6'].value!r}"
    )

    # Member 1 (row 7) — band1 formula uses MIN(TSB, threshold_2) - threshold_1
    d7 = ws["D7"].value
    assert d7 and d7.startswith("=")
    assert "MIN" in d7 and "threshold_2" in d7 and "threshold_1" in d7

    # Member 1 (row 7) — band2 formula uses (TSB - threshold_2)
    e7 = ws["E7"].value
    assert e7 and e7.startswith("=")
    assert "threshold_2" in e7
    # band2 should NOT reference threshold_1 — only the >$10m slice.
    assert "threshold_1" not in e7


def test_inputs_control_panel_removed(tmp_path: Path):
    """v3.0: control panel section deleted — no toggles at B5/B6/B7."""
    out = tmp_path / "out.xlsx"
    wb = build_workbook()
    wb.save(out)
    ws = load_workbook(out)["Inputs"]

    # In v3.0 the band at row 5 should read "1. Members" — not "Control panel".
    a5 = str(ws["A5"].value or "")
    assert "Control panel" not in a5, f"Control panel should be gone, got A5={a5!r}"
    assert "Members" in a5, f"Row 5 should be Members band, got A5={a5!r}"

    # Previous control-panel cells must not hold toggle defaults anymore.
    # B5 is now blank (band only). B6 is Members header. B7 is Member 1 TSB.
    assert ws["B7"].value == 12_000_000, (
        f"B7 should be Member 1 TSB in v3.0, got {ws['B7'].value!r}"
    )


def test_sample_data_badge_present(tmp_path: Path):
    """Row 2 badge tells staff to overwrite sample data."""
    out = tmp_path / "out.xlsx"
    wb = build_workbook()
    wb.save(out)
    ws = load_workbook(out)["Inputs"]
    badge = ws["A2"].value
    assert badge is not None
    assert "Sample data" in badge
    assert "overwrite" in badge.lower()


def test_all_tabs_protected(tmp_path: Path):
    out = tmp_path / "out.xlsx"
    wb = build_workbook()
    wb.save(out)
    wb_re = load_workbook(out)
    for tab in ("Inputs", "Analyser", "Notes", "Comparison"):
        assert wb_re[tab].protection.sheet is True, f"{tab} is not protected"


def test_all_tabs_allow_column_resize_under_protection(tmp_path: Path):
    out = tmp_path / "out.xlsx"
    wb = build_workbook()
    wb.save(out)
    wb_re = load_workbook(out)
    for sheet in ("Inputs", "Analyser", "Notes", "Comparison"):
        ws = wb_re[sheet]
        assert ws.protection.sheet is True, f"{sheet} should be protected"
        assert ws.protection.formatColumns in (False, "0", 0), (
            f"{sheet} formatColumns must be unlocked under v2+ protection scope"
        )
        assert ws.protection.formatRows in (False, "0", 0), (
            f"{sheet} formatRows must be unlocked under v2+ protection scope"
        )


def test_input_cells_unlocked(tmp_path: Path):
    """v3.0: Member 1 TSB (B7) and Original cost base (C16) must stay editable."""
    out = tmp_path / "out.xlsx"
    wb = build_workbook()
    wb.save(out)
    ws = load_workbook(out)["Inputs"]
    # v3.0 shifted: Member 1 TSB is at B7; Asset register original cost base at C16.
    assert ws["B7"].protection.locked is False
    assert ws["C16"].protection.locked is False


# --- Held>12m paste-in normalisation (v3.1.2) ---------------------------
#
# The Inputs!I dropdown only validates direct typing; pastes from another
# sheet or CSV import bypass it entirely. v3.1.2 added a hidden Inputs!J
# helper column that runs TRIM+UPPER on I and emits a clean "Yes"/"No",
# and redirected every downstream formula to read J instead of I.
# These tests pin both the wiring (formulas reference J) and the
# behaviour (paste-in "Yes " / "yes" still produce the discounted gain).


def test_inputs_j_column_is_hidden_normaliser(tmp_path: Path):
    """Inputs!J16:J65 = =IF(I=...,..,IF(TRIM(UPPER(I))="YES","Yes","No"));
    column is hidden so the user never sees the helper."""
    out = tmp_path / "out.xlsx"
    wb = build_workbook()
    wb.save(out)
    ws = load_workbook(out)["Inputs"]

    assert ws.column_dimensions["J"].hidden, "Inputs!J must be hidden"

    for row in (16, 40, 65):
        f = ws[f"J{row}"].value
        assert isinstance(f, str) and f.startswith("="), (
            f"Inputs!J{row} should be a formula, got {f!r}"
        )
        assert "TRIM" in f and "UPPER" in f, (
            f"Inputs!J{row} should normalise via TRIM+UPPER, got {f!r}"
        )
        assert f'I{row}' in f, (
            f"Inputs!J{row} should reference its own I{row} source, got {f!r}"
        )


def test_downstream_formulas_read_held_from_j_not_i(tmp_path: Path):
    """Every formula that classifies a gain as discountable must read the
    NORMALISED held flag (Inputs!J), not the raw user input (Inputs!I)."""
    out = tmp_path / "out.xlsx"
    wb = build_workbook()
    wb.save(out)
    wb_re = load_workbook(out)

    analyser = wb_re["Analyser"]
    # Per-asset Ord taxable gain (E17, first register row).
    e17 = analyser["E17"].value
    assert "'Inputs'!J16" in e17, (
        f"Analyser E17 must read held from Inputs!J16, got {e17!r}"
    )
    assert "'Inputs'!I16" not in e17, (
        f"Analyser E17 must not read raw Inputs!I16, got {e17!r}"
    )
    # Per-asset Div 296 adjusted gain (H17).
    h17 = analyser["H17"].value
    assert "'Inputs'!J16" in h17 and "'Inputs'!I16" not in h17, h17
    # Reconciliation helper M70 — SUMIFS over the held range.
    m70 = analyser["M70"].value
    assert "'Inputs'!J16:J65" in m70, (
        f"Recon helper M70 must aggregate over Inputs!J16:J65, got {m70!r}"
    )
    assert "'Inputs'!I16:I65" not in m70, m70

    # Comparison per-register helper grid (cols N/O — div296 adj gains).
    comparison = wb_re["Comparison"]
    # Find the per-register grid row aligned to Inputs row 16; the grid uses
    # the same row numbers as Inputs (n_first = REGISTER_FIRST_DATA_ROW).
    n16 = comparison["N16"].value
    assert "'Inputs'!J16" in n16 and "'Inputs'!I16" not in n16, n16


@pytest.mark.slow
def test_paste_in_dirty_held_values_are_normalised(tmp_path: Path):
    """Behavioural test: write paste-style dirty values ('Yes ', 'yes',
    ' YES ') into Inputs!I, recalc, and check that Analyser per-asset
    E (Ord taxable gain) and H (Div 296 adj gain) both apply the 1/3
    CGT discount as if the cell had said "Yes" exactly."""
    formulas = pytest.importorskip(  # noqa: F811
        "formulas",
        reason="`formulas` package required for paste-in normalisation test",
    )
    import numpy as np

    out = tmp_path / "out.xlsx"
    wb = build_workbook()

    # Simulate paste-in: trailing space (row 16), lowercase (row 17),
    # surrounding whitespace + uppercase (row 18). All three sample rows.
    ws_in = wb["Inputs"]
    ws_in["I16"] = "Yes "
    ws_in["I17"] = "yes"
    ws_in["I18"] = " YES "
    wb.save(out)

    try:
        sol = formulas.ExcelModel().loads(str(out)).finish().calculate()
    except MemoryError:
        pytest.skip("`formulas` OOM'd on v3.1.2 workbook recalc")

    file_token = f"[{out.name}]"

    def at(sheet: str, cell: str):
        key = f"'{file_token}{sheet.upper()}'!{cell}"
        v = sol[key].value
        if isinstance(v, np.ndarray):
            v = v.flat[0]
        return v

    # Inputs!J should normalise every variant to "Yes".
    for r in (16, 17, 18):
        j = at("Inputs", f"J{r}")
        assert j == "Yes", (
            f"Inputs!J{r} must normalise paste-in to 'Yes', got {j!r}"
        )

    # Analyser E17 — P1 (orig 800k, proceeds 2.6m) → raw gain 1.8m → with
    # discount applied: 1.8m * (1 - 1/3) = 1,200,000. If normalisation
    # didn't happen, the formula would skip the discount branch and return
    # the raw 1.8m.
    e17 = float(at("Analyser", "E17"))
    assert abs(e17 - 1_200_000) < 1, (
        f"E17 should be discounted long-held gain 1.2m, got {e17}"
    )

    # Analyser H17 — Div 296 adj gain, cost base = MV 2.4m, proceeds 2.6m
    # → raw 200k → discounted ≈ 133,333.33.
    h17 = float(at("Analyser", "H17"))
    assert abs(h17 - 133_333.33) < 1, (
        f"H17 should be discounted Div 296 adj gain ≈ 133,333, got {h17}"
    )

    # Analyser B71 (Fund Ord CGT) routes through the recon helpers
    # M70/N70/O70 which read Inputs!J via SUMIFS. M70 = discount-eligible
    # gains > 0. With all three sample rows held>12m (paste-dirty), the
    # discount-eligible gains are P1 (1.8m) + S1 (0.3m) = 2,100,000.
    # L1 has a projected loss (-300k) → in O70, not M70.
    # If normalisation didn't happen, "Yes ", "yes", " YES " would all fail
    # the SUMIFS exact-match against "Yes" and M70 would collapse to 0.
    m70 = float(at("Analyser", "M70"))
    assert abs(m70 - 2_100_000) < 1, (
        f"Recon helper M70 should aggregate paste-normalised long-held "
        f"gains to 2,100,000, got {m70}"
    )


# --- Analyser tab ---------------------------------------------------------

def _analyser(tmp_path: Path):
    out = tmp_path / "out.xlsx"
    wb = build_workbook()
    wb.save(out)
    return load_workbook(out)["Analyser"]


class TestAnalyser:
    def test_no_lever_mirror_rows(self, tmp_path: Path):
        """v3.0: lever mirror at rows 5-7 (and band at row 4) is deleted."""
        ws = _analyser(tmp_path)
        # Rows 5-7 should NOT contain the v2.x lever-mirror Inputs!B refs.
        for row in (5, 6, 7):
            b = ws.cell(row=row, column=2).value
            assert b is None or not (isinstance(b, str) and b.startswith("='Inputs'!B")), (
                f"v3.0 should have no lever mirror at B{row}, got {b!r}"
            )

    def test_state_strip_parameters_in_effect(self, tmp_path: Path):
        """v3.0: state strip row 2 is the 'Parameters in effect' read-only line.
        References named ranges for rates / thresholds / discount, NOT toggles."""
        ws = _analyser(tmp_path)
        cell = ws["A2"]
        v = cell.value
        assert v and isinstance(v, str) and v.startswith("=")
        assert "Parameters in effect" in v
        # Must reference these named ranges live.
        for nm in ("rate_tier1", "rate_tier2", "threshold_1", "threshold_2",
                   "discount_rate", "fund_cgt_rate"):
            assert nm in v, f"state strip should reference {nm}"
        # Must NOT reference any removed toggle.
        for removed in ("reset_on", "tier10_on", "discount_on"):
            assert removed not in v, (
                f"state strip should not reference removed named range {removed}: {v!r}"
            )

    def test_fund_summary_side_by_side_headers(self, tmp_path: Path):
        """v3.0: fund-summary header row (row 7) has scenario column headers."""
        ws = _analyser(tmp_path)
        # Col C, D, E hold scenario column headers.
        assert "no reset" in str(ws["C7"].value).lower()
        assert "elected" in str(ws["D7"].value).lower()
        assert "difference" in str(ws["E7"].value).lower()

    def test_fund_earnings_two_scenarios(self, tmp_path: Path):
        """v3.1: Fund Div 296 earnings nets gains and losses intra-year,
        floored at zero — formula is MAX(0, SUM(...)). v3.0 used
        SUMIF(...,">0") which floored per-asset before summing.
        - C8 (no reset) nets helper col L (per-asset no-reset gain).
        - D8 (elected) nets col H (per-asset adjusted gain, elected scenario).
        - E8 (difference) = D8 - C8."""
        ws = _analyser(tmp_path)
        c8 = ws["C8"].value
        d8 = ws["D8"].value
        e8 = ws["E8"].value
        assert c8 == '=MAX(0, SUM(L17:L66))', f"C8 wrong: {c8!r}"
        assert d8 == '=MAX(0, SUM(H17:H66))', f"D8 wrong: {d8!r}"
        assert e8 == "=D8-C8", f"E8 (diff) should be D8-C8, got {e8!r}"

    def test_headline_row_sums_member_taxes_per_scenario(self, tmp_path: Path):
        """v3.0: Headline at row 13, three scenario columns each summing rows 9-12."""
        ws = _analyser(tmp_path)
        assert ws["C13"].value == "=SUM(C9:C12)"
        assert ws["D13"].value == "=SUM(D9:D12)"
        assert ws["E13"].value == "=SUM(E9:E12)"

    def test_member_tax_formula_reads_band1_band2_from_inputs(self, tmp_path: Path):
        """v3.0: per-member tax (row 9 Member 1) reads Inputs!D7 (band1) and
        Inputs!E7 (band2) directly — single source of truth with Inputs."""
        ws = _analyser(tmp_path)
        # No-reset scenario at C9
        c9 = ws["C9"].value
        assert c9 and isinstance(c9, str)
        for ref in ("'Inputs'!B7", "'Inputs'!C7", "'Inputs'!D7", "'Inputs'!E7"):
            assert ref in c9, f"member 1 no-reset formula missing {ref}: {c9!r}"
        assert "rate_tier1" in c9 and "rate_tier2" in c9
        # Must not reference removed toggle.
        assert "tier10_on" not in c9, f"v3.0 formula should not reference tier10_on: {c9!r}"

        # Elected scenario at D9 — same formula, different earnings cell.
        d9 = ws["D9"].value
        assert d9 and isinstance(d9, str)
        assert "rate_tier1" in d9 and "rate_tier2" in d9

        # Difference at E9
        assert ws["E9"].value == "=D9-C9"

    def test_per_asset_first_row_columns(self, tmp_path: Path):
        """v3.0: per-asset detail starts at row 17 (was 21). Cost base hardcoded
        to mv (no reset_on IF). Discount hardcoded to held>12mo (no discount_on)."""
        ws = _analyser(tmp_path)
        # Row 17 first data row; visible cols 2..10.
        a = ws.cell(row=17, column=2).value   # Asset display
        c = ws.cell(row=17, column=3).value   # Proceeds
        d = ws.cell(row=17, column=4).value   # Original cost base
        e = ws.cell(row=17, column=5).value   # Ordinary taxable gain
        f = ws.cell(row=17, column=6).value   # Ordinary CGT
        g = ws.cell(row=17, column=7).value   # Div 296 cost base
        h = ws.cell(row=17, column=8).value   # Div 296 adjusted gain
        i = ws.cell(row=17, column=9).value   # Div 296 tax (pro-rata of headline)
        j = ws.cell(row=17, column=10).value  # Reset impact

        assert "'Inputs'!A16" in a and "'Inputs'!B16" in a
        assert "'Inputs'!G16" in c    # Proceeds
        assert "'Inputs'!C16" in d    # Original CB
        assert "discount_rate" in e and "'Inputs'!G16" in e
        assert "fund_cgt_rate" in f and "E17" in f
        # Col G — Div 296 cost base. v3.0: no IF(reset_on), just mv.
        assert "'Inputs'!E16" in g
        assert "reset_on" not in g, f"v3.0 col G must not reference reset_on: {g!r}"
        # Col H — Div 296 adj gain, no reset_on
        assert "discount_rate" in h
        assert "reset_on" not in h
        # Col H must also not reference discount_on toggle.
        assert "discount_on" not in h
        # Col I — Div 296 tax (pro-rata of headline). Headline now at $D$13 (elected).
        assert "$D$13" in i and 'SUMIF(H17:H66,">0")' in i
        # Col J — Reset impact = K - L
        assert "K17" in j and "L17" in j

    def test_helper_columns_hidden(self, tmp_path: Path):
        ws = _analyser(tmp_path)
        assert ws.column_dimensions["K"].hidden
        assert ws.column_dimensions["L"].hidden

    def test_totals_row_v3(self, tmp_path: Path):
        """v3.1.1: totals row at 67. The three info-only cols (E, F, H) all
        show "(see fund total)" placeholder text — none of them sum, because
        per-asset post-discount values don't aggregate meaningfully once
        fund-level loss netting is in play (s102-5). Authoritative fund
        figures live in the Reconciliation panel."""
        ws = _analyser(tmp_path)
        assert ws["C67"].value == "=SUM(C17:C66)"
        # v3.1.1: info-only columns are not summed.
        assert ws["E67"].value == "(see fund total)"
        assert ws["F67"].value == "(see fund total)"
        assert ws["H67"].value == "(see fund total)"
        assert ws["I67"].value == "=SUM(I17:I66)"

    def test_f_info_footnote_present(self, tmp_path: Path):
        """v3.1: row 68 carries a footnote explaining col F is info only."""
        ws = _analyser(tmp_path)
        footnote = ws["A68"].value
        assert footnote is not None
        assert "info only" in footnote
        assert "fund" in footnote.lower()

    def test_per_asset_ord_cgt_col_header_grey_label(self, tmp_path: Path):
        """v3.1: col F header renamed to 'Per-asset Ord CGT (info only)'."""
        ws = _analyser(tmp_path)
        f16 = ws["F16"].value
        assert "info only" in f16

    def test_per_asset_ord_cgt_formula_shows_dash_for_losses(self, tmp_path: Path):
        """v3.1: col F formula returns '—' for loss rows (E<=0)."""
        ws = _analyser(tmp_path)
        f17 = ws["F17"].value
        assert "—" in f17 and "IF(E17<=0" in f17

    def test_reconciliation_panel_v31(self, tmp_path: Path):
        """v3.1: recon band at row 70. Fund Ordinary CGT uses the s102-5
        netted formula (not =SUM(F:F)). Carry-forward losses are net unused
        gross losses at the fund level (not a per-asset sum)."""
        ws = _analyser(tmp_path)
        # Band shifted down by 1 row due to new F-info footnote at row 68.
        assert ws["A70"].value == "Reconciliation"
        assert ws["A71"].value == "Fund Ordinary CGT (after intra-year netting)"
        b71 = ws["B71"].value
        # New formula references hidden helpers at M70/N70/O70 and the
        # discount_rate / fund_cgt_rate named ranges.
        assert b71.startswith("=")
        assert "M70" in b71 and "N70" in b71 and "O70" in b71
        assert "discount_rate" in b71 and "fund_cgt_rate" in b71

        assert ws["A72"].value == "Div 296 tax payable (elected-reset headline)"
        assert ws["B72"].value == "=D13"

        assert ws["A73"].value == "Capital losses carried forward"
        # v3.1: net unused gross loss, computed from the same 3 helpers.
        b73 = ws["B73"].value
        assert b73 == "=MAX(0, O70 - (M70 + N70))"

    def test_recon_helpers_hidden(self, tmp_path: Path):
        """v3.1: helper cells (M70, N70, O70) live in hidden cols M/N/O."""
        ws = _analyser(tmp_path)
        for col in ("M", "N", "O"):
            assert ws.column_dimensions[col].hidden, f"col {col} not hidden"
        # Helpers must be formulas (not blank).
        for cell in ("M70", "N70", "O70"):
            v = ws[cell].value
            assert v and v.startswith("="), f"{cell} empty/non-formula: {v!r}"

    def test_recon_helpers_use_sumifs_not_sumproduct(self, tmp_path: Path):
        """v3.1.1: M70/N70/O70 use SUMIFS/SUMIF, NOT SUMPRODUCT.

        The briefly-shipped SUMPRODUCT pattern `SUMPRODUCT(ISNUMBER(H)*(H>0)*...*H)`
        errored with #VALUE! because the trailing `*H` multiplied empty-row text
        values ("" from the Inputs!H IF formula) and the ISNUMBER guard doesn't
        short-circuit. SUMIFS handles mixed-type sum_range natively.
        """
        ws = _analyser(tmp_path)
        m70 = ws["M70"].value
        n70 = ws["N70"].value
        o70 = ws["O70"].value
        # Must use SUMIFS (or SUMIF for the loss helper) — not SUMPRODUCT.
        assert m70.startswith("=SUMIFS("), f"M70 must use SUMIFS: {m70!r}"
        assert n70.startswith("=SUMIFS("), f"N70 must use SUMIFS: {n70!r}"
        assert o70.startswith("=-SUMIF("), f"O70 must use -SUMIF: {o70!r}"
        for cell, v in (("M70", m70), ("N70", n70), ("O70", o70)):
            assert "SUMPRODUCT" not in v, (
                f"{cell} should NOT use SUMPRODUCT (it errors on empty-row "
                f"text values): {v!r}"
            )

    def test_trap_conditional_formatting_applied(self, tmp_path: Path):
        """v3.0: trap CF range now A17:J66 (was A21:J70)."""
        ws = _analyser(tmp_path)
        cf_rules = list(ws.conditional_formatting._cf_rules.items())
        assert cf_rules, "no conditional formatting rules on Analyser"
        ranges = [str(r[0]) for r in cf_rules]
        assert any("A17" in r and "J66" in r for r in ranges)

    def test_row_number_column(self, tmp_path: Path):
        """v3.0: col A row-num 1..50 spans rows 17-66; totals at row 67."""
        ws = _analyser(tmp_path)
        assert ws["A16"].value == "#"        # header row
        assert ws["A17"].value == 1          # first data row
        assert ws["A66"].value == 50         # last data row
        assert ws["A67"].value == "Total"    # totals row

    def test_print_titles_repeat_header(self, tmp_path: Path):
        """v3.0: per-asset header repeats on every printed page (row 16)."""
        ws = _analyser(tmp_path)
        assert ws.print_title_rows in ("16:16", "$16:$16")

    def test_no_freeze_panes(self, tmp_path: Path):
        ws = _analyser(tmp_path)
        assert ws.freeze_panes is None

    def test_trap_cf_formula_uses_relative_rows(self, tmp_path: Path):
        """The CF formula must NOT have $ before row numbers (must adjust per row)."""
        import re
        ws = _analyser(tmp_path)
        for _rng, rules in ws.conditional_formatting._cf_rules.items():
            for rule in rules:
                if rule.formula:
                    for f in rule.formula:
                        bad = re.findall(r"\$\d+", f)
                        assert not bad, (
                            f"CF formula has absolute-row reference(s) {bad}: {f!r}"
                        )


# --- Comparison tab -------------------------------------------------------

def _comparison(tmp_path: Path):
    out = tmp_path / "out.xlsx"
    wb = build_workbook()
    wb.save(out)
    return load_workbook(out)["Comparison"]


class TestComparison:
    def test_watermark_banner_and_print_header(self, tmp_path: Path):
        ws = _comparison(tmp_path)
        assert ws["A1"].value == "ILLUSTRATIVE — NOT ADVICE"
        assert "ILLUSTRATIVE" in ws.oddHeader.center.text

    def test_context_strip_present(self, tmp_path: Path):
        """v3.0: Members & TSB strip still uses MEMBERS_FIRST_DATA_ROW (now 7)
        for the SUM reference."""
        ws = _comparison(tmp_path)
        assert "Members & TSB" in str(ws["A12"].value)
        assert ws["A13"].value == "Members"
        assert ws["B13"].value == "Total Super Balance"
        assert ws["A14"].value == "Member 1"
        assert ws["A15"].value == "Member 2"
        assert ws["A16"].value == "Member 3"
        assert ws["A17"].value == "Member 4"
        assert ws["A18"].value == "Total"
        # v3.0: members on Inputs now at rows 7-10 (was 11-14).
        assert "SUM('Inputs'!B7:B10)" in ws["B18"].value
        for col in ("D", "E", "F", "G", "H", "I", "J", "K"):
            for row in range(14, 19):
                val = ws[f"{col}{row}"].value
                assert val is None, (
                    f"{col}{row} should be empty (right-side tiles dropped); got {val!r}"
                )

    def test_metric_cards_present(self, tmp_path: Path):
        ws = _comparison(tmp_path)
        assert ws["A21"].value == "If no Div 296 CostBase Reset (default)"
        assert ws["E21"].value == "If elected to reset Div 296 CostBase Reset"
        assert ws["I21"].value == "Difference (Net Div 296 Tax)"
        assert "L$6" in ws["A22"].value or "L6" in ws["A22"].value
        assert "M$6" in ws["E22"].value or "M6" in ws["E22"].value
        net = ws["I22"].value
        assert ("M" in net and "L" in net and "-" in net), f"unexpected difference formula {net!r}"
        assert net.index("M") < net.index("L")

    def test_subtotals_table(self, tmp_path: Path):
        ws = _comparison(tmp_path)
        assert ws["B26"].value == "If no reset (default)"
        assert ws["C26"].value == "If elected to reset"
        assert ws["D26"].value == "Difference"
        assert ws["A27"].value == "Div 296 earnings"
        assert "Ordinary CGT" in ws["A28"].value and "unchanged" in ws["A28"].value
        assert ws["A29"].value == "Div 296 tax (headline)"
        assert "TOTAL TAX BURDEN" in ws["A30"].value
        assert "Analyser" in ws["B28"].value
        assert ws["B28"].value == ws["C28"].value
        assert ws["B30"].value == "=B28+B29"
        assert ws["C30"].value == "=C28+C29"
        for row in (27, 28, 29, 30):
            assert ws[f"D{row}"].value == f"=C{row}-B{row}"

    def test_panel_a_uses_original_cost_base(self, tmp_path: Path):
        ws = _comparison(tmp_path)
        cb_formula = ws[f"C{CMP_DATA_FIRST_ROW}"].value
        assert "'Inputs'!$C:$C" in cb_formula
        assert "'Inputs'!$E:$E" not in cb_formula

    def test_panel_b_uses_market_value(self, tmp_path: Path):
        ws = _comparison(tmp_path)
        cb_formula = ws[f"H{CMP_DATA_FIRST_ROW}"].value
        assert "'Inputs'!$E:$E" in cb_formula
        assert "'Inputs'!$C:$C" not in cb_formula

    def test_panels_independent_of_master_reset_toggle(self, tmp_path: Path):
        """Neither panel formula may reference reset_on (which no longer exists)."""
        ws = _comparison(tmp_path)
        for row in range(CMP_DATA_FIRST_ROW, CMP_DATA_LAST_ROW + 1):
            for col_letter in ("B", "C", "D", "E", "G", "H", "I", "J"):
                f = ws[f"{col_letter}{row}"].value
                if f and isinstance(f, str):
                    assert "reset_on" not in f, (
                        f"{col_letter}{row} references reset_on (removed in v3.0)"
                    )

    def test_per_member_formulas_have_no_tier_or_discount_toggle(self, tmp_path: Path):
        """v3.0: Comparison's per-member tax formula must not reference removed toggles."""
        ws = _comparison(tmp_path)
        # Helper col L holds the Scenario A per-member tax (rows starting at HELPER_MEMBER_TAX_FIRST_ROW=2).
        for row in range(2, 6):
            for col_letter in ("L", "M"):
                f = ws[f"{col_letter}{row}"].value
                if f and isinstance(f, str):
                    for removed in ("tier10_on", "discount_on"):
                        assert removed not in f, (
                            f"{col_letter}{row} references removed toggle {removed}: {f!r}"
                        )

    def test_delta_column_formula(self, tmp_path: Path):
        ws = _comparison(tmp_path)
        delta = ws[f"K{CMP_DATA_FIRST_ROW}"].value
        assert f"J{CMP_DATA_FIRST_ROW}" in delta and f"E{CMP_DATA_FIRST_ROW}" in delta
        assert "-" in delta
        assert delta.index(f"J{CMP_DATA_FIRST_ROW}") < delta.index(f"E{CMP_DATA_FIRST_ROW}")

    def test_helper_columns_hidden(self, tmp_path: Path):
        ws = _comparison(tmp_path)
        for col_letter in ("L", "M", "N", "O", "P", "Q", "R"):
            assert ws.column_dimensions[col_letter].hidden, f"col {col_letter} should be hidden"
        for col_letter in ("A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K"):
            assert not ws.column_dimensions[col_letter].hidden, f"col {col_letter} must be visible"

    def test_top_10_sorted_rendered(self, tmp_path: Path):
        ws = _comparison(tmp_path)
        first = ws[f"A{CMP_DATA_FIRST_ROW}"].value
        last = ws[f"A{CMP_DATA_LAST_ROW}"].value
        assert first and first.startswith("=")
        assert last and last.startswith("=")
        overflow = ws[f"A{CMP_OVERFLOW_ROW}"].value
        assert overflow and "top 10" in overflow.lower()
        assert "Δ" not in overflow

    def test_per_asset_detail_uses_large_match_sort(self, tmp_path: Path):
        ws = _comparison(tmp_path)
        matched_formula = ws[f"R{CMP_DATA_FIRST_ROW}"].value
        assert "LARGE" in matched_formula and "MATCH" in matched_formula
        assert (f"$R{CMP_DATA_FIRST_ROW}" in ws[f"B{CMP_DATA_FIRST_ROW}"].value
                or f"R{CMP_DATA_FIRST_ROW}" in ws[f"B{CMP_DATA_FIRST_ROW}"].value)

    def test_no_recommendation_language(self, tmp_path: Path):
        ws = _comparison(tmp_path)
        forbidden = ("saved", "saves", "you should", "we recommend", "we suggest")
        for row in ws.iter_rows(values_only=True):
            for v in row:
                if isinstance(v, str):
                    for word in forbidden:
                        assert word.lower() not in v.lower(), (
                            f"Comparison tab contains forbidden recommendation word "
                            f"{word!r} in cell value {v!r}"
                        )

    def test_print_setup_landscape_a4(self, tmp_path: Path):
        ws = _comparison(tmp_path)
        assert ws.page_setup.orientation == ws.ORIENTATION_LANDSCAPE
        assert int(ws.page_setup.paperSize) == int(ws.PAPERSIZE_A4)
        assert int(ws.page_setup.fitToWidth) == 1
        assert ws.print_area is not None and "K" in ws.print_area

    def test_no_chart_v25(self, tmp_path: Path):
        ws = _comparison(tmp_path)
        assert len(ws._charts) == 0


# --- Notes tab ------------------------------------------------------------

def _notes(tmp_path: Path):
    out = tmp_path / "out.xlsx"
    wb = build_workbook()
    wb.save(out)
    return load_workbook(out)["Notes"]


class TestNotes:
    def test_required_caveats_present(self, tmp_path: Path):
        """Every locked caveat from README must appear on the Notes tab."""
        ws = _notes(tmp_path)
        all_text = " ".join(
            str(c.value) for row in ws.iter_rows() for c in row if c.value is not None
        )
        for required in (
            "Illustrative only",
            "Pension phase is NOT modelled",
            # v3.1.1: "Loss-offset divergence" replaced by
            # "Prior-year capital losses are NOT modelled" — v3.1 now
            # implements intra-year netting per s102-5, so the previous
            # divergence caveat is no longer factually correct.
            "Prior-year capital losses are NOT modelled",
            "Reset OFF scenario is realised-only",
            "Wash sale / Part IVA",
            "Transaction costs",
            "actuarial certificate",
            "all-or-nothing and irrevocable",
            "recontribution",
        ):
            assert required in all_text, f"Notes missing required caveat: {required!r}"

    def test_no_recommendation_language(self, tmp_path: Path):
        ws = _notes(tmp_path)
        forbidden = ("you should", "we recommend", "we suggest")
        for row in ws.iter_rows(values_only=True):
            for v in row:
                if isinstance(v, str):
                    for word in forbidden:
                        assert word.lower() not in v.lower(), (
                            f"Notes contains forbidden advice phrase {word!r}: {v!r}"
                        )

    def test_valuation_log_mirrors_register(self, tmp_path: Path):
        """Valuation log has one row per asset, formulas pointing at Inputs."""
        ws = _notes(tmp_path)
        ref_count = 0
        for row in ws.iter_rows():
            for c in row:
                if isinstance(c.value, str) and "'Inputs'!F" in c.value:
                    ref_count += 1
        assert ref_count >= 50, f"Expected ≥50 valuation-source mirrors, got {ref_count}"

    def test_provenance_cells_present(self, tmp_path: Path):
        ws = _notes(tmp_path)
        labels = {
            str(c.value): c.row
            for row in ws.iter_rows()
            for c in row
            if c.column == 1 and isinstance(c.value, str)
        }
        assert "build_version" in labels
        assert "build_date" in labels
        assert "git_short_sha" in labels
        for label in ("build_version", "build_date", "git_short_sha"):
            assert ws.row_dimensions[labels[label]].hidden, (
                f"Provenance row for {label!r} should be hidden"
            )
