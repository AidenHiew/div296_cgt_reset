"""Integration tests: build the workbook, optionally recalc, read it back.

v1.5: layouts changed in Inputs (zones reordered: control panel, members,
asset register, advanced assumptions), Analyser (Div 296 cost base moved
from col D to col F), and Comparison (full redesign — subtotals at top,
metric cards, fund-context strip, 5-col panels + Δ column, 15 visible
data rows). Tests below mirror the new coords.
"""

from pathlib import Path

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


def test_inputs_sample_data_preloaded(tmp_path: Path):
    """v2.3: Quantity col dropped; cols shift D→C (orig CB), F→E (MV today),
    F→E (MV 30 Jun), G→F (val source), H→G (proceeds). New col H = projected G/L."""
    out = tmp_path / "out.xlsx"
    wb = build_workbook()
    wb.save(out)
    ws = load_workbook(out)["Inputs"]

    # Row 20: Commercial property
    assert ws["A20"].value == "P1"
    assert ws["B20"].value == "Commercial property"
    assert ws["C20"].value == 800_000             # Original cost base (was D)
    assert ws["E20"].value == 2_400_000           # MV at 30 Jun 2026 (was F)
    assert ws["G20"].value == 2_600_000           # Projected sale proceeds (was H)
    assert ws["I20"].value == "Yes"               # Held>12m (still col I)

    # Row 21: Listed shares parcel
    assert ws["A21"].value == "S1"
    assert ws["C21"].value == 300_000
    assert ws["E21"].value == 520_000

    # Row 22: Loss-making holding
    assert ws["A22"].value == "L1"
    assert ws["C22"].value == 500_000
    assert ws["E22"].value == 100_000

    # NEW v2.3: Projected gain/loss formula in col H for every row.
    pg_formula = ws["H20"].value
    assert pg_formula and "G20" in pg_formula and "C20" in pg_formula

    # Sole member on row 11: TSB $12m.
    # v2.4 FB-1: Split % is now an auto formula derived from TSB share,
    # not a typed value. With only Member 1 having a TSB, the formula
    # returns 1.0 at runtime, but openpyxl reads back the formula string
    # (data_only=False). Verify the formula shape instead.
    assert ws["B11"].value == 12_000_000
    c11 = ws["C11"].value
    assert isinstance(c11, str) and c11.startswith("="), (
        f"C11 should be a formula in v2.4+, got {c11!r}"
    )
    assert "SUM($B$11:$B$14)" in c11 and "B11/" in c11


def test_control_panel_defaults(tmp_path: Path):
    """v1.7: control panel = 3 levers on rows 5-7 (Manual earnings dropped)."""
    out = tmp_path / "out.xlsx"
    wb = build_workbook()
    wb.save(out)
    ws = load_workbook(out)["Inputs"]

    assert ws["B5"].value == "ON"     # Reset election
    assert ws["B6"].value == "ON"     # $10m / +25% tier (v2.5: flipped to ON — Bill-correct)
    assert ws["B7"].value == "ON"     # CGT discount
    # B8/B9 used to hold earnings source + manual override — removed in v1.7.


def test_sample_data_badge_present(tmp_path: Path):
    """v1.5: warning badge on row 2 tells staff to overwrite sample data."""
    out = tmp_path / "out.xlsx"
    wb = build_workbook()
    wb.save(out)
    ws = load_workbook(out)["Inputs"]
    badge = ws["A2"].value
    assert badge is not None
    assert "Sample data" in badge
    assert "overwrite" in badge.lower()


def test_all_tabs_protected(tmp_path: Path):
    """v2.5 step 13: Comparison re-locked after Aiden's formatting pass was
    ported back into source. All tabs now protected."""
    out = tmp_path / "out.xlsx"
    wb = build_workbook()
    wb.save(out)
    wb_re = load_workbook(out)
    for tab in ("Inputs", "Analyser", "Notes", "Comparison"):
        assert wb_re[tab].protection.sheet is True, f"{tab} is not protected"


def test_all_tabs_allow_column_resize_under_protection(tmp_path: Path):
    """v2.0.0: protection.sheet stays True but formatColumns/formatRows
    are unlocked so users can drag column borders without unlocking cells.
    v2.5 step 13: Comparison re-locked → invariant applies to it too."""
    out = tmp_path / "out.xlsx"
    wb = build_workbook()
    wb.save(out)
    wb_re = load_workbook(out)
    for sheet in ("Inputs", "Analyser", "Notes", "Comparison"):
        ws = wb_re[sheet]
        assert ws.protection.sheet is True, f"{sheet} should be protected"
        # openpyxl exposes these as bool-ish; allow either False or "0"
        assert ws.protection.formatColumns in (False, "0", 0), (
            f"{sheet} formatColumns must be unlocked under v2 protection scope"
        )
        assert ws.protection.formatRows in (False, "0", 0), (
            f"{sheet} formatRows must be unlocked under v2 protection scope"
        )


def test_input_cells_unlocked(tmp_path: Path):
    """Inputs control-panel and register cells must stay editable under protection."""
    out = tmp_path / "out.xlsx"
    wb = build_workbook()
    wb.save(out)
    ws = load_workbook(out)["Inputs"]
    # Control panel B5-B7 must be unlocked.
    for row in range(5, 8):
        assert ws.cell(row=row, column=2).protection.locked is False, (
            f"Inputs!B{row} should be unlocked"
        )
    # v2.3: Original cost base shifted to col C after Quantity dropped.
    assert ws["C20"].protection.locked is False
    # Member 1 TSB (B11) must be unlocked.
    assert ws["B11"].protection.locked is False


# --- Analyser tab ---------------------------------------------------------

def _analyser(tmp_path: Path):
    out = tmp_path / "out.xlsx"
    wb = build_workbook()
    wb.save(out)
    return load_workbook(out)["Analyser"]


class TestAnalyser:
    def test_mirror_levers_reference_inputs(self, tmp_path: Path):
        """v2.0.0: state strip pushes everything down by 1; mirror strip now at rows 5-7."""
        ws = _analyser(tmp_path)
        for offset in range(3):
            analyser_row = 5 + offset
            inputs_row = 5 + offset
            assert ws.cell(row=analyser_row, column=2).value == f"='Inputs'!B{inputs_row}"

    def test_fund_earnings_formula(self, tmp_path: Path):
        """v2.0.0: fund earnings = SUMIF over col-H positive Div 296 gains
        (col shifted G→H due to row-num col A insert; range shifted to 21-70)."""
        ws = _analyser(tmp_path)
        f = ws["B12"].value
        assert f == '=SUMIF(H21:H70,">0")'

    def test_headline_is_sum_of_member_tax(self, tmp_path: Path):
        """v2.0.0: headline now at B17 (was B16); members 13-16 (were 12-15)."""
        ws = _analyser(tmp_path)
        assert ws["B17"].value == "=SUM(B13:B16)"

    def test_member_tax_formula_references_inputs_member_rows(self, tmp_path: Path):
        """v2.0.0: member 1 Analyser row → 13 (was 12); Inputs row 11 unchanged."""
        ws = _analyser(tmp_path)
        f = ws["B13"].value
        for ref in ("'Inputs'!B11", "'Inputs'!C11", "'Inputs'!D11", "'Inputs'!E11"):
            assert ref in f, f"member 1 formula missing {ref}"
        assert "rate_tier1" in f
        assert "rate_tier2" in f
        assert "threshold_1" in f
        assert "threshold_2" in f

    def test_first_data_row_columns(self, tmp_path: Path):
        """v2.3: Inputs col shifts — D→C (orig CB), F→E (MV at 30 Jun),
        H→G (proceeds). Held>12m still col I."""
        ws = _analyser(tmp_path)
        # Analyser row 21 = first asset; visible data cols 2..10.
        a, b, c, d, e, f, g, h, i = (ws.cell(row=21, column=col).value for col in range(2, 11))
        assert "'Inputs'!A20" in a and "'Inputs'!B20" in a
        assert "'Inputs'!G20" in b                      # Proceeds (was H)
        assert "'Inputs'!C20" in c                      # Original CB (was D)
        # Col E = Ordinary taxable gain
        assert "discount_rate" in d and "'Inputs'!G20" in d and "'Inputs'!C20" in d
        # Col F = Ordinary CGT = MAX(0, E{row}) * fund_cgt_rate
        assert "fund_cgt_rate" in e and "E21" in e
        # Col G = Div 296 cost base
        assert 'reset_on="ON"' in f and "'Inputs'!E20" in f and "'Inputs'!C20" in f
        # Col H = Div 296 adjusted gain
        assert "discount_rate" in g and 'reset_on="ON"' in g
        # Col I = Div 296 tax (pro-rata of headline)
        assert "$B$17" in h and 'SUMIF(H21:H70,">0")' in h
        # Col J = Reset impact (K − L helper diff)
        assert "K21" in i and "L21" in i

    def test_helper_columns_hidden(self, tmp_path: Path):
        """v2.0.0: helpers shifted from J/K to K/L."""
        ws = _analyser(tmp_path)
        assert ws.column_dimensions["K"].hidden
        assert ws.column_dimensions["L"].hidden

    def test_totals_row_sums_correct_ranges(self, tmp_path: Path):
        """v2.0.0: totals row now 71; cols shifted by +1 (B→C, E→F, G→H, H→I)."""
        ws = _analyser(tmp_path)
        assert ws["C71"].value == "=SUM(C21:C70)"        # Proceeds (was B70)
        assert ws["F71"].value == "=SUM(F21:F70)"        # Ord CGT (was E70)
        assert ws["H71"].value == "=SUM(H21:H70)"        # Div 296 adj gain (was G70)
        assert ws["I71"].value == "=SUM(I21:I70)"        # Div 296 tax (was H70)

    def test_reconciliation_panel(self, tmp_path: Path):
        """v2.0.0: recon block rows shift down by 1 (73→74 etc); cols unchanged
        in recon (fund block stays A:B), but per-asset col E→F for the SUMIF."""
        ws = _analyser(tmp_path)
        assert ws["A74"].value == "Ordinary CGT payable"
        assert ws["B74"].value == "=SUM(F21:F70)"
        assert ws["A75"].value == "Div 296 tax payable (headline)"
        assert ws["B75"].value == "=B17"
        assert ws["A76"].value == "Capital losses carried forward"
        cf = ws["B76"].value
        assert cf.startswith("=")
        assert cf.count("MAX(0,'Inputs'!C") == 50  # v2.3: orig CB col D→C

    def test_trap_conditional_formatting_applied(self, tmp_path: Path):
        """v2.0.0: trap CF range now A21:J70 (was A20:I69)."""
        ws = _analyser(tmp_path)
        cf_rules = list(ws.conditional_formatting._cf_rules.items())
        assert cf_rules, "no conditional formatting rules on Analyser"
        ranges = [str(r[0]) for r in cf_rules]
        assert any("A21" in r and "J70" in r for r in ranges)

    def test_state_strip_present(self, tmp_path: Path):
        """v2.0.0: state strip in row 2 references all 3 named ranges + headline."""
        ws = _analyser(tmp_path)
        cell = ws["A2"]
        assert "reset_on" in cell.value
        assert "discount_on" in cell.value
        assert "tier10_on" in cell.value
        assert "B17" in cell.value  # HEADLINE_ROW after shift

    def test_row_number_column(self, tmp_path: Path):
        """v2.0.0: col A is row numbers 1..50 matching Inputs register offset."""
        ws = _analyser(tmp_path)
        assert ws["A20"].value == "#"        # header row (PERASSET_HEADER_ROW)
        assert ws["A21"].value == 1          # first data row
        assert ws["A70"].value == 50         # last data row
        # v2.2.0: Σ glyph replaced with the word "Total" for client readability.
        assert ws["A71"].value == "Total"    # totals row

    def test_print_titles_repeat_header(self, tmp_path: Path):
        """v2.0.0: per-asset header row repeats on every printed page.
        openpyxl round-trips with $-prefix on absolute row refs."""
        ws = _analyser(tmp_path)
        assert ws.print_title_rows in ("20:20", "$20:$20")

    def test_no_freeze_panes(self, tmp_path: Path):
        """v2.0.0: free scroll — no freeze panes."""
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
        """v2.5 polish: strip is per-member TSB only (cols A:C, rows 14-18).
        The earlier v2.5 right-side tiles (proportion / discount / tier) were
        dropped — cols D:K of the strip are deliberately blank."""
        ws = _comparison(tmp_path)
        # Row 12 section band — renamed to "Members & TSB" (no parenthetical).
        assert "Members & TSB" in str(ws["A12"].value)
        # Row 13 (v2.5 step 13 — Aiden polish): 2-column header row, not a
        # merged sub-header anymore. A13 = "Members", B13 = "Total Super Balance".
        assert ws["A13"].value == "Members"
        assert ws["B13"].value == "Total Super Balance"
        # Per-member rows: A14..A17 carry placeholder labels regardless of TSB.
        assert ws["A14"].value == "Member 1"
        assert ws["A15"].value == "Member 2"
        assert ws["A16"].value == "Member 3"
        assert ws["A17"].value == "Member 4"
        # Total row 18 sums all member TSBs from Inputs.
        assert ws["A18"].value == "Total"
        assert "SUM('Inputs'!B11:B14)" in ws["B18"].value
        # Right-side tiles must be GONE — every cell D14..K18 should be None.
        for col in ("D", "E", "F", "G", "H", "I", "J", "K"):
            for row in range(14, 19):
                val = ws[f"{col}{row}"].value
                assert val is None, (
                    f"{col}{row} should be empty (right-side tiles dropped); got {val!r}"
                )

    def test_metric_cards_present(self, tmp_path: Path):
        """v2.5 step 13 (Aiden polish): headline cards use VERBOSE labels —
        long form on the big tile, short form everywhere else. Rows shifted
        +1 from spacer row inserted at r19."""
        ws = _comparison(tmp_path)
        # Card labels at row 21 use the verbose variant (CONTEXT.md headline
        # override). Short form lives on subtotal r26 + per-member r34.
        assert ws["A21"].value == "If no Div 296 CostBase Reset (default)"
        assert ws["E21"].value == "If elected to reset Div 296 CostBase Reset"
        assert ws["I21"].value == "Difference (Net Div 296 Tax)"
        # Card values at row 22 point at the headline cells L6/M6.
        assert "L$6" in ws["A22"].value or "L6" in ws["A22"].value
        assert "M$6" in ws["E22"].value or "M6" in ws["E22"].value
        # Difference card is SIGNED — reset − default (negative = saving).
        net = ws["I22"].value
        assert ("M" in net and "L" in net and "-" in net), f"unexpected difference formula {net!r}"
        assert net.index("M") < net.index("L"), (
            f"Difference card should be reset − default (M-L), got {net!r}"
        )

    def test_subtotals_table(self, tmp_path: Path):
        """v2.5 step 13 (Aiden polish): subtotal rows shifted +2 from spacer
        inserts. Header at r26, data r27-30. Short-form labels (glossary)."""
        ws = _comparison(tmp_path)
        # Header row 26
        assert ws["B26"].value == "If no reset (default)"
        assert ws["C26"].value == "If elected to reset"
        assert ws["D26"].value == "Difference"

        # Row labels (27-30)
        assert ws["A27"].value == "Div 296 earnings"
        assert "Ordinary CGT" in ws["A28"].value and "unchanged" in ws["A28"].value
        assert ws["A29"].value == "Div 296 tax (headline)"
        assert "TOTAL TAX BURDEN" in ws["A30"].value

        # Ordinary CGT row pulls from Analyser (same value in both scenarios)
        assert "Analyser" in ws["B28"].value
        assert ws["B28"].value == ws["C28"].value

        # Total burden = ord CGT + Div 296 tax for each scenario
        assert ws["B30"].value == "=B28+B29"
        assert ws["C30"].value == "=C28+C29"

        # Change col is SIGNED (reset − default).
        for row in (27, 28, 29, 30):
            assert ws[f"D{row}"].value == f"=C{row}-B{row}", (
                f"D{row} should be =C{row}-B{row}, got {ws[f'D{row}'].value!r}"
            )

    def test_panel_a_uses_original_cost_base(self, tmp_path: Path):
        """Panels use INDEX(Inputs!col, matched_row).
        v2.3: orig CB shifted from col D to col C (Quantity dropped)."""
        ws = _comparison(tmp_path)
        cb_formula = ws[f"C{CMP_DATA_FIRST_ROW}"].value
        assert "'Inputs'!$C:$C" in cb_formula
        assert "'Inputs'!$E:$E" not in cb_formula

    def test_panel_b_uses_market_value(self, tmp_path: Path):
        """v2.5 FB-7: Panel B moved to F-J. Cost base is now col H."""
        ws = _comparison(tmp_path)
        cb_formula = ws[f"H{CMP_DATA_FIRST_ROW}"].value
        assert "'Inputs'!$E:$E" in cb_formula
        assert "'Inputs'!$C:$C" not in cb_formula

    def test_panels_independent_of_master_reset_toggle(self, tmp_path: Path):
        """Neither panel formula may reference the reset_on named range."""
        ws = _comparison(tmp_path)
        for row in range(CMP_DATA_FIRST_ROW, CMP_DATA_LAST_ROW + 1):
            # v2.5 FB-7: Panel A B-E, Panel B G-J (col F = Panel B asset label).
            for col_letter in ("B", "C", "D", "E", "G", "H", "I", "J"):
                f = ws[f"{col_letter}{row}"].value
                if f and isinstance(f, str):
                    assert "reset_on" not in f, (
                        f"{col_letter}{row} references reset_on; panels must be independent"
                    )

    def test_delta_column_formula(self, tmp_path: Path):
        """v2.5 FB-7: Change col K = panel B tax (J) − panel A tax (E) — signed tax delta."""
        ws = _comparison(tmp_path)
        delta = ws[f"K{CMP_DATA_FIRST_ROW}"].value
        assert f"J{CMP_DATA_FIRST_ROW}" in delta and f"E{CMP_DATA_FIRST_ROW}" in delta, (
            f"Change formula must reference panel B tax (J) and panel A tax (E); got {delta!r}"
        )
        assert "-" in delta
        # Reset (J) − default (E): the J reference comes BEFORE the E reference.
        assert delta.index(f"J{CMP_DATA_FIRST_ROW}") < delta.index(f"E{CMP_DATA_FIRST_ROW}"), (
            f"Change must be reset − default (J before E), got {delta!r}"
        )

    def test_helper_columns_hidden(self, tmp_path: Path):
        """v1.7: helpers in cols L/M plus per-register grid N/O/P + matched-row R."""
        ws = _comparison(tmp_path)
        for col_letter in ("L", "M", "N", "O", "P", "Q", "R"):
            assert ws.column_dimensions[col_letter].hidden, f"col {col_letter} should be hidden"
        for col_letter in ("A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K"):
            assert not ws.column_dimensions[col_letter].hidden, f"col {col_letter} must be visible"

    def test_top_10_sorted_rendered(self, tmp_path: Path):
        """Display top 10 assets by |Δ (B − A)| descending."""
        ws = _comparison(tmp_path)
        first = ws[f"A{CMP_DATA_FIRST_ROW}"].value
        last = ws[f"A{CMP_DATA_LAST_ROW}"].value
        assert first and first.startswith("=")
        assert last and last.startswith("=")
        overflow = ws[f"A{CMP_OVERFLOW_ROW}"].value
        assert overflow and "top 10" in overflow.lower()
        # v2.2.0: greek delta dropped — must read as plain English now.
        assert "Δ" not in overflow

    def test_per_asset_detail_uses_large_match_sort(self, tmp_path: Path):
        """Each visible row's matched register row is computed via LARGE/MATCH
        in the hidden col R, and panel cells INDEX into Inputs by that row."""
        ws = _comparison(tmp_path)
        matched_formula = ws[f"R{CMP_DATA_FIRST_ROW}"].value
        assert "LARGE" in matched_formula and "MATCH" in matched_formula
        # Panel cells should reference the matched row.
        assert (f"$R{CMP_DATA_FIRST_ROW}" in ws[f"B{CMP_DATA_FIRST_ROW}"].value
                or f"R{CMP_DATA_FIRST_ROW}" in ws[f"B{CMP_DATA_FIRST_ROW}"].value)

    def test_no_recommendation_language(self, tmp_path: Path):
        """Comparison tab must use neutral net-effect language only."""
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
        # Print area now spans through col K (widened from G in v1.5).
        assert ws.print_area is not None and "K" in ws.print_area

    def test_no_chart_v25(self, tmp_path: Path):
        """v2.5 FB-3: chart removed — Aiden's feedback was it didn't tell a story."""
        ws = _comparison(tmp_path)
        assert len(ws._charts) == 0, "Comparison should have no chart (removed v2.5)"


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
            "Loss-offset divergence",
            "Reset OFF scenario is realised-only",
            "Wash sale / Part IVA",
            "Transaction costs",
            "actuarial certificate",
            "all-or-nothing and irrevocable",
            "recontribution",
        ):
            assert required in all_text, f"Notes missing required caveat: {required!r}"

    def test_no_recommendation_language(self, tmp_path: Path):
        """Caveats are factual disclosures only — no advice language."""
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
        """Valuation log has one row per asset, formulas pointing at Inputs.
        v2.3: Valuation source/date moved from Inputs col G to col F (Quantity dropped)."""
        ws = _notes(tmp_path)
        ref_count = 0
        for row in ws.iter_rows():
            for c in row:
                if isinstance(c.value, str) and "'Inputs'!F" in c.value:
                    ref_count += 1
        # 50 valuation-source mirrors + 50 MV-at-30Jun mirrors (col E, not col F)
        # → only valuation-source contributes 'Inputs'!F refs in the log block.
        # However Notes also references Inputs!F elsewhere; assert >= 50.
        assert ref_count >= 50, f"Expected ≥50 valuation-source mirrors, got {ref_count}"

    def test_provenance_cells_present(self, tmp_path: Path):
        """Hidden provenance block carries build_version / build_date / git SHA."""
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
