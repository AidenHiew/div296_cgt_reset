# Ongoing Division 296 Calculator — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Build a single-year, multi-member ongoing Division 296 tax calculator as a new `src/div296_calc/` package that emits a 2-tab Excel workbook (`Calculator`, `Notes`), with a fully-tested pure-Python calc engine behind it.

**Architecture:** A pure calc layer (`assumptions.py`, `calcs.py`) holds all domain logic with **no openpyxl import**. An Excel layer (`_formulas.py` formula-string builders + `tabs/calculator.py` + `tabs/notes.py`) is a projection of the proven pure layer. `build.py` assembles the workbook and runs a **strict** recalc gate (zero error cells, no skip-list). The frozen `src/div296/` package is **not touched**; we reuse `div296.styles` for look-and-feel and mirror its s102-5 netting and formula-builder patterns.

**Tech Stack:** Python 3.11+, openpyxl, pytest (+ pytest-xdist, `slow` marker), the `formulas` package for recalc, ruff (pinned 0.15.14). Spec: [`docs/superpowers/specs/2026-06-20-div296-ongoing-calculator-design.md`](../specs/2026-06-20-div296-ongoing-calculator-design.md).

---

## File structure (created by this plan)

| File | Responsibility |
|---|---|
| `src/div296_calc/__init__.py` | Package marker + `__version__`. |
| `src/div296_calc/assumptions.py` | Year→thresholds table (only confirmed years), enacted rates, `thresholds_for()`, `UnknownYearError`. |
| `src/div296_calc/calcs.py` | Pure engine: CGT netting, pooled total, tax kernel, TSB_REF rule, per-member compute, fund roll-up, share guard. No openpyxl. |
| `src/div296_calc/named_ranges.py` | Defined-name registry (`rate_tier1`, `rate_tier2`, `discount_rate`, `t1_sel`, `t2_sel`, `greater_of_sel`). |
| `src/div296_calc/_formulas.py` | Excel formula-string builders (SUM-safe, guard-first). |
| `src/div296_calc/tabs/__init__.py` | Tabs subpackage marker. |
| `src/div296_calc/tabs/calculator.py` | The single data-entry + output sheet. |
| `src/div296_calc/tabs/notes.py` | Disclaimer / law-basis / threshold-table / caveats sheet. |
| `src/div296_calc/build.py` | `build_workbook()`, strict `validate_recalc()`, `main()` CLI. |
| `tests/div296_calc/test_calc_assumptions.py` | Year table + frozen-parity. |
| `tests/div296_calc/test_calc_engine.py` | Kernel, TSB rule, CGT, pooled, roll-up, guard, four goldens. |
| `tests/div296_calc/test_calc_formulas_golden.py` | Formula-string goldens. |
| `tests/div296_calc/test_calc_build.py` | Structural: tabs, named ranges, key cell formulas, sample values. |
| `tests/div296_calc/test_calc_integration.py` | `slow`: build → recalc → assert sample-fund cells + zero error cells. |

**Note on test filenames:** the repo's `tests/` is flat with unique basenames (pytest default import mode rejects duplicate basenames). All new test files use the `test_calc_*` prefix so they never collide with the existing `tests/test_calcs.py` etc., even though they live in `tests/div296_calc/`.

**Packaging:** `pyproject.toml` already uses `[tool.setuptools.packages.find] where = ["src"]`, which auto-discovers `div296_calc` — **no pyproject edit needed for packaging.**

---

### Task 0: Package scaffold

**Files:**
- Create: `src/div296_calc/__init__.py`
- Create: `src/div296_calc/tabs/__init__.py`
- Create: `tests/div296_calc/test_calc_assumptions.py` (import smoke only for now)

- [ ] **Step 1: Write the failing test**

`tests/div296_calc/test_calc_assumptions.py`:
```python
"""Tests for div296_calc.assumptions — year table + frozen-law parity."""

def test_package_version_present():
    import div296_calc
    assert div296_calc.__version__ == "0.1.0"
```

- [ ] **Step 2: Run test to verify it fails**

Run: `pytest tests/div296_calc/test_calc_assumptions.py -q`
Expected: FAIL — `ModuleNotFoundError: No module named 'div296_calc'`.

- [ ] **Step 3: Write minimal implementation**

`src/div296_calc/__init__.py`:
```python
"""Ongoing Division 296 calculator — single-year, multi-member.

Sibling to the frozen `div296` CGT-reset package, which this does not
import for calculation logic (only `div296.styles` for look-and-feel).
See docs/superpowers/specs/2026-06-20-div296-ongoing-calculator-design.md.
"""

__version__ = "0.1.0"
```

`src/div296_calc/tabs/__init__.py`:
```python
"""Worksheet builders for the ongoing Division 296 calculator."""
```

- [ ] **Step 4: Run test to verify it passes**

Run: `pytest tests/div296_calc/test_calc_assumptions.py -q`
Expected: PASS (1 passed).

- [ ] **Step 5: Commit**

```bash
git add src/div296_calc/__init__.py src/div296_calc/tabs/__init__.py tests/div296_calc/test_calc_assumptions.py
git commit -m "feat(calc): scaffold div296_calc package"
```

---

### Task 1: Assumptions — year table + enacted rates

**Files:**
- Create: `src/div296_calc/assumptions.py`
- Modify: `tests/div296_calc/test_calc_assumptions.py`

- [ ] **Step 1: Write the failing tests**

Append to `tests/div296_calc/test_calc_assumptions.py`:
```python
import pytest

from div296_calc.assumptions import (
    RATE_TIER1, RATE_TIER2, MEMBER_COUNT,
    YearThresholds, thresholds_for, UnknownYearError,
)


def test_2026_27_thresholds():
    yt = thresholds_for("2026-27")
    assert yt == YearThresholds(threshold_1=3_000_000, threshold_2=10_000_000,
                                use_greater_of=False)


def test_unknown_year_raises_with_actionable_message():
    with pytest.raises(UnknownYearError) as exc:
        thresholds_for("2027-28")
    assert "2027-28" in str(exc.value)
    assert "add" in str(exc.value).lower()


def test_rates_match_frozen_div296_assumptions():
    # The ongoing tool must not drift from the shipped CGT-reset tool's rates.
    from div296.assumptions import ASSUMPTIONS
    assert RATE_TIER1 == ASSUMPTIONS.rate_tier1 == 0.15
    assert RATE_TIER2 == ASSUMPTIONS.rate_tier2 == 0.25


def test_2026_27_thresholds_match_frozen_constants():
    from div296.assumptions import ASSUMPTIONS
    yt = thresholds_for("2026-27")
    assert yt.threshold_1 == ASSUMPTIONS.threshold_1
    assert yt.threshold_2 == ASSUMPTIONS.threshold_2


def test_member_cap_is_four():
    assert MEMBER_COUNT == 4
```

- [ ] **Step 2: Run to verify it fails**

Run: `pytest tests/div296_calc/test_calc_assumptions.py -q`
Expected: FAIL — `ImportError: cannot import name ... from 'div296_calc.assumptions'`.

- [ ] **Step 3: Write minimal implementation**

`src/div296_calc/assumptions.py`:
```python
"""Rates, member cap, and the year→thresholds table.

The table is SEEDED WITH ONLY CONFIRMED YEARS. At v0.1 that is just
2026-27 ($3M / $10M, un-indexed). Later years' thresholds are CPI-indexed
(in $150k / $500k steps) and not yet published — add them by hand to
YEAR_TABLE as the ATO announces them. Looking up an absent year raises
UnknownYearError with an actionable message (no silent default).

rate_tier1 / rate_tier2 deliberately mirror the frozen div296.assumptions
so the two tools never drift; a test pins the equality.
"""

from __future__ import annotations

from dataclasses import dataclass

# Enacted Division 296 additional rates (slice form: 15% on $3M–$10M,
# 25% = 15%+extra 10% on the slice above $10M). See the spec §5 footgun note.
RATE_TIER1: float = 0.15
RATE_TIER2: float = 0.25

# SMSF 1/3 CGT discount — used by the CGT netting helper.
DISCOUNT_RATE: float = 1.0 / 3.0

# Up to 4 members per SMSF.
MEMBER_COUNT: int = 4


@dataclass(frozen=True)
class YearThresholds:
    threshold_1: int          # lower-super-balance threshold (LSBT), e.g. $3M
    threshold_2: int          # very-large-super-balance threshold (VLSBT), e.g. $10M
    use_greater_of: bool      # False for 2026-27 (closing TSB); True from 2027-28


# Only confirmed years. ADD A ROW per ATO-published indexed thresholds.
YEAR_TABLE: dict[str, YearThresholds] = {
    "2026-27": YearThresholds(3_000_000, 10_000_000, use_greater_of=False),
}


class UnknownYearError(KeyError):
    """Raised when an income year is not in YEAR_TABLE."""


def thresholds_for(year: str) -> YearThresholds:
    """Return the thresholds row for `year`, or raise UnknownYearError."""
    try:
        return YEAR_TABLE[year]
    except KeyError:
        raise UnknownYearError(
            f"No threshold row for income year {year!r}. Add the ATO-published "
            f"indexed $3M/$10M thresholds for {year} to "
            f"div296_calc.assumptions.YEAR_TABLE before running."
        ) from None
```

- [ ] **Step 4: Run to verify it passes**

Run: `pytest tests/div296_calc/test_calc_assumptions.py -q`
Expected: PASS (6 passed).

- [ ] **Step 5: Commit**

```bash
git add src/div296_calc/assumptions.py tests/div296_calc/test_calc_assumptions.py
git commit -m "feat(calc): year thresholds table + enacted rates with frozen-parity tests"
```

---

### Task 2: Tax kernel (slice form) + TSB_REF rule

**Files:**
- Create: `src/div296_calc/calcs.py`
- Create: `tests/div296_calc/test_calc_engine.py`

- [ ] **Step 1: Write the failing tests**

`tests/div296_calc/test_calc_engine.py`:
```python
"""Pure-engine tests for div296_calc.calcs."""

import pytest

from div296_calc.assumptions import RATE_TIER1, RATE_TIER2
from div296_calc.calcs import member_div296_tax, tsb_ref


T1, T2 = 3_000_000, 10_000_000


def _tax(net, tsb):
    return member_div296_tax(net, tsb, T1, T2, RATE_TIER1, RATE_TIER2)


# --- mandatory >$10M two-tier anchor (the footgun guard) ---
def test_emma_two_tier_above_10m():
    assert _tax(840_000, 12_900_000) == pytest.approx(115_581.40, abs=0.01)


def test_jess_single_tier():
    assert _tax(476_625, 4_500_000) == pytest.approx(23_831.25, abs=0.01)


def test_jill_just_above_3m():
    assert _tax(100_000, 3_100_000) == pytest.approx(483.87, abs=0.01)


# --- boundaries ---
def test_tsb_exactly_3m_is_zero():
    assert _tax(500_000, 3_000_000) == 0.0


def test_tsb_exactly_10m_is_tier1_only():
    # band1 = (10M-3M)/10M = 0.7, band2 = 0
    assert _tax(100_000, 10_000_000) == pytest.approx(100_000 * 0.7 * 0.15)


def test_tsb_below_3m_with_earnings_is_zero():
    assert _tax(100_000, 2_500_000) == 0.0


def test_zero_tsb_is_zero():
    assert _tax(100_000, 0) == 0.0


def test_zero_or_negative_net_is_zero():
    assert _tax(0, 4_000_000) == 0.0
    assert _tax(-50_000, 4_000_000) == 0.0


# --- TSB_REF rule ---
def test_tsb_ref_2026_27_uses_closing():
    from div296_calc.calcs import Member
    m = Member(name="A", opening_tsb=5_000_000, closing_tsb=4_000_000, share=1.0)
    assert tsb_ref(m, use_greater_of=False) == 4_000_000


def test_tsb_ref_greater_of_when_opening_higher():
    from div296_calc.calcs import Member
    m = Member(name="A", opening_tsb=5_000_000, closing_tsb=4_000_000, share=1.0)
    assert tsb_ref(m, use_greater_of=True) == 5_000_000


def test_tsb_ref_greater_of_when_closing_higher():
    from div296_calc.calcs import Member
    m = Member(name="A", opening_tsb=4_000_000, closing_tsb=6_000_000, share=1.0)
    assert tsb_ref(m, use_greater_of=True) == 6_000_000
```

- [ ] **Step 2: Run to verify it fails**

Run: `pytest tests/div296_calc/test_calc_engine.py -q`
Expected: FAIL — `ModuleNotFoundError`/`ImportError` for `div296_calc.calcs`.

- [ ] **Step 3: Write minimal implementation**

`src/div296_calc/calcs.py` (kernel + Member + tsb_ref — more added in later tasks):
```python
"""Pure-Python Division 296 ongoing-year calc engine.

The Excel workbook is the deliverable; this module mirrors every formula
so the test suite can verify the numbers without Excel. NO openpyxl import
in this file (or assumptions.py) — the calc layer is openpyxl-free.

Tax kernel uses the disjoint-SLICE form: 15% on the $3M–$10M slice, 25%
(=15%+extra 10%) on the slice above $10M. This is algebraically identical
to the law's "15% above $3M plus extra 10% above $10M"; pairing the slice
band2 with 0.10 instead of 0.25 undercounts (the documented footgun).
"""

from __future__ import annotations

from dataclasses import dataclass


@dataclass(frozen=True)
class Member:
    name: str
    opening_tsb: float
    closing_tsb: float
    share: float                      # 0.0–1.0 of pooled earnings
    prior_loss: float = 0.0           # prior-year Div 296 carried-forward loss (>=0)
    override: float | None = None     # segregated earnings; None ⇒ pooled


def tsb_ref(member: Member, use_greater_of: bool) -> float:
    """The balance the proportions are measured against.

    2026-27: closing TSB. 2027-28+: greater of opening/closing (anti-avoidance).
    """
    if use_greater_of:
        return max(member.opening_tsb, member.closing_tsb)
    return member.closing_tsb


def member_div296_tax(
    net_earnings: float,
    tsb_ref_value: float,
    threshold_1: float,
    threshold_2: float,
    rate_tier1: float,
    rate_tier2: float,
) -> float:
    """Div 296 tax on a member's net Div 296 earnings.

    Returns 0 when there are no positive net earnings or the balance is at
    or below threshold_1.
    """
    if net_earnings <= 0 or tsb_ref_value <= threshold_1:
        return 0.0
    band1 = max(0.0, min(tsb_ref_value, threshold_2) - threshold_1) / tsb_ref_value
    band2 = max(0.0, tsb_ref_value - threshold_2) / tsb_ref_value
    return net_earnings * (band1 * rate_tier1 + band2 * rate_tier2)
```

- [ ] **Step 4: Run to verify it passes**

Run: `pytest tests/div296_calc/test_calc_engine.py -q`
Expected: PASS (11 passed). The Emma case confirms `rate_tier2 = 0.25` paired with the slice band → $115,581.40.

- [ ] **Step 5: Commit**

```bash
git add src/div296_calc/calcs.py tests/div296_calc/test_calc_engine.py
git commit -m "feat(calc): slice-form tax kernel + TSB_REF rule (Emma >\$10M anchor)"
```

---

### Task 3: CGT netting helper (s102-5)

**Files:**
- Modify: `src/div296_calc/calcs.py`
- Modify: `tests/div296_calc/test_calc_engine.py`

- [ ] **Step 1: Write the failing tests**

Append to `tests/div296_calc/test_calc_engine.py`:
```python
from div296_calc.assumptions import DISCOUNT_RATE
from div296_calc.calcs import CgtInputs, CgtResult, net_capital_gain


def test_cgt_long_held_gain_gets_one_third_discount():
    r = net_capital_gain(CgtInputs(gross_gains_over_12m=300_000,
                                   gross_gains_under_12m=0,
                                   capital_losses=0), DISCOUNT_RATE)
    assert r.net_realised_cg == pytest.approx(200_000)   # 300k × 2/3
    assert r.unused_capital_loss == 0


def test_cgt_losses_hit_non_discount_gains_first():
    # 60k loss eats the 50k short-held gain, then 10k of the long-held gain.
    # remaining long-held = 240k → ×2/3 = 160k.
    r = net_capital_gain(CgtInputs(gross_gains_over_12m=250_000,
                                   gross_gains_under_12m=50_000,
                                   capital_losses=60_000), DISCOUNT_RATE)
    assert r.net_realised_cg == pytest.approx(160_000)
    assert r.unused_capital_loss == 0


def test_cgt_losses_exceeding_all_gains_carry_forward():
    r = net_capital_gain(CgtInputs(gross_gains_over_12m=100_000,
                                   gross_gains_under_12m=20_000,
                                   capital_losses=200_000), DISCOUNT_RATE)
    assert r.net_realised_cg == 0
    assert r.unused_capital_loss == pytest.approx(80_000)   # 200k − 120k


def test_cgt_result_is_frozen_dataclass():
    r = net_capital_gain(CgtInputs(0, 0, 0), DISCOUNT_RATE)
    assert isinstance(r, CgtResult)
```

- [ ] **Step 2: Run to verify it fails**

Run: `pytest tests/div296_calc/test_calc_engine.py -q -k cgt`
Expected: FAIL — `ImportError` for `CgtInputs`/`net_capital_gain`.

- [ ] **Step 3: Write minimal implementation**

Append to `src/div296_calc/calcs.py`:
```python
@dataclass(frozen=True)
class CgtInputs:
    gross_gains_over_12m: float       # discountable
    gross_gains_under_12m: float      # non-discountable
    capital_losses: float             # current-year + brought-forward


@dataclass(frozen=True)
class CgtResult:
    net_realised_cg: float
    unused_capital_loss: float        # carries forward to next year's losses


def net_capital_gain(cgt: CgtInputs, discount_rate: float) -> CgtResult:
    """s102-5 method: losses hit non-discount gains first, then the 1/3
    discount applies to the surviving long-held remainder."""
    loss = cgt.capital_losses
    nondisc_net = max(0.0, cgt.gross_gains_under_12m - loss)
    loss = max(0.0, loss - cgt.gross_gains_under_12m)
    disc_net = max(0.0, cgt.gross_gains_over_12m - loss)
    net = nondisc_net + disc_net * (1.0 - discount_rate)
    unused = max(0.0, loss - cgt.gross_gains_over_12m)
    return CgtResult(net_realised_cg=net, unused_capital_loss=unused)
```

- [ ] **Step 4: Run to verify it passes**

Run: `pytest tests/div296_calc/test_calc_engine.py -q -k cgt`
Expected: PASS (4 passed).

- [ ] **Step 5: Commit**

```bash
git add src/div296_calc/calcs.py tests/div296_calc/test_calc_engine.py
git commit -m "feat(calc): s102-5 CGT netting helper with 1/3 discount + carry-forward"
```

---

### Task 4: Pooled total + per-member allocation

**Files:**
- Modify: `src/div296_calc/calcs.py`
- Modify: `tests/div296_calc/test_calc_engine.py`

- [ ] **Step 1: Write the failing tests**

Append to `tests/div296_calc/test_calc_engine.py`:
```python
from div296_calc.calcs import PooledIncome, pooled_total, member_earnings


def test_pooled_total_sums_income_less_expenses():
    p = PooledIncome(dividends_grossed=120_000, interest=30_000, rent=80_000,
                     other=0, net_realised_cg=160_000, deductible_expenses=40_000)
    assert pooled_total(p) == 350_000


def test_pooled_total_can_go_negative():
    p = PooledIncome(dividends_grossed=10_000, interest=0, rent=0, other=0,
                     net_realised_cg=0, deductible_expenses=60_000)
    assert pooled_total(p) == -50_000


def test_member_earnings_pooled_is_share_times_pool():
    m = Member("A", 4_000_000, 4_000_000, share=0.6)
    assert member_earnings(m, 350_000) == pytest.approx(210_000)


def test_member_earnings_override_wins_and_ignores_pool():
    m = Member("A", 4_000_000, 4_000_000, share=0.6, override=90_000)
    assert member_earnings(m, 350_000) == 90_000


def test_member_earnings_override_zero_is_honoured_not_blank():
    m = Member("A", 4_000_000, 4_000_000, share=0.6, override=0.0)
    assert member_earnings(m, 350_000) == 0.0


def test_member_earnings_negative_pool_allocates_negative():
    m = Member("A", 4_000_000, 4_000_000, share=0.5)
    assert member_earnings(m, -50_000) == pytest.approx(-25_000)
```

- [ ] **Step 2: Run to verify it fails**

Run: `pytest tests/div296_calc/test_calc_engine.py -q -k "pooled or earnings"`
Expected: FAIL — `ImportError` for `PooledIncome`/`member_earnings`.

- [ ] **Step 3: Write minimal implementation**

Append to `src/div296_calc/calcs.py`:
```python
@dataclass(frozen=True)
class PooledIncome:
    dividends_grossed: float          # incl. franking credits, per the fund return
    interest: float
    rent: float
    other: float
    net_realised_cg: float            # from net_capital_gain()
    deductible_expenses: float


def pooled_total(p: PooledIncome) -> float:
    """Realised income less deductible expenses. May be negative."""
    return (p.dividends_grossed + p.interest + p.rent + p.other
            + p.net_realised_cg - p.deductible_expenses)


def member_earnings(member: Member, pool_total: float) -> float:
    """Segregated members earn only their override; otherwise share × pool.

    `override is None` ⇒ pooled. An override of 0.0 is honoured (not treated
    as 'no override').
    """
    if member.override is not None:
        return member.override
    return member.share * pool_total
```

- [ ] **Step 4: Run to verify it passes**

Run: `pytest tests/div296_calc/test_calc_engine.py -q -k "pooled or earnings"`
Expected: PASS (6 passed).

- [ ] **Step 5: Commit**

```bash
git add src/div296_calc/calcs.py tests/div296_calc/test_calc_engine.py
git commit -m "feat(calc): pooled total (less expenses) + override-aware member allocation"
```

---

### Task 5: Per-member result + fund roll-up + share guard

**Files:**
- Modify: `src/div296_calc/calcs.py`
- Modify: `tests/div296_calc/test_calc_engine.py`

- [ ] **Step 1: Write the failing tests**

Append to `tests/div296_calc/test_calc_engine.py`:
```python
from div296_calc.assumptions import thresholds_for
from div296_calc.calcs import (
    MemberResult, compute_member, compute_fund, fund_total_tax,
    pooled_share_status,
)

YT = thresholds_for("2026-27")


def _member_result(m, pool):
    return compute_member(m, pool, YT, RATE_TIER1, RATE_TIER2)


def test_compute_member_below_threshold_status():
    m = Member("Below", 2_500_000, 2_500_000, share=1.0)
    r = _member_result(m, 100_000)
    assert r.tax == 0.0
    assert r.below_threshold is True
    assert r.new_loss == 0.0


def test_compute_member_emma_via_override():
    m = Member("Emma", 12_900_000, 12_900_000, share=0.0, override=840_000)
    r = _member_result(m, 0)
    assert r.tax == pytest.approx(115_581.40, abs=0.01)
    assert r.tier2_tax > 0          # the >$10M slice is taxed
    assert r.below_threshold is False


def test_compute_member_prior_loss_partially_absorbs():
    m = Member("P", 4_000_000, 4_000_000, share=1.0, prior_loss=100_000)
    r = _member_result(m, 200_000)            # net = 100,000
    assert r.net_earnings == pytest.approx(100_000)
    assert r.tax == pytest.approx(100_000 * 0.25 * 0.15)   # band1 = 1M/4M


def test_compute_member_prior_loss_exceeds_earnings_carries_forward():
    m = Member("L", 4_000_000, 4_000_000, share=1.0, prior_loss=250_000)
    r = _member_result(m, 200_000)            # net = -50,000
    assert r.tax == 0.0
    assert r.new_loss == pytest.approx(50_000)


def test_jamal_negative_earnings_carry_forward():
    m = Member("Jamal", 5_000_000, 5_000_000, share=1.0, override=-200_000)
    r = _member_result(m, 0)
    assert r.tax == 0.0
    assert r.new_loss == pytest.approx(200_000)


def test_compute_fund_rolls_up_sum_including_zero_member():
    members = [
        Member("Above", 4_000_000, 4_000_000, share=0.5),
        Member("Below", 2_000_000, 2_000_000, share=0.5),
    ]
    results = compute_fund(members, 400_000, YT, RATE_TIER1, RATE_TIER2)
    assert len(results) == 2
    # Above: earnings 200k, band1 = 1M/4M = .25 → 200k×.25×.15 = 7,500
    assert fund_total_tax(results) == pytest.approx(7_500)


def test_pooled_share_status_100_percent_ok():
    members = [Member("A", 4e6, 4e6, share=0.6), Member("B", 5e6, 5e6, share=0.4)]
    status = pooled_share_status(members)
    assert status.pooled_count == 2
    assert status.share_sum == pytest.approx(1.0)
    assert status.ok is True
    assert status.all_segregated is False


def test_pooled_share_status_soft_warn_when_not_100():
    members = [Member("A", 4e6, 4e6, share=0.6), Member("B", 5e6, 5e6, share=0.3)]
    status = pooled_share_status(members)
    assert status.ok is False


def test_pooled_share_status_excludes_overridden_members():
    members = [
        Member("Pooled", 4e6, 4e6, share=1.0),
        Member("Seg", 5e6, 5e6, share=0.0, override=50_000),
    ]
    status = pooled_share_status(members)
    assert status.pooled_count == 1
    assert status.share_sum == pytest.approx(1.0)
    assert status.ok is True


def test_pooled_share_status_all_overridden_is_suppressed():
    members = [
        Member("S1", 4e6, 4e6, share=0.0, override=10_000),
        Member("S2", 5e6, 5e6, share=0.0, override=20_000),
    ]
    status = pooled_share_status(members)
    assert status.pooled_count == 0
    assert status.all_segregated is True
    assert status.ok is True          # no false warning
```

- [ ] **Step 2: Run to verify it fails**

Run: `pytest tests/div296_calc/test_calc_engine.py -q -k "compute or share or jamal"`
Expected: FAIL — `ImportError` for `compute_member` etc.

- [ ] **Step 3: Write minimal implementation**

Append to `src/div296_calc/calcs.py`:
```python
from typing import Sequence


@dataclass(frozen=True)
class MemberResult:
    name: str
    earnings: float
    tsb_ref: float
    used_greater_of: bool
    net_earnings: float
    band1: float
    band2: float
    tier1_tax: float
    tier2_tax: float
    tax: float
    new_loss: float
    below_threshold: bool


def compute_member(
    member: Member,
    pool_total: float,
    yt,                               # YearThresholds
    rate_tier1: float,
    rate_tier2: float,
) -> MemberResult:
    earnings = member_earnings(member, pool_total)
    ref = tsb_ref(member, yt.use_greater_of)
    net = earnings - member.prior_loss
    below = ref <= yt.threshold_1

    if ref > 0:
        band1 = max(0.0, min(ref, yt.threshold_2) - yt.threshold_1) / ref
        band2 = max(0.0, ref - yt.threshold_2) / ref
    else:
        band1 = band2 = 0.0

    if net <= 0 or below:
        tier1 = tier2 = 0.0
        new_loss = max(0.0, -net)
    else:
        tier1 = net * band1 * rate_tier1
        tier2 = net * band2 * rate_tier2
        new_loss = 0.0

    return MemberResult(
        name=member.name, earnings=earnings, tsb_ref=ref,
        used_greater_of=yt.use_greater_of, net_earnings=net,
        band1=band1, band2=band2, tier1_tax=tier1, tier2_tax=tier2,
        tax=tier1 + tier2, new_loss=new_loss, below_threshold=below,
    )


def compute_fund(
    members: Sequence[Member],
    pool_total: float,
    yt,
    rate_tier1: float,
    rate_tier2: float,
) -> list[MemberResult]:
    return [compute_member(m, pool_total, yt, rate_tier1, rate_tier2)
            for m in members]


def fund_total_tax(results: Sequence[MemberResult]) -> float:
    return sum(r.tax for r in results)


@dataclass(frozen=True)
class ShareStatus:
    pooled_count: int
    share_sum: float
    all_segregated: bool
    ok: bool


def pooled_share_status(members: Sequence[Member]) -> ShareStatus:
    """Soft guard on pooled members' shares. Suppressed (ok=True) when every
    member is segregated (overridden), since there is then no pool to split."""
    pooled = [m for m in members if m.override is None]
    share_sum = sum(m.share for m in pooled)
    all_seg = len(pooled) == 0
    ok = all_seg or abs(share_sum - 1.0) < 1e-9
    return ShareStatus(pooled_count=len(pooled), share_sum=share_sum,
                       all_segregated=all_seg, ok=ok)
```

- [ ] **Step 4: Run to verify it passes**

Run: `pytest tests/div296_calc/test_calc_engine.py -q`
Expected: PASS (all engine tests green).

- [ ] **Step 5: Commit**

```bash
git add src/div296_calc/calcs.py tests/div296_calc/test_calc_engine.py
git commit -m "feat(calc): per-member result, fund roll-up, all-overridden-aware share guard"
```

---

### Task 6: Excel formula-string builders + named ranges

**Files:**
- Create: `src/div296_calc/named_ranges.py`
- Create: `src/div296_calc/_formulas.py`
- Create: `tests/div296_calc/test_calc_formulas_golden.py`

- [ ] **Step 1: Write the failing tests**

`tests/div296_calc/test_calc_formulas_golden.py`:
```python
"""Golden-string tests pinning the exact arithmetic of generated formulas.

If one fails after an INTENTIONAL formula change, re-verify the new string
against div296_calc.calcs (the pure mirror) and the spec §5/§7, then update
the golden in the same commit.
"""

from div296_calc import _formulas as F


def test_pooled_total_formula_uses_sum_not_plus():
    assert F.pooled_total_formula("B6", "B7", "B8", "B9", "B10", "B11") == (
        "=SUM(B6,B7,B8,B9,B10)-B11"
    )


def test_cgt_net_formula_golden():
    assert F.cgt_net_formula("B15", "B16", "B17") == (
        "=MAX(0,B16-B17)+MAX(0,B15-MAX(0,B17-B16))*(1-discount_rate)"
    )


def test_cgt_unused_loss_formula_golden():
    assert F.cgt_unused_loss_formula("B15", "B16", "B17") == (
        "=MAX(0,MAX(0,B17-B16)-B15)"
    )


def test_tsb_ref_formula_golden():
    assert F.tsb_ref_formula("B23", "B24", "greater_of_sel") == (
        '=IF(B24="","",IF(greater_of_sel=1,MAX(B23,B24),B24))'
    )


def test_earnings_formula_override_then_pooled_with_name_guard():
    assert F.earnings_formula("B22", "B27", "B25", "$B$12") == (
        '=IF(B22="","",IF(B27<>"",B27,IF(OR(B25="",$B$12=""),"",B25*$B$12)))'
    )


def test_net_earnings_formula_golden():
    assert F.net_earnings_formula("B28", "B26") == (
        '=IF(B28="","",B28-IF(B26="",0,B26))'
    )


def test_band1_formula_guard_first():
    assert F.band1_formula("B29", "t1_sel", "t2_sel") == (
        '=IF(OR(B29="",B29<=0),"",MAX(0,MIN(B29,t2_sel)-t1_sel)/B29)'
    )


def test_band2_formula_only_references_t2():
    assert F.band2_formula("B29", "t2_sel") == (
        '=IF(OR(B29="",B29<=0),"",MAX(0,B29-t2_sel)/B29)'
    )


def test_tier1_tax_formula_golden():
    assert F.tier1_tax_formula("B30", "B31", "B29", "t1_sel") == (
        '=IF(OR(B30="",B30<=0,B29<=t1_sel,B31=""),0,B30*B31*rate_tier1)'
    )


def test_tier2_tax_formula_golden():
    assert F.tier2_tax_formula("B30", "B32", "B29", "t1_sel") == (
        '=IF(OR(B30="",B30<=0,B29<=t1_sel,B32=""),0,B30*B32*rate_tier2)'
    )


def test_new_loss_formula_golden():
    assert F.new_loss_formula("B30") == '=IF(B30="","",MAX(0,-B30))'


def test_status_formula_golden():
    assert F.status_formula("B29", "t1_sel") == (
        '=IF(B29="","",IF(B29<=t1_sel,"Below $3M — not liable","Liable"))'
    )


def test_threshold_lookup_uses_sumifs_not_index_match():
    assert F.threshold_lookup_formula("$P$2:$P$10", "$O$2:$O$10", "$B$3") == (
        "=SUMIFS($P$2:$P$10,$O$2:$O$10,$B$3)"
    )


def test_year_known_guard_formula_actionable():
    out = F.year_known_guard_formula("$O$2:$O$10", "$B$3")
    assert out.startswith('=IF(COUNTIF($O$2:$O$10,$B$3)=0,')
    assert "add" in out.lower()


def test_pooled_share_contrib_formula_golden():
    assert F.pooled_share_contrib_formula("B22", "B27", "B25") == (
        '=IF(AND(B22<>"",B27=""),IF(B25="",0,B25),0)'
    )


def test_pooled_count_formula_uses_boolean_sumproduct():
    assert F.pooled_count_formula("B22:E22", "B27:E27") == (
        '=SUMPRODUCT((B22:E22<>"")*(B27:E27=""))'
    )


def test_share_guard_status_formula_handles_all_segregated():
    out = F.share_guard_status_formula("B38", "B39")
    assert "all members segregated" in out
    assert "100%" in out
```

- [ ] **Step 2: Run to verify it fails**

Run: `pytest tests/div296_calc/test_calc_formulas_golden.py -q`
Expected: FAIL — `ModuleNotFoundError` for `div296_calc._formulas`.

- [ ] **Step 3: Write minimal implementation**

`src/div296_calc/named_ranges.py`:
```python
"""Defined-name registry for the ongoing calculator.

Rates + discount mirror the frozen tool's names. Threshold cells are the
RESOLVED per-year values (looked up from the year table), not raw constants.
"""

RATE_TIER1 = "rate_tier1"
RATE_TIER2 = "rate_tier2"
DISCOUNT_RATE = "discount_rate"
T1_SEL = "t1_sel"               # resolved threshold_1 for the selected year
T2_SEL = "t2_sel"               # resolved threshold_2 for the selected year
GREATER_OF_SEL = "greater_of_sel"   # resolved use_greater_of (1/0)

ALL_NAMES = (RATE_TIER1, RATE_TIER2, DISCOUNT_RATE, T1_SEL, T2_SEL, GREATER_OF_SEL)
```

`src/div296_calc/_formulas.py`:
```python
"""Excel formula-string builders for the Calculator tab.

Discipline (spec §7): aggregation via SUM (text coerces to 0, '+' propagates
#VALUE!); guard-first IF for every division/blank-render; threshold lookup via
exact-match SUMIFS (never INDEX/MATCH/VLOOKUP — recalc-gate false positives).
Boolean-only SUMPRODUCT (no text in the multiplied ranges) is allowed.
"""

from __future__ import annotations


def pooled_total_formula(div_c, int_c, rent_c, other_c, cg_c, exp_c) -> str:
    return f"=SUM({div_c},{int_c},{rent_c},{other_c},{cg_c})-{exp_c}"


def cgt_net_formula(over_c, under_c, loss_c) -> str:
    return (f"=MAX(0,{under_c}-{loss_c})"
            f"+MAX(0,{over_c}-MAX(0,{loss_c}-{under_c}))*(1-discount_rate)")


def cgt_unused_loss_formula(over_c, under_c, loss_c) -> str:
    return f"=MAX(0,MAX(0,{loss_c}-{under_c})-{over_c})"


def tsb_ref_formula(open_c, close_c, greater_of_c) -> str:
    return (f'=IF({close_c}="","",'
            f'IF({greater_of_c}=1,MAX({open_c},{close_c}),{close_c}))')


def earnings_formula(name_c, override_c, share_c, pool_c) -> str:
    return (f'=IF({name_c}="","",'
            f'IF({override_c}<>"",{override_c},'
            f'IF(OR({share_c}="",{pool_c}=""),"",{share_c}*{pool_c})))')


def net_earnings_formula(earnings_c, prior_loss_c) -> str:
    return (f'=IF({earnings_c}="","",'
            f'{earnings_c}-IF({prior_loss_c}="",0,{prior_loss_c}))')


def band1_formula(tsb_ref_c, t1_c, t2_c) -> str:
    return (f'=IF(OR({tsb_ref_c}="",{tsb_ref_c}<=0),"",'
            f'MAX(0,MIN({tsb_ref_c},{t2_c})-{t1_c})/{tsb_ref_c})')


def band2_formula(tsb_ref_c, t2_c) -> str:
    return (f'=IF(OR({tsb_ref_c}="",{tsb_ref_c}<=0),"",'
            f'MAX(0,{tsb_ref_c}-{t2_c})/{tsb_ref_c})')


def tier1_tax_formula(net_c, band1_c, tsb_ref_c, t1_c) -> str:
    return (f'=IF(OR({net_c}="",{net_c}<=0,{tsb_ref_c}<={t1_c},{band1_c}=""),0,'
            f'{net_c}*{band1_c}*rate_tier1)')


def tier2_tax_formula(net_c, band2_c, tsb_ref_c, t1_c) -> str:
    return (f'=IF(OR({net_c}="",{net_c}<=0,{tsb_ref_c}<={t1_c},{band2_c}=""),0,'
            f'{net_c}*{band2_c}*rate_tier2)')


def new_loss_formula(net_c) -> str:
    return f'=IF({net_c}="","",MAX(0,-{net_c}))'


def status_formula(tsb_ref_c, t1_c) -> str:
    return (f'=IF({tsb_ref_c}="","",'
            f'IF({tsb_ref_c}<={t1_c},"Below $3M — not liable","Liable"))')


def threshold_lookup_formula(value_range, year_range, year_c) -> str:
    return f"=SUMIFS({value_range},{year_range},{year_c})"


def year_known_guard_formula(year_range, year_c) -> str:
    return (f'=IF(COUNTIF({year_range},{year_c})=0,'
            f'"⚠ Income year not in the threshold table — add the '
            f'ATO-published $3M/$10M row before relying on this.",'
            f'"✓ thresholds loaded")')


def pooled_share_contrib_formula(name_c, override_c, share_c) -> str:
    return (f'=IF(AND({name_c}<>"",{override_c}=""),'
            f'IF({share_c}="",0,{share_c}),0)')


def pooled_count_formula(name_range, override_range) -> str:
    return f'=SUMPRODUCT(({name_range}<>"")*({override_range}=""))'


def share_guard_status_formula(pooled_count_c, share_sum_c) -> str:
    return (f'=IF({pooled_count_c}=0,"n/a — all members segregated",'
            f'IF(ROUND({share_sum_c},4)=1,"✓ pooled shares = 100%",'
            f'"⚠ pooled shares ≠ 100% — check member shares"))')
```

- [ ] **Step 4: Run to verify it passes**

Run: `pytest tests/div296_calc/test_calc_formulas_golden.py -q`
Expected: PASS (all golden-string tests green).

> Note: the builders emit literal `⚠`/`✓`/`—`/`≠` glyphs (matching the frozen tool's `inputs.py`, which already uses `⚠`/`➤`/`✓`). The golden tests compare literal-to-literal, so the assertions hold. ruff permits non-ASCII in string literals.

- [ ] **Step 5: Commit**

```bash
git add src/div296_calc/named_ranges.py src/div296_calc/_formulas.py tests/div296_calc/test_calc_formulas_golden.py
git commit -m "feat(calc): SUM-safe, guard-first Excel formula builders + named ranges"
```

---

### Task 7: Calculator tab

**Files:**
- Create: `src/div296_calc/tabs/calculator.py`
- Create: `tests/div296_calc/test_calc_build.py`

This tab is **members-as-columns** (spec §4): col A holds row labels, cols B–E hold members 1–4. A hidden year-table block lives in cols O–R. The build is one complete function; the test pins the layout constants, the named ranges, the load-bearing cell formulas (reusing the Task 6 builders), and the seeded sample fund.

- [ ] **Step 1: Write the failing tests**

`tests/div296_calc/test_calc_build.py`:
```python
"""Structural tests for the built workbook (fast — no recalc)."""

import pytest

from div296_calc.build import build_workbook
from div296_calc.tabs import calculator as C
from div296_calc import _formulas as F


@pytest.fixture(scope="module")
def wb():
    return build_workbook()


def test_two_tabs_in_order(wb):
    assert wb.sheetnames == ["Calculator", "Notes"]


def test_named_ranges_defined(wb):
    for name in ("rate_tier1", "rate_tier2", "discount_rate",
                 "t1_sel", "t2_sel", "greater_of_sel"):
        assert name in wb.defined_names


def test_pooled_total_cell_formula(wb):
    ws = wb["Calculator"]
    c = ws[f"B{C.POOLED_TOTAL_ROW}"].value
    assert c == F.pooled_total_formula(
        f"B{C.DIVIDENDS_ROW}", f"B{C.INTEREST_ROW}", f"B{C.RENT_ROW}",
        f"B{C.OTHER_ROW}", f"B{C.NET_CG_ROW}", f"B{C.EXPENSES_ROW}",
    )


def test_net_cg_cell_pulls_from_cgt_helper(wb):
    ws = wb["Calculator"]
    assert ws[f"B{C.NET_CG_ROW}"].value == F.cgt_net_formula(
        f"B{C.CGT_OVER_ROW}", f"B{C.CGT_UNDER_ROW}", f"B{C.CGT_LOSS_ROW}")


def test_member1_tier2_tax_formula(wb):
    ws = wb["Calculator"]
    col = "B"   # member 1
    assert ws[f"{col}{C.TIER2_TAX_ROW}"].value == F.tier2_tax_formula(
        f"{col}{C.NET_EARNINGS_ROW}", f"{col}{C.BAND2_ROW}",
        f"{col}{C.TSB_REF_ROW}", "t1_sel")


def test_fund_total_is_sum_of_member_tax(wb):
    ws = wb["Calculator"]
    val = ws[f"B{C.FUND_TOTAL_ROW}"].value
    assert val == f"=SUM(B{C.TOTAL_TAX_ROW}:E{C.TOTAL_TAX_ROW})"


def test_sample_fund_seeded(wb):
    ws = wb["Calculator"]
    # Alice (col B) 60% / $4M closing; Bob (col C) 40% / $12.9M closing.
    assert ws[f"B{C.CLOSING_TSB_ROW}"].value == 4_000_000
    assert ws[f"C{C.CLOSING_TSB_ROW}"].value == 12_900_000
    assert ws[f"B{C.SHARE_ROW}"].value == 0.6
    assert ws[f"C{C.SHARE_ROW}"].value == 0.4
    assert ws[f"B{C.DIVIDENDS_ROW}"].value == 120_000


def test_sheet_protected_with_unlocked_inputs(wb):
    ws = wb["Calculator"]
    assert ws.protection.sheet is True
    # A seeded input cell is unlocked; a result cell stays locked.
    assert ws[f"B{C.CLOSING_TSB_ROW}"].protection.locked is False
    assert ws[f"B{C.TOTAL_TAX_ROW}"].protection.locked is True
```

- [ ] **Step 2: Run to verify it fails**

Run: `pytest tests/div296_calc/test_calc_build.py -q`
Expected: FAIL — `ModuleNotFoundError` for `div296_calc.build` (built next task) / `div296_calc.tabs.calculator`.

> The build entrypoint is Task 8. To keep this task self-checking, temporarily run only the calculator-tab pieces after Task 8 wires `build_workbook`. If executing strictly task-by-task, expect this test file to stay red until the end of Task 8, then go green — that is the intended red→green seam between Tasks 7 and 8.

- [ ] **Step 3: Write the implementation**

`src/div296_calc/tabs/calculator.py`:
```python
"""Calculator tab — the single data-entry + output sheet.

Layout (members as columns B–E; labels in col A):

  Row 1   Title
  Row 2   Sample-data badge
  Row 3   Income year (input B3) + threshold-loaded status (D3)
  Row 5   Band: 1. Fund pooled income
  Row 6   Dividends/distributions (grossed-up)      [input]
  Row 7   Interest                                  [input]
  Row 8   Rent                                      [input]
  Row 9   Other realised income                     [input]
  Row 10  Net realised capital gain (from helper)   [formula]
  Row 11  less: Deductible expenses                 [input]
  Row 12  POOLED TOTAL                              [formula]
  Row 14  Band: 2. CGT netting helper
  Row 15  Gross gains held >12m                     [input]
  Row 16  Gross gains held <12m                     [input]
  Row 17  Capital losses (current + brought-fwd)    [input]
  Row 18  (spare)
  Row 19  Unused capital loss carried forward       [formula output]
  Row 21  Band: 3. Members
  Row 22  Member name        (B–E)                  [input]
  Row 23  Opening TSB        (B–E)                   [input]
  Row 24  Closing TSB        (B–E)                   [input]
  Row 25  Share % of pooled earnings (B–E)          [input]
  Row 26  Prior-year Div 296 loss   (B–E)           [input]
  Row 27  Earnings override (optional) (B–E)        [input]
  Row 28  Earnings used      (B–E)                  [formula]
  Row 29  TSB used (ref)     (B–E)                  [formula]
  Row 30  Net Div 296 earnings (B–E)                [formula]
  Row 31  band1 proportion   (B–E)                  [formula]
  Row 32  band2 proportion   (B–E)                  [formula]
  Row 33  Tier-1 tax (15%)   (B–E)                  [formula]
  Row 34  Tier-2 tax (extra) (B–E)                  [formula]
  Row 35  TOTAL Division 296 tax (B–E)              [formula]
  Row 36  New carried-forward loss (B–E)            [formula]
  Row 37  Status             (B–E)                  [formula]
  Row 38  (hidden) pooled-share contribution (B–E)  [helper]
  Row 39  Pooled-share sum / count                  [helper]
  Row 40  FUND TOTAL Division 296 tax               [formula]
  Row 42-44  Liability block (3 plain lines)

  Hidden cols O–R: year table  (O=year, P=t1, Q=t2, R=greater_of)
  Resolved selectors (named): t1_sel, t2_sel, greater_of_sel, and the
  rate/discount constants live in col B of an assumptions strip (rows 46-48).
"""

from __future__ import annotations

from openpyxl.styles import Font, PatternFill, Protection
from openpyxl.utils import get_column_letter
from openpyxl.workbook import Workbook
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.formatting.rule import FormulaRule

from div296_calc import _formulas as F
from div296_calc import named_ranges as nr
from div296_calc.assumptions import (
    DISCOUNT_RATE, MEMBER_COUNT, RATE_TIER1, RATE_TIER2, YEAR_TABLE,
)
from div296.styles import (
    BODY_FONT, FMT_CURRENCY, FMT_PERCENT, FMT_TEXT,
    INPUT_FILL, INPUT_FONT, SECTION_BAND_FILL, SECTION_BAND_FONT,
    THIN_BOX, TITLE_FONT,
)

SHEET = "Calculator"

# --- row constants (single source of truth) ---
YEAR_ROW = 3
DIVIDENDS_ROW = 6
INTEREST_ROW = 7
RENT_ROW = 8
OTHER_ROW = 9
NET_CG_ROW = 10
EXPENSES_ROW = 11
POOLED_TOTAL_ROW = 12
CGT_OVER_ROW = 15
CGT_UNDER_ROW = 16
CGT_LOSS_ROW = 17
UNUSED_LOSS_ROW = 19
NAME_ROW = 22
OPENING_TSB_ROW = 23
CLOSING_TSB_ROW = 24
SHARE_ROW = 25
PRIOR_LOSS_ROW = 26
OVERRIDE_ROW = 27
EARNINGS_ROW = 28
TSB_REF_ROW = 29
NET_EARNINGS_ROW = 30
BAND1_ROW = 31
BAND2_ROW = 32
TIER1_TAX_ROW = 33
TIER2_TAX_ROW = 34
TOTAL_TAX_ROW = 35
NEW_LOSS_ROW = 36
STATUS_ROW = 37
POOLED_CONTRIB_ROW = 38
SHARE_GUARD_ROW = 39
FUND_TOTAL_ROW = 40
LIABILITY_ROW = 42

ASSUMP_FIRST_ROW = 46          # rate_tier1, rate_tier2, discount_rate strip
YEAR_TABLE_FIRST_ROW = 2       # hidden cols O–R start at row 2 (row 1 = headers)
YEAR_TABLE_LAST_ROW = 11       # capacity for ~10 future year rows

# Member columns B..E (1 + MEMBER_COUNT-1).
MEMBER_COLS = [get_column_letter(2 + i) for i in range(MEMBER_COUNT)]   # B,C,D,E

# Seeded coherent sample fund (a single pool; shares sum to 100%).
SAMPLE_POOLED = {
    DIVIDENDS_ROW: 120_000, INTEREST_ROW: 30_000, RENT_ROW: 80_000,
    OTHER_ROW: 0, EXPENSES_ROW: 40_000,
}
SAMPLE_CGT = {CGT_OVER_ROW: 300_000, CGT_UNDER_ROW: 0, CGT_LOSS_ROW: 60_000}
# Alice (B): pooled 60%, $4M.  Bob (C): pooled 40%, $12.9M (two-tier anchor).
SAMPLE_MEMBERS = {
    "B": {NAME_ROW: "Alice", OPENING_TSB_ROW: 3_800_000, CLOSING_TSB_ROW: 4_000_000,
          SHARE_ROW: 0.6, PRIOR_LOSS_ROW: 0},
    "C": {NAME_ROW: "Bob", OPENING_TSB_ROW: 12_500_000, CLOSING_TSB_ROW: 12_900_000,
          SHARE_ROW: 0.4, PRIOR_LOSS_ROW: 0},
}


def _band(ws, row, text):
    ws.cell(row=row, column=1, value=text).font = SECTION_BAND_FONT
    ws.merge_cells(f"A{row}:E{row}")
    for col in range(1, 6):
        ws.cell(row=row, column=col).fill = SECTION_BAND_FILL


def _label(ws, row, text):
    ws.cell(row=row, column=1, value=text).font = BODY_FONT


def _input(ws, coord, value=None, number_format=FMT_CURRENCY):
    cell = ws[coord]
    if value is not None:
        cell.value = value
    cell.font = INPUT_FONT
    cell.fill = INPUT_FILL
    cell.border = THIN_BOX
    cell.protection = Protection(locked=False)
    cell.number_format = number_format


def _formula(ws, coord, value, number_format=FMT_CURRENCY):
    cell = ws[coord]
    cell.value = value
    cell.number_format = number_format
    cell.font = Font(name="Arial", size=10, italic=True, color="1D3B34")


def _define_name(wb, name, coord):
    wb.defined_names[name] = DefinedName(
        name=name, attr_text=f"'{SHEET}'!${coord[0]}${coord[1:]}")


def build(wb: Workbook) -> Worksheet:
    ws = wb.create_sheet(SHEET)

    # --- Title ---
    ws["A1"] = "Ongoing Division 296 Calculator"
    ws["A1"].font = TITLE_FONT
    ws.merge_cells("A1:E1")

    # --- Hidden year table (cols O–R) ---
    ws["O1"] = "Year"; ws["P1"] = "T1"; ws["Q1"] = "T2"; ws["R1"] = "GreaterOf"
    for i, (year, yt) in enumerate(sorted(YEAR_TABLE.items())):
        r = YEAR_TABLE_FIRST_ROW + i
        ws[f"O{r}"] = year
        ws[f"P{r}"] = yt.threshold_1
        ws[f"Q{r}"] = yt.threshold_2
        ws[f"R{r}"] = 1 if yt.use_greater_of else 0
    for col in ("O", "P", "Q", "R"):
        ws.column_dimensions[col].hidden = True

    year_rng_o = f"$O${YEAR_TABLE_FIRST_ROW}:$O${YEAR_TABLE_LAST_ROW}"
    t1_rng = f"$P${YEAR_TABLE_FIRST_ROW}:$P${YEAR_TABLE_LAST_ROW}"
    t2_rng = f"$Q${YEAR_TABLE_FIRST_ROW}:$Q${YEAR_TABLE_LAST_ROW}"
    gof_rng = f"$R${YEAR_TABLE_FIRST_ROW}:$R${YEAR_TABLE_LAST_ROW}"

    # --- Row 3: income year selector + resolved thresholds ---
    _label(ws, YEAR_ROW, "Income year")
    _input(ws, f"B{YEAR_ROW}", value="2026-27", number_format=FMT_TEXT)
    # constrain to table years
    # NOTE: a range-reference list DV takes the bare range (no leading '='),
    # mirroring openpyxl's idiom; the '="Yes,No"' form is only for literal lists.
    dv = DataValidation(type="list", formula1=year_rng_o, allow_blank=False,
                        showErrorMessage=True, errorTitle="Unknown income year",
                        error="Pick a year present in the threshold table (cols O–R).")
    ws.add_data_validation(dv)
    dv.add(f"B{YEAR_ROW}")
    ws[f"D{YEAR_ROW}"] = F.year_known_guard_formula(year_rng_o, f"$B${YEAR_ROW}")
    ws[f"D{YEAR_ROW}"].font = Font(name="Arial", size=10, bold=True, color="666666")

    # resolved selectors (named) — placed in hidden helper cells P/Q/R col S
    ws["S2"] = F.threshold_lookup_formula(t1_rng, year_rng_o, f"$B${YEAR_ROW}")
    ws["S3"] = F.threshold_lookup_formula(t2_rng, year_rng_o, f"$B${YEAR_ROW}")
    ws["S4"] = F.threshold_lookup_formula(gof_rng, year_rng_o, f"$B${YEAR_ROW}")
    ws.column_dimensions["S"].hidden = True
    _define_name(wb, nr.T1_SEL, "S2")
    _define_name(wb, nr.T2_SEL, "S3")
    _define_name(wb, nr.GREATER_OF_SEL, "S4")

    # --- Assumptions strip (rows 46-48): rate constants as named cells ---
    for offset, (name, label, value, fmt) in enumerate([
        (nr.RATE_TIER1, "Div 296 rate — tier 1 ($3m–$10m)", RATE_TIER1, FMT_PERCENT),
        (nr.RATE_TIER2, "Div 296 rate — tier 2 (above $10m)", RATE_TIER2, FMT_PERCENT),
        (nr.DISCOUNT_RATE, "CGT discount (1/3)", DISCOUNT_RATE, "0.000%"),
    ]):
        r = ASSUMP_FIRST_ROW + offset
        _label(ws, r, label)
        _input(ws, f"B{r}", value=value, number_format=fmt)
        _define_name(wb, name, f"B{r}")

    # --- Band 1: pooled income ---
    _band(ws, 5, "1. Fund pooled income (realised, before fund 15% tax; exclude contributions)")
    for row, text in [
        (DIVIDENDS_ROW, "Dividends / distributions (grossed-up, incl. franking)"),
        (INTEREST_ROW, "Interest"), (RENT_ROW, "Rent"),
        (OTHER_ROW, "Other realised income"),
    ]:
        _label(ws, row, text)
        _input(ws, f"B{row}", value=SAMPLE_POOLED.get(row), number_format=FMT_CURRENCY)
    _label(ws, NET_CG_ROW, "Net realised capital gain (from CGT helper below)")
    _formula(ws, f"B{NET_CG_ROW}",
             F.cgt_net_formula(f"B{CGT_OVER_ROW}", f"B{CGT_UNDER_ROW}", f"B{CGT_LOSS_ROW}"))
    _label(ws, EXPENSES_ROW, "less: Deductible expenses")
    _input(ws, f"B{EXPENSES_ROW}", value=SAMPLE_POOLED.get(EXPENSES_ROW), number_format=FMT_CURRENCY)
    _label(ws, POOLED_TOTAL_ROW, "POOLED TOTAL")
    _formula(ws, f"B{POOLED_TOTAL_ROW}", F.pooled_total_formula(
        f"B{DIVIDENDS_ROW}", f"B{INTEREST_ROW}", f"B{RENT_ROW}",
        f"B{OTHER_ROW}", f"B{NET_CG_ROW}", f"B{EXPENSES_ROW}"))
    ws[f"B{POOLED_TOTAL_ROW}"].font = Font(name="Arial", size=11, bold=True, color="0F6E56")

    # --- Band 2: CGT helper ---
    _band(ws, 14, "2. CGT netting helper (raw figures off a CLASS/BGL report)")
    for row, text, val in [
        (CGT_OVER_ROW, "Gross capital gains — held > 12 months (discountable)", SAMPLE_CGT[CGT_OVER_ROW]),
        (CGT_UNDER_ROW, "Gross capital gains — held < 12 months", SAMPLE_CGT[CGT_UNDER_ROW]),
        (CGT_LOSS_ROW, "Capital losses (current year + brought-forward)", SAMPLE_CGT[CGT_LOSS_ROW]),
    ]:
        _label(ws, row, text)
        _input(ws, f"B{row}", value=val, number_format=FMT_CURRENCY)
    _label(ws, UNUSED_LOSS_ROW, "→ Unused capital loss carried forward (to NEXT year's losses)")
    _formula(ws, f"B{UNUSED_LOSS_ROW}", F.cgt_unused_loss_formula(
        f"B{CGT_OVER_ROW}", f"B{CGT_UNDER_ROW}", f"B{CGT_LOSS_ROW}"))

    # --- Band 3: members (cols B–E) ---
    _band(ws, 21, "3. Members (enter up to 4; segregated members use the override row)")
    row_labels = {
        NAME_ROW: "Member name", OPENING_TSB_ROW: "Opening TSB (1 Jul, ALL super)",
        CLOSING_TSB_ROW: "Closing TSB (30 Jun, ALL super)",
        SHARE_ROW: "Share % of pooled earnings", PRIOR_LOSS_ROW: "Prior-year Div 296 loss",
        OVERRIDE_ROW: "Earnings override (segregated; optional)",
        EARNINGS_ROW: "Earnings used", TSB_REF_ROW: "TSB used (ref)",
        NET_EARNINGS_ROW: "Net Div 296 earnings", BAND1_ROW: "Proportion $3m–$10m",
        BAND2_ROW: "Proportion above $10m", TIER1_TAX_ROW: "Tier-1 tax (15%)",
        TIER2_TAX_ROW: "Tier-2 tax (extra 10%)", TOTAL_TAX_ROW: "TOTAL Division 296 tax",
        NEW_LOSS_ROW: "New carried-forward Div 296 loss", STATUS_ROW: "Status",
    }
    for row, text in row_labels.items():
        _label(ws, row, text)

    for col in MEMBER_COLS:
        seed = SAMPLE_MEMBERS.get(col, {})
        _input(ws, f"{col}{NAME_ROW}", value=seed.get(NAME_ROW), number_format=FMT_TEXT)
        _input(ws, f"{col}{OPENING_TSB_ROW}", value=seed.get(OPENING_TSB_ROW), number_format=FMT_CURRENCY)
        _input(ws, f"{col}{CLOSING_TSB_ROW}", value=seed.get(CLOSING_TSB_ROW), number_format=FMT_CURRENCY)
        _input(ws, f"{col}{SHARE_ROW}", value=seed.get(SHARE_ROW), number_format=FMT_PERCENT)
        _input(ws, f"{col}{PRIOR_LOSS_ROW}", value=seed.get(PRIOR_LOSS_ROW), number_format=FMT_CURRENCY)
        _input(ws, f"{col}{OVERRIDE_ROW}", value=None, number_format=FMT_CURRENCY)

        _formula(ws, f"{col}{EARNINGS_ROW}", F.earnings_formula(
            f"{col}{NAME_ROW}", f"{col}{OVERRIDE_ROW}", f"{col}{SHARE_ROW}",
            f"$B${POOLED_TOTAL_ROW}"))
        _formula(ws, f"{col}{TSB_REF_ROW}", F.tsb_ref_formula(
            f"{col}{OPENING_TSB_ROW}", f"{col}{CLOSING_TSB_ROW}", nr.GREATER_OF_SEL))
        _formula(ws, f"{col}{NET_EARNINGS_ROW}", F.net_earnings_formula(
            f"{col}{EARNINGS_ROW}", f"{col}{PRIOR_LOSS_ROW}"))
        _formula(ws, f"{col}{BAND1_ROW}", F.band1_formula(
            f"{col}{TSB_REF_ROW}", nr.T1_SEL, nr.T2_SEL), number_format=FMT_PERCENT)
        _formula(ws, f"{col}{BAND2_ROW}", F.band2_formula(
            f"{col}{TSB_REF_ROW}", nr.T2_SEL), number_format=FMT_PERCENT)
        _formula(ws, f"{col}{TIER1_TAX_ROW}", F.tier1_tax_formula(
            f"{col}{NET_EARNINGS_ROW}", f"{col}{BAND1_ROW}", f"{col}{TSB_REF_ROW}", nr.T1_SEL))
        _formula(ws, f"{col}{TIER2_TAX_ROW}", F.tier2_tax_formula(
            f"{col}{NET_EARNINGS_ROW}", f"{col}{BAND2_ROW}", f"{col}{TSB_REF_ROW}", nr.T1_SEL))
        _formula(ws, f"{col}{TOTAL_TAX_ROW}", f"=SUM({col}{TIER1_TAX_ROW},{col}{TIER2_TAX_ROW})")
        ws[f"{col}{TOTAL_TAX_ROW}"].font = Font(name="Arial", size=11, bold=True, color="0F6E56")
        _formula(ws, f"{col}{NEW_LOSS_ROW}", F.new_loss_formula(f"{col}{NET_EARNINGS_ROW}"))
        _formula(ws, f"{col}{STATUS_ROW}", F.status_formula(
            f"{col}{TSB_REF_ROW}", nr.T1_SEL), number_format=FMT_TEXT)
        # hidden pooled-share contribution helper
        _formula(ws, f"{col}{POOLED_CONTRIB_ROW}", F.pooled_share_contrib_formula(
            f"{col}{NAME_ROW}", f"{col}{OVERRIDE_ROW}", f"{col}{SHARE_ROW}"),
            number_format=FMT_PERCENT)
    ws.row_dimensions[POOLED_CONTRIB_ROW].hidden = True

    # --- Share guard ---
    last_col = MEMBER_COLS[-1]
    _label(ws, SHARE_GUARD_ROW, "Pooled-share check")
    ws[f"B{SHARE_GUARD_ROW}"] = F.pooled_count_formula(
        f"B{NAME_ROW}:{last_col}{NAME_ROW}", f"B{OVERRIDE_ROW}:{last_col}{OVERRIDE_ROW}")
    ws[f"C{SHARE_GUARD_ROW}"] = f"=SUM(B{POOLED_CONTRIB_ROW}:{last_col}{POOLED_CONTRIB_ROW})"
    ws[f"C{SHARE_GUARD_ROW}"].number_format = FMT_PERCENT
    ws[f"D{SHARE_GUARD_ROW}"] = F.share_guard_status_formula(
        f"B{SHARE_GUARD_ROW}", f"C{SHARE_GUARD_ROW}")
    ws.conditional_formatting.add(
        f"D{SHARE_GUARD_ROW}",
        FormulaRule(formula=[f'ISNUMBER(SEARCH("⚠",D{SHARE_GUARD_ROW}))'],
                    fill=PatternFill("solid", fgColor="FBE9E9")))
    ws.conditional_formatting.add(
        f"D{SHARE_GUARD_ROW}",
        FormulaRule(formula=[f'ISNUMBER(SEARCH("✓",D{SHARE_GUARD_ROW}))'],
                    fill=PatternFill("solid", fgColor="E1F5EE")))

    # --- Fund total ---
    _label(ws, FUND_TOTAL_ROW, "FUND TOTAL Division 296 tax")
    ws[f"B{FUND_TOTAL_ROW}"] = f"=SUM(B{TOTAL_TAX_ROW}:{last_col}{TOTAL_TAX_ROW})"
    ws[f"B{FUND_TOTAL_ROW}"].number_format = FMT_CURRENCY
    ws[f"B{FUND_TOTAL_ROW}"].font = Font(name="Arial", size=12, bold=True, color="0F6E56")

    # --- Liability block (3 plain lines) ---
    for i, line in enumerate([
        "Division 296 is assessed to the INDIVIDUAL, on top of the fund's own 15% tax.",
        "Payable personally OR by election to release from a super fund (release-authority deadline applies).",
        "Two carry-forwards: the member Div 296 loss (row 36) rolls into next year's prior-loss; the unused capital loss (row 19) rolls into next year's CGT-helper losses.",
    ]):
        c = ws.cell(row=LIABILITY_ROW + i, column=1, value=line)
        c.font = Font(name="Arial", size=9, italic=True, color="666666")
        ws.merge_cells(f"A{LIABILITY_ROW + i}:E{LIABILITY_ROW + i}")

    # --- Sample-data badge ---
    badge = ws.cell(row=2, column=1, value=(
        f'=IF(AND(B{DIVIDENDS_ROW}={SAMPLE_POOLED[DIVIDENDS_ROW]},'
        f'B{CLOSING_TSB_ROW}={SAMPLE_MEMBERS["B"][CLOSING_TSB_ROW]}),'
        f'"⚠ Sample data preloaded — overwrite with your fund\'s figures before sharing.","")'))
    badge.font = Font(name="Arial", size=10, bold=True, italic=True, color="8A6D00")
    badge.fill = PatternFill("solid", fgColor="FFF4CE")
    ws.merge_cells("A2:E2")

    # --- Column widths, protection, print ---
    ws.column_dimensions["A"].width = 46
    for col in MEMBER_COLS:
        ws.column_dimensions[col].width = 18
    ws.freeze_panes = "B4"

    ws.protection.sheet = True
    ws.protection.formatColumns = False
    ws.protection.formatRows = False
    ws.protection.selectLockedCells = False
    ws.protection.selectUnlockedCells = False

    ws.oddHeader.center.text = "ILLUSTRATIVE — NOT ADVICE"
    ws.oddHeader.center.size = 28
    ws.oddHeader.center.color = "CCCCCC"
    ws.page_setup.orientation = "portrait"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.print_area = f"A1:E{LIABILITY_ROW + 2}"
    return ws
```

- [ ] **Step 4: Run after Task 8 wires the build**

(See Task 8 Step 4 — `test_calc_build.py` goes green once `build_workbook` exists.)

- [ ] **Step 5: Commit**

```bash
git add src/div296_calc/tabs/calculator.py tests/div296_calc/test_calc_build.py
git commit -m "feat(calc): Calculator tab (members-as-columns) projecting the pure layer"
```

---

### Task 8: Notes tab + build entrypoint + strict recalc gate

**Files:**
- Create: `src/div296_calc/tabs/notes.py`
- Create: `src/div296_calc/build.py`

- [ ] **Step 1: Write the failing test**

Add to `tests/div296_calc/test_calc_build.py`:
```python
def test_notes_tab_has_law_basis_and_caveats(wb):
    text = "\n".join(
        str(c.value) for row in wb["Notes"].iter_rows() for c in row if c.value)
    assert "Royal Assent" in text
    assert "thresholds verified" in text.lower()
    assert "not advice" in text.lower() or "not personal advice" in text.lower()
    assert "carry-forward" in text.lower() or "carried forward" in text.lower()


def test_build_main_writes_file(tmp_path):
    from div296_calc.build import main
    out = tmp_path / "calc.xlsx"
    rc = main(["-o", str(out), "--no-validate"])
    assert rc == 0
    assert out.exists()
```

- [ ] **Step 2: Run to verify it fails**

Run: `pytest tests/div296_calc/test_calc_build.py -q`
Expected: FAIL — `ModuleNotFoundError` for `div296_calc.build`.

- [ ] **Step 3: Write the implementation**

`src/div296_calc/tabs/notes.py`:
```python
"""Notes / disclaimer tab."""

from __future__ import annotations

from openpyxl.styles import Alignment, Font
from openpyxl.workbook import Workbook
from openpyxl.worksheet.worksheet import Worksheet

from div296_calc.assumptions import YEAR_TABLE
from div296.styles import SECTION_BAND_FILL, SECTION_BAND_FONT, TITLE_FONT

SHEET = "Notes"
THRESHOLDS_VERIFIED = "2026-06-20"   # update when YEAR_TABLE changes

LINES = [
    ("band", "Division 296 — ongoing calculator notes"),
    ("p", "Law basis: Treasury Laws Amendment (Building a Stronger and Fairer "
          "Super System) Act 2026 — Royal Assent 13 March 2026. First income "
          "year 2026-27 (first test date 30 June 2027)."),
    ("band", "Thresholds by year (user-maintained)"),
    ("p", f"thresholds verified {THRESHOLDS_VERIFIED}. At v0.1 only 2026-27 is "
          "confirmed ($3M / $10M, un-indexed). Later years are CPI-indexed "
          "($150k / $500k steps) — add each year's ATO-published figures to the "
          "hidden year table on the Calculator tab (cols O–R) as released."),
    ("p", "Years currently loaded: " + ", ".join(sorted(YEAR_TABLE))),
    ("band", "Earnings basis"),
    ("p", "Enacted method (approximated here): earnings ≈ taxable income − "
          "assessable contributions + net ECPI − NALI, measured BEFORE the "
          "fund's own 15% tax. v0.1 takes realised income components (less "
          "deductible expenses) directly. Contributions are NOT added back and "
          "there is NO $3M loss-floor (both were 2023-draft features)."),
    ("p", "Franking: dividends should be the grossed-up amount (cash + franking "
          "credits) per the fund return; exact treatment under final "
          "regulations to be confirmed."),
    ("p", "CGT discount caveat: the helper applies the 1/3 discount to ALL "
          "realised gains held >12 months regardless of pension/accumulation "
          "phase — an approximation pending final regs."),
    ("band", "Reset cost base"),
    ("p", "For 2026-27 onward, realised capital gains should be measured from "
          "the 30-Jun-2026 reset cost base IF the fund made that election — see "
          "the separate year-one CGT-reset tool."),
    ("band", "Two carry-forwards — do not conflate"),
    ("p", "Member Division 296 loss (Calculator row 36) → next year's per-member "
          "prior-loss input. Unused capital loss (Calculator row 19) → next "
          "year's CGT-helper capital-losses input."),
    ("band", "Scope & disclaimer"),
    ("p", "Accumulation and account-based pensions only; one calc per member "
          "across ALL their super (not per fund). Defined-benefit interests, "
          "NALI, and multi-fund aggregation are out of scope. Outputs are "
          "ESTIMATES subject to final regulations — not personal advice."),
]


def build(wb: Workbook) -> Worksheet:
    ws = wb.create_sheet(SHEET)
    r = 1
    for kind, text in LINES:
        if kind == "band":
            c = ws.cell(row=r, column=1, value=text)
            c.font = SECTION_BAND_FONT if r > 1 else TITLE_FONT
            if r > 1:
                c.fill = SECTION_BAND_FILL
            ws.merge_cells(f"A{r}:H{r}")
        else:
            c = ws.cell(row=r, column=1, value=text)
            c.font = Font(name="Arial", size=10)
            c.alignment = Alignment(wrap_text=True, vertical="top")
            ws.merge_cells(f"A{r}:H{r}")
            ws.row_dimensions[r].height = 46
        r += 1

    ws.column_dimensions["A"].width = 20
    ws.oddHeader.center.text = "ILLUSTRATIVE — NOT ADVICE"
    ws.oddHeader.center.size = 28
    ws.oddHeader.center.color = "CCCCCC"
    ws.page_setup.orientation = "portrait"
    ws.page_setup.fitToWidth = 1
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    return ws
```

`src/div296_calc/build.py`:
```python
"""Build the Ongoing Division 296 Calculator workbook.

    python -m div296_calc.build
    python -m div296_calc.build --no-validate

Writes dist/Division_296_Calculator_v<version>.xlsx (Calculator, Notes).
The recalc gate is STRICT: ANY Excel error cell fails the build (no
skip-list). The workbook is small enough that the pure-Python `formulas`
engine recalculates it without the OOM the year-one model hits.
"""

from __future__ import annotations

import argparse
import re
from pathlib import Path

from openpyxl import Workbook

from div296_calc import __version__
from div296_calc.tabs import calculator, notes

EXCEL_ERROR_RE = re.compile(r"#(REF|DIV/0|VALUE|NAME|NULL|NUM|N/A)!?\b")
AUTHOR = "Aiden Hiew"
TITLE = "Ongoing Division 296 Calculator"


def _dist_dir() -> Path:
    return Path(__file__).resolve().parents[2] / "dist"


def build_workbook() -> Workbook:
    wb = Workbook()
    wb.remove(wb.active)
    calculator.build(wb)
    notes.build(wb)
    props = wb.properties
    props.creator = props.lastModifiedBy = AUTHOR
    props.title = props.subject = TITLE
    props.description = f"v{__version__} — ongoing Div 296 calculator"
    for ws in wb.worksheets:
        ws.oddFooter.left.text = f"Prepared by {AUTHOR}"
        ws.oddFooter.left.size = 8
        ws.oddFooter.right.text = f"v{__version__}  |  Page &P of &N"
        ws.oddFooter.right.size = 8
    return wb


def validate_recalc(xlsx_path: Path) -> list[str]:
    """Strict recalc — return every cell that resolves to an Excel error."""
    import formulas  # noqa: PLC0415

    xl = formulas.ExcelModel().loads(str(xlsx_path)).finish()
    sol = xl.calculate()
    errors: list[str] = []
    for key, cell in sol.items():
        if "'!" not in str(key):
            continue
        try:
            value = cell.value
        except Exception as exc:  # noqa: BLE001
            errors.append(f"{key} -> <unreadable: {exc!r}>")
            continue
        if EXCEL_ERROR_RE.search(str(value)):
            errors.append(f"{key} -> {value}")
    return errors


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Build the ongoing Div 296 calculator.")
    parser.add_argument("--output", "-o", type=Path, default=None)
    parser.add_argument("--no-validate", action="store_true")
    args = parser.parse_args(argv)

    out = args.output or (_dist_dir() / f"Division_296_Calculator_v{__version__}.xlsx")
    out.parent.mkdir(parents=True, exist_ok=True)
    build_workbook().save(out)
    print(f"Wrote {out}")

    if args.no_validate:
        return 0
    try:
        errors = validate_recalc(out)
    except ImportError:
        print("Skipped recalc: `formulas` not installed (pip install -e .[dev]).")
        return 0
    if errors:
        print(f"FAILED recalc — {len(errors)} error cell(s):")
        for e in errors[:20]:
            print(f"  {e}")
        return 1
    print("Recalc validation: OK (no Excel error cells).")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
```

- [ ] **Step 4: Run the full fast suite for the package**

Run: `pytest tests/div296_calc/ -q -m "not slow"`
Expected: PASS — `test_calc_build.py` (incl. Task 7's tests) now green; all engine/formula/assumptions tests green.

- [ ] **Step 5: Build once locally and eyeball**

Run: `python -m div296_calc.build`
Expected: prints `Wrote dist/Division_296_Calculator_v0.1.0.xlsx` then `Recalc validation: OK (no Excel error cells).` (exit 0). If a cell errors, fix the formula before continuing.

- [ ] **Step 6: Commit**

```bash
git add src/div296_calc/tabs/notes.py src/div296_calc/build.py tests/div296_calc/test_calc_build.py
git commit -m "feat(calc): Notes tab + build entrypoint with STRICT recalc gate"
```

---

### Task 9: Integration test (slow) — sample fund recalc

**Files:**
- Create: `tests/div296_calc/test_calc_integration.py`

- [ ] **Step 1: Write the failing test**

`tests/div296_calc/test_calc_integration.py`:
```python
"""Slow integration: build → recalc with `formulas` → assert sample-fund cells.

The shipped coherent sample fund (Alice pooled 60% / $4M; Bob pooled 40% /
$12.9M; one pool) is the live-workbook truth source. The four kernel goldens
(Jess/Emma/Jill/Jamal) are pure-layer unit tests in test_calc_engine.py — the
two-tier path is exercised live here via Bob's >$10M balance.

Hand-derivation:
  CGT helper: gains>12m 300k, losses 60k → net = (300k−60k)×2/3 = 160,000
  POOLED TOTAL = 120k+30k+80k+0+160k − 40k = 350,000
  Alice: earnings 0.6×350k = 210,000; TSB_REF 4.0M; band1 = 1/4 = .25
         tax = 210,000 × .25 × .15 = 7,875
  Bob:   earnings 0.4×350k = 140,000; TSB_REF 12.9M
         band1 = 7/12.9, band2 = 2.9/12.9
         tax = 140,000 × (7/12.9×.15 + 2.9/12.9×.25) ≈ 19,263.57
  FUND TOTAL ≈ 27,138.57
"""

from __future__ import annotations

import numpy as np
import pytest

from div296_calc.build import build_workbook
from div296_calc.tabs import calculator as C

formulas = pytest.importorskip(
    "formulas", reason="`formulas` required — pip install -e .[dev]")
pytestmark = pytest.mark.slow


def _unwrap(v):
    if isinstance(v, np.ndarray):
        v = v.flat[0]
    if isinstance(v, str):
        return v
    try:
        return float(v)
    except (TypeError, ValueError):
        return v


@pytest.fixture(scope="module")
def sol(tmp_path_factory):
    out = tmp_path_factory.mktemp("calc") / "calc.xlsx"
    build_workbook().save(out)
    xl = formulas.ExcelModel().loads(str(out)).finish()
    s = xl.calculate()
    token = f"[{out.name}]"

    def at(cell):
        return _unwrap(s[f"'{token}CALCULATOR'!{cell}"].value)

    return at


def _r(x):
    assert isinstance(x, (int, float)), f"non-numeric: {x!r}"
    return int(round(x))


def test_net_capital_gain_cell(sol):
    assert _r(sol(f"B{C.NET_CG_ROW}")) == 160_000


def test_pooled_total_cell(sol):
    assert _r(sol(f"B{C.POOLED_TOTAL_ROW}")) == 350_000


def test_alice_total_tax(sol):
    assert _r(sol(f"B{C.TOTAL_TAX_ROW}")) == 7_875


def test_bob_two_tier_total_tax(sol):
    assert _r(sol(f"C{C.TOTAL_TAX_ROW}")) == 19_264
    assert _r(sol(f"C{C.TIER2_TAX_ROW}")) > 0       # >$10M slice taxed live


def test_fund_total_tax(sol):
    assert _r(sol(f"B{C.FUND_TOTAL_ROW}")) == 27_139


def test_recalc_gate_zero_error_cells(tmp_path):
    from div296_calc.build import build_workbook, validate_recalc
    out = tmp_path / "calc.xlsx"
    build_workbook().save(out)
    assert validate_recalc(out) == []
```

- [ ] **Step 2: Run to verify it fails (or errors before impl is correct)**

Run: `pytest tests/div296_calc/test_calc_integration.py -q`
Expected: initially may FAIL if any sample formula is off. Fix `calculator.py`/`_formulas.py` until all six pass. (If `formulas` isn't installed it SKIPS — install with `pip install -e .[dev]`.)

- [ ] **Step 3: Make it pass**

No new production code if Tasks 6–8 are correct. If a value is wrong, the most likely culprits: a `+`-chain that should be `SUM`, a band referencing the wrong threshold name, or the year-table SUMIFS resolving 0 (check `t1_sel`/`t2_sel` named ranges point at `S2`/`S3`). Fix and re-run.

- [ ] **Step 4: Run to verify it passes**

Run: `pytest tests/div296_calc/test_calc_integration.py -q`
Expected: PASS (6 passed) — or SKIPPED if `formulas` absent.

- [ ] **Step 5: Commit**

```bash
git add tests/div296_calc/test_calc_integration.py
git commit -m "test(calc): slow integration — live recalc of the sample fund + zero-error gate"
```

---

### Task 10: CI wiring + final gate

**Files:**
- Modify: `.github/workflows/ci.yml`

- [ ] **Step 1: Add the strict build step for the new package**

In `.github/workflows/ci.yml`, immediately AFTER the existing `Build (recalc validation skipped on CI ...)` step (the `python -m div296.build --no-validate` step), insert:
```yaml
      - name: Build ongoing calculator (STRICT recalc — small enough to validate)
        run: python -m div296_calc.build
```
The existing `Lint (ruff)` step already runs `ruff check src tests scripts` (covers `div296_calc` + new tests). The existing `Test (pytest ... -m "not slow")` step already covers `tests/div296_calc/` fast tests. Only the strict build line is new.

- [ ] **Step 2: Run the full local CI-equivalent gate**

Run, in order (this is exactly what CI runs, per the `run_ci_gate_locally_before_push` memory):
```bash
ruff check src tests scripts
pytest -q -n auto -m "not slow"
python -m div296.build --no-validate
python -m div296_calc.build
```
Expected: ruff clean; all fast tests pass; both builds exit 0 (the new one prints `Recalc validation: OK`).

> **Lint caveat (the code in this plan is illustrative, not pre-linted):** ruff (line-length 100) will likely flag a few `E501` long lines in `calculator.py`/`tabs/notes.py` (label dicts, comment banners, the liability/Notes prose) and any `F401` unused import. Wrap/fix exactly what ruff reports before committing — do **not** widen the line-length or add blanket `noqa`. This is mechanical and expected; the formula/engine logic is unaffected.

- [ ] **Step 3: Run the slow suite once locally (CI skips it)**

Run: `pytest tests/div296_calc/test_calc_integration.py -q`
Expected: PASS (or SKIP if `formulas`/`numpy` missing).

- [ ] **Step 4: Commit**

```bash
git add .github/workflows/ci.yml
git commit -m "ci(calc): build the ongoing calculator with the strict recalc gate"
```

- [ ] **Step 5: Show the pre-push summary (do NOT push without Aiden's go-ahead)**

Per the global commit/push workflow, output: `git status`, `git diff --stat` (vs branch point), a one-paragraph summary, the test results from Step 2–3, and known risks. The branch is `div296-ongoing-calc/v0.1` (a feature branch — push allowed once Aiden approves; this plan does not push).

---

## Self-Review

**1. Spec coverage**
- §2 single-year/4-member/per-member+roll-up → Tasks 4–8. ✓
- §2 pooled less expenses + override → Task 4, `pooled_total`/`member_earnings`. ✓
- §2 CGT netting helper → Task 3 + Calculator rows 15–19. ✓
- §2 two tiers / year table / greater-of / carry-forward-as-IO → Tasks 1, 2, 5. ✓
- §2 sample pre-filled + clear path → Task 7 sample seed + badge (clear = overwrite/delete input cells; documented in liability/notes). ✓
- §2/§4 user-maintained threshold table, only 2026-27, unknown-year actionable → Task 1 `UnknownYearError`, Task 7 `year_known_guard_formula` + DV. ✓
- §4 pre-fund-tax basis label → Task 7 band-1 header text. ✓
- §4 all-overridden guard suppression → Task 5 `pooled_share_status`, Task 6 `share_guard_status_formula`, Task 7 wiring. ✓
- §5 slice-form kernel + footgun anchor → Task 2 (Emma) + Task 9 (live Bob >$10M). ✓
- §6 distinct carry-forwards labelled → Task 7 liability line + Task 8 Notes. ✓
- §7 SUM-not-+, guard-first IF, SUMIFS-not-INDEX/MATCH, locking, print, strict gate → Task 6 builders + Tasks 7–8. ✓
- §8 Notes content (law, verified-date, franking, discount caveat, reset link, carry-forwards, scope) → Task 8. ✓
- §9 four goldens (kernel) + boundary/path/CGT/guard + integration → Tasks 2–5, 9. ✓
- §10 TDD order → Tasks 1→9 follow it. ✓
- §3 CI builds+tests + strict build → Task 10. ✓

**2. Placeholder scan:** No "TBD"/"handle edge cases"/"similar to". Every code step shows full code. ✓

**3. Type/name consistency:** `member_div296_tax`, `tsb_ref`, `net_capital_gain`, `pooled_total`, `member_earnings`, `compute_member`, `compute_fund`, `fund_total_tax`, `pooled_share_status` consistent across tasks. Named ranges `t1_sel`/`t2_sel`/`greater_of_sel`/`rate_tier1`/`rate_tier2`/`discount_rate` consistent between `named_ranges.py`, `_formulas.py` goldens, and `calculator.py`. Row constants (`POOLED_TOTAL_ROW` etc.) referenced identically in `calculator.py` and both build/integration tests. ✓

One known seam (intentional, called out in-plan): `test_calc_build.py` is authored in Task 7 but only goes green after Task 8 creates `build.py`.

---

## Execution Handoff

Plan complete. Two execution options:

**1. Subagent-Driven (recommended)** — a fresh subagent per task, review between tasks, fast iteration.
**2. Inline Execution** — execute tasks in this session via executing-plans, batch with checkpoints.
