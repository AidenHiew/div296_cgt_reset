# Fable Review — Acceptance Numbers & Invariants

> These are the numbers and rules that **must not change**. The maintainer's
> latitude to you is "medium open, law fixed" — so any redesign you propose has to
> reproduce these exactly. Use them as your correctness sanity check: if a proposed
> change would move any of these, it's out of bounds (or you've found a real bug —
> flag it loudly, don't quietly propose it).

---

## Why this matters for a *design* review

You're allowed to propose big structural changes (new medium, shared engine,
consolidation). The safety rail on all of them is: **the tax outputs are already
correct and pinned**. So your job is to preserve these numbers while improving
everything around them. Treat them as the regression contract for any refactor you
recommend.

---

## Reset / transition workbook (`src/div296/`) — spec §12 acceptance

Sample fund: single member, TSB $12m, 3 assets incl. a $300k loss asset.

| Figure | Value |
|--------|-------|
| Fund Ordinary CGT (after intra-year netting) | **$180,000** |
| Div 296 earnings — no reset | **$1,100,000** |
| Div 296 earnings — elected reset | **$253,333** |
| Div 296 headline tax — no reset | **$142,083** |
| Div 296 headline tax — elected reset | **$32,722** |
| Carry-forward losses (sample) | **$0** (the $2.1m gross gains absorb the $300k loss) |

These are stable from v3.1.0 onward and unchanged through v3.4 + Adviser Edition.
The Adviser Edition (omits the CLASS Import tab) is **byte-identical** to the full
edition on these numbers.

## Ongoing calculator (`src/div296_calc/`) — golden anchors

**Pure-kernel scenario anchors (`tests/div296_calc/test_calc_engine.py`):**

| Scenario | Meaning | Anchor |
|----------|---------|--------|
| **Emma** | TSB > $10m — mandatory two-tier | **$115,581.40** ⚠ the slice-form footgun guard |
| Jess | single-tier ($3m–$10m) | (unit-pinned) |
| Jill | just above $3m | (unit-pinned) |
| Jamal | negative earnings → carry-forward | $0 tax + carry-forward accrues |

**Seeded sample fund (live recalc, `test_calc_integration.py`):** Alice 60% / $4.0m
+ Bob 40% / $12.9m.

| Figure | Value |
|--------|-------|
| Net capital gain (CGT helper) | **$160,000** |
| Pooled total | **$350,000** |
| Alice Div 296 tax | **$7,875** |
| Bob Div 296 tax (exercises >$10m tier) | **≈ $19,264** |
| Fund total | **≈ $27,139** |

## Web calculator (`web/`) — parity contract

`web/tests/parity.test.js` re-derives the **same reset-tool §12 numbers** in JS
(23 checks). The web calculator is a **reset calculator** — it mirrors
`src/div296/calcs.py`, not the ongoing tool. Its numbers must match the reset
workbook's §12 figures above.

---

## Structural invariants (must survive any redesign)

1. **The two-tier slice form pairs band2 with 25%, never 10%.** Reproduce **Emma =
   $115,581.40**. (See `01_DOMAIN_AND_LAW.md` for why this is a footgun.)
2. **Unknown income year must fail LOUD, not silent.** The ongoing tool's threshold
   lookup is wrapped `IF(COUNTIF(year_range,year)=0, NA(), SUMIFS(...))` so an
   unrecognised year propagates `#N/A` rather than emitting a fictitious >$10m-tier
   liability. This was a real bug caught in cold review — do not let any redesign
   reintroduce a silent-zero lookup.
3. **s102-5 netting order:** losses applied to non-discount (short-held) gains
   first, then the 1/3 discount on the surviving long-held slice. Div 296 earnings
   floored at zero.
4. **Per-asset tax reconciles to the headline** — pro-rata of positive-gain share;
   loss assets bear $0 even when they reduce the fund headline.
5. **Sign convention:** Difference `= (if reset) − (if no reset)`; negative = saving
   (green), positive = the reset costs money (red). Load-bearing across the UI.
6. **"ILLUSTRATIVE — NOT ADVICE"** posture everywhere; neutral net-effect framing,
   no "you should…" recommendations.

---

## Existing correctness assurance (why correctness is a sanity pass, not the job)

The engines have already been through:
- v3.4 post-audit adversarial review — 5 read-only subagents, incl. a
  200k-scenario Excel-vs-`calcs.py`-mirror parity check; found no correctness bug.
- Ongoing-calc pre-merge cold review — 3 subagents; caught the real MAJOR
  (unknown-year → fictitious tax), now fixed + regression-tested.
- Golden formula-string tests on both Excel tools (sign-flip tripwires) + the web
  parity test.

So: sanity-check the math, reproduce Emma, but **spend your depth on design,
presentation, and product shape**, not re-auditing arithmetic.
