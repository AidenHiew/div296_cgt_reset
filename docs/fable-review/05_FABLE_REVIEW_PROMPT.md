# Fable Review — Your Task & Required Output

> This is your actual assignment. Files `00`–`04` are your briefing; this file is
> what to *do* and how to *deliver* it. Read `00`–`04` first.

---

## Your role

You are a senior reviewer with fresh eyes on the Division 296 project (three
surfaces, `main` @ `e184725`). You have been brought in specifically for
**design, presentation, and product-shape judgement** — the maintainer values an
independent, opinionated perspective and has explicitly invited you to **challenge**
his choices. Be direct. Praise what's genuinely good, name what isn't, and back
every claim with a specific file/line or a concrete user scenario.

## Hard constraints (from `00` — do not violate)

- **Do not modify code.** Deliverable is a written report.
- **Law/methodology is fixed.** Don't propose changing the tax math. Optimise *how*
  it's computed/presented. (If you think you've found a genuine correctness bug,
  flag it as a boxed **⚠ CORRECTNESS** callout — separate from design suggestions —
  and reproduce it against the anchors in `03`.)
- **Medium/architecture is open.** You may propose moving between mediums, merging
  or killing surfaces, a shared engine, etc.
- **Concentrate on the PRIMARY surface (quota-constrained).** Go **deep on the
  ongoing calculator (`src/div296_calc/`)** + the cross-cutting product-shape
  question. The reset workbook and web get a **light/optional** pass only if budget
  remains — don't spend scarce review budget deep-diving the mature reset tool. (See
  `00` §Scope for the full priority order; it supersedes the earlier equal-weight
  plan.)
- **Correctness is a sanity pass**, already well-reviewed. Reproduce **Emma =
  $115,581.40** as your kernel sanity check, spot-check the invariants in `03`, then
  spend your depth on design.
- **Don't relitigate settled scope decisions** (`00`) or re-report the team's own
  known-deferred items (`web/PROGRESS.md`) — build on them.

## Efficiency & seeing the output

Use `04_CODE_READING_GUIDE.md` — read the ★★★ files, skim ★★, and skip the whole
tree, the generated `standalone/` bundle, and the full test suites. **Before you
critique Excel presentation, read `04`'s "Seeing the rendered output" note:** build
the workbook and inspect the actual `.xlsx` (and, where you can, the PDF) — judging a
tearsheet from its openpyxl builder alone will mislead you. Scope any visual verdict
to what you could actually verify.

---

## What to produce — required report structure

Write a single markdown report. Use exactly these sections.

### 1. Executive summary (½–1 page)
Your headline judgement. If a reader reads only this, what should they know? Include
your **top 5 recommendations** as a ranked list, each one line, each tagged with its
disruption tier (T0/T1/T2 — defined below).

### 2. PRIMARY — Ongoing calculator (`src/div296_calc/`), in depth
This is the bulk of the report. Cover:
- **What it is & who it's for** (one paragraph — confirm you understood it).
- **Engine sanity** (short — did the math check out at a sanity level? reproduce
  Emma; any ⚠ CORRECTNESS flags?).
- **Presentation critique** — the substance. Is a single-sheet Excel workbook the
  right medium for a tool a fund will run **every year from 2027**? Judge the
  members-as-columns layout, the manual year-table upkeep, the manual carry-forward
  roll, the 4-member cap, the hidden-column mechanics. What confuses, what's
  over-built, what's missing, what breaks at year 2? Ground it in specific files and
  the built `.xlsx` (see `04`).
- **Concrete improvements** — a prioritised list, each tagged T0/T1/T2.

### 2b. LIGHT / OPTIONAL — Reset workbook + Web (only if budget remains)
A **brief** pass each (a few paragraphs, not the full treatment above): standout
strengths, standout problems, and any T0/T1/T2 items worth recording. The reset tool
is mature — don't re-derive its design history; just flag what a fresh eye sees.
**If you're low on budget, skip these and say so** — the ongoing tool + §3 matter
more.

### 3. Cross-cutting findings
The things only visible across all three. At minimum, take a clear position on:
- **The triple-maintained math / drift risk** (reset `calcs.py` → web `calcs.js`
  hand-port; ongoing standalone; three duplicated assumption/fixture sets). Is a
  single-source-of-truth architecture worth it here, or is hand-port + parity tests
  fine at this scale? Recommend, with reasoning.
- **Product shape / audience.** Should the public **web** calculator expose the
  **ongoing** tax (the annual reality from 2027), not just the one-off 2026 reset?
  Should these three converge toward one product, or stay separate tools for
  separate jobs? Make the call.
- **Excel-as-medium.** How much of the workbook complexity (hidden helpers, layout-
  constant fragility, recalc-gate blind spots, print-density) is inherent to "an
  adviser must open and audit it in Excel," vs. accidental — and what would a better
  medium change?
- **Docs coherence.** The root README/docs don't describe the monorepo's three
  deliverables. (Cheap fix — note it.)

### 4. Optimisation / redesign plan (the headline deliverable)
A concrete, actionable plan, **tiered by disruption** so the maintainer can adopt
per item:
- **Tier 0 — in-place:** tweaks within a surface as-is (relabels, layout, a11y
  fixes, wiring the web parity test into CI, etc.).
- **Tier 1 — medium/architecture:** move or merge surfaces, introduce a shared
  engine / codegen, add a web front-end for the ongoing tool, etc.
- **Tier 2 — greenfield:** a from-scratch product shape, **only if you believe one
  is warranted** — make the argument, don't just assert it.

For **each** item use this format:

> **[T0/T1/T2] Title**
> **Problem:** … **Proposal:** … **Payoff:** … **Cost/risk:** …
> **Lands in:** `path/to/file` (or "new")

Order items by your recommended sequence (what to do first). Call out dependencies
between items. Where you'd want the maintainer's decision before proceeding, say so.

### 5. What's already good (be honest here too)
Don't only critique. Name the genuinely strong choices (e.g. Python-builder as
source of truth, the fail-loud year guard, the golden-string tests, the sign-colour
conventions) so the maintainer knows what NOT to throw away in a redesign.

---

## Tone & standard

- Opinionated, specific, honest. Challenge freely; you were asked to.
- Every criticism carries a concrete fix or a clear "here's why it's wrong."
- Distinguish **evidence** (what the code does) from **judgement** (what you'd
  change) — the maintainer wants to see your reasoning, not just conclusions.
- No performative hedging and no false praise. If something's fine, say "this is
  fine, leave it."

## Deliverable

Return the report as markdown. (If run as a subagent, your final message IS the
report.) Suggested save location if writing to disk:
`docs/fable-review/FINDINGS.md`.
