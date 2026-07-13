# Fable Review — Mandate & Scope

> **You are Fable.** This folder is your complete briefing. Read the files in order
> (00 → 05). You are reviewing the Division 296 project on `main` (tip `e184725`).
> This document defines *what* you are being asked to do and the *hard constraints*
> on your suggestions. `05_FABLE_REVIEW_PROMPT.md` is your actual task list and
> required output format.

---

## Who commissioned this and why

The maintainer (Aiden) wants a **fresh-eyes review** of the whole Division 296
codebase — the engine(s), the calculator(s), and the way it is presented — plus a
concrete **optimisation / redesign plan**. The engine's *correctness* has already
been through several adversarial passes (see `03`), so the highest-value thing you
can add is **design, presentation, and product-shape judgement**, not another
correctness audit. Confirm correctness at a sanity level; spend your depth on
"is this the right shape, and how would you do it better."

This is a review + proposal. **You are not asked to change any code.** Your
deliverable is a written report + optimisation plan (format in `05`).

---

## What exists (three "surfaces")

This is a **monorepo with three independent deliverables** that all compute or
explain Division 296 outcomes:

| # | Surface | Path | Medium | Maturity |
|---|---------|------|--------|----------|
| 1 | **Reset / transition workbook** | `src/div296/` | Python → Excel `.xlsx` | Mature — v3.4 + Adviser Edition |
| 2 | **Ongoing Div 296 calculator** | `src/div296_calc/` | Python → Excel `.xlsx` | New — v0.1 + hardening |
| 3 | **Web live calculator** | `web/` | Static site (HTML/JS/CSS) | Evolving |

They model **two different things** and **two different versions of the law** —
this is the single most important thing to understand before reviewing. See
`01_DOMAIN_AND_LAW.md`.

---

## Scope decisions (already made by the maintainer — do not relitigate)

1. **All three surfaces, weighted equally.** Give each a comparable mid-depth
   pass. Do **not** deep-dive one at the expense of the others. There is no single
   "primary" surface.

2. **Latitude: medium is OPEN, the law is FIXED.**
   - You **may** propose moving between mediums or consolidating surfaces — e.g.
     "the ongoing calculator should be web-first," "these three should share one
     engine," "kill surface X." Nothing about the *delivery medium or architecture*
     is sacred.
   - You **may not** propose changing the **enacted-law calculation methodology**.
     The tax math (two-tier 15% / +25% slice form, thresholds, 1/3 CGT discount,
     s102-5 netting, realised-earnings basis for the ongoing tool) is fixed and
     correct. Optimise *how* it is computed/presented, not *what* is computed.

3. **Correctness is a sanity pass, not the headline.** Already well-reviewed
   (`03`). Flag anything you genuinely find, but the depth budget goes to design /
   presentation / optimisation.

---

## What Fable must produce (exact structure in `05` — follow it)

A single markdown report: an executive summary, a per-surface review (×3),
cross-cutting findings, a **tiered optimisation plan**, and a "what's already good"
section. `05_FABLE_REVIEW_PROMPT.md` is the authoritative spec — use its section
order verbatim.

The optimisation plan is the headline deliverable and is **tiered by disruption** so
the maintainer can adopt per item. This vocabulary (T0/T1/T2) is used throughout:

- **Tier 0 — in-place:** tweaks within a surface as-is.
- **Tier 1 — medium/architecture:** move or merge surfaces, a shared engine, a web
  front-end for the ongoing tool.
- **Tier 2 — greenfield:** a from-scratch product shape — only if warranted, and
  argued, not asserted.

Be opinionated and honest. If something is over-engineered, say so. If a whole
surface is redundant, make the case. If the current approach is actually right, say
so — don't invent work.

---

## Efficiency guidance (the maintainer explicitly cares about resource use)

This briefing package is written so you do **not** need to read the whole tree.
`04_CODE_READING_GUIDE.md` gives you a **curated, ranked file list per surface** plus
a **"Seeing the rendered output"** note — read the ranked files directly, and lean on
the maps in `02` for everything else. Don't exhaustively read tests or frozen `dist/`
artifacts. Spend reads where your design judgement needs ground truth: the calc
kernels and the presentation layers. **Because this is a presentation review, build
the actual artifact where you can (see `04`) — reading a workbook's openpyxl builder
alone will mislead you about how it looks.**
