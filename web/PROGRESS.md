# Website — progress & memory

Working note for the `web/` site (the Division 296 reset explainer + live
calculator). Branch: `claude/website-brainstorm-915it1`. Keep this current so a
future session can resume without re-reading the whole history.

## What exists

A static, backend-free single-page site in `web/`:

| File | Role |
|------|------|
| `index.html` | Page: persistent not-advice bar, header/nav, hero + countdown, plain-English explainer (3 cards + "for the accountant" deep-dive), live calculator, "why model it first" trap callout, workbook-download CTA, footer. |
| `styles.css` | Styling — echoes the workbook palette (dark-teal titles, sand/slate/sage, gold accent) and the sign-colour convention (muted green = saving, red = cost). Includes the `@media print` tearsheet rules. |
| `calcs.js` | **1:1 port of `src/div296/calcs.py`** (the source of truth). Engine only, no DOM. |
| `app.js` | UI: editable members + asset register → live render; countdown; sample badge; applicability gate; print tearsheet. |
| `tests/parity.test.js` | Pins the §12 acceptance numbers against `calcs.js`. Run: `node web/tests/parity.test.js` → expect **23/23**. |
| `build-standalone.mjs` | Bundles everything into one self-contained HTML (inlines CSS, strips ES-module syntax). Run: `node web/build-standalone.mjs`. |
| `standalone/div296-reset-calculator.html` | The generated single-file build. **Regenerate after any change** to index/styles/calcs/app — it is committed, so don't let it drift. Sent to the user (no repo on their Mac mini) for double-click viewing. |

## Calculator behaviour (verified)

Seeded with the §12 sample fund (single member $12m TSB; P1/S1/L1). On load:
- If no reset **$142,083** · If elected **$32,722** · Difference **($109,361)** (green/saving)
- Fund ordinary CGT **$180,000** · L1 flagged as the **reset trap**

Three workbook-mirrored refinements shipped:
1. **Applicability gate** — banner above results: green "Div 296 doesn't apply" when no member TSB > $3m (figures $0 either way), else "Div 296 applies — N member(s) above $3m".
2. **Sample-data badge** — amber warning while seeded codes (P1/S1/L1) or sample TSBs ($12m / $3.5m) remain; clears once overwritten.
3. **Print / Save-as-PDF tearsheet** — `🖶` button → `window.print()`; `@media print` drops chrome + input panels and renders a clean A4 one-pager (print-only header with date + ILLUSTRATIVE banner, result cards, per-member, most-affected assets, fund figures).

## Key correctness decision (do not regress)

**30 June 2026 is the reset VALUATION date, not the election deadline.** The
election is lodged with the 2026-27 return (in 2027); 30 June 2026 is the date
market value is fixed for the reset cost base. All user-facing copy uses the
valuation-date framing — the countdown, title, meta, hero lede, reset-election
card, and trap callout were all corrected. (Note: `CONTEXT.md` still uses the
older "pre-30-June-2026 election" phrasing — left as-is; flag if it should
change too.)

## Source-of-truth rule

`src/div296/calcs.py` is canonical. `calcs.js` is a hand-maintained mirror —
change both in lockstep and keep `tests/parity.test.js` green so drift is caught.

## Open / deferred (next session)

- **Website hosting / GitHub Pages — user is still thinking about it.** Not
  started. The repo is now public and MIT-licensed (2026), so GitHub Pages is a
  viable host and a Pages site being public is no longer a blocker. Don't deploy
  without an explicit go-ahead.
- Refinements not yet done (user picked 1–3 only so far):
  - #4 Shareable link / save scenario (encode inputs in URL hash) + "reset to sample" button.
  - #5 Thousands separators while typing; tighter mobile layout for tables.
  - #6 Short "how it's calculated" worked example (band1/band2 on the $12m member).
- No PR opened (user hasn't asked).
- No headless browser in the build env → screenshots not possible here; verify visually by opening the standalone file or serving `web/` locally.
