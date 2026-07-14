# Fable Review — FINDINGS_WEB (deep pass on `web/`)

*Reviewer: Fable, second pass. Scope: the web surface only, deep — building on `FINDINGS.md` (§2b web brief pass, §3 product-shape, §4 item 8), not repeating it. Reviewed in the `fable-brief` worktree; site driven live in a browser (localhost server, desktop + 375 px layout viewport). Date: 2026-07-13.*

> **Opus verification (post-review):** spot-checked the cheapest concrete claims against source — `th { white-space: nowrap }` (`web/styles.css:600`, the mobile-overflow mechanism), `--gold: #c8962e` (the failing white-on-gold pairing), `web/README.md:4` literally says "a countdown to the **30 June 2026** election deadline" (regressing `PROGRESS.md`'s own "do not regress" valuation-vs-deadline rule), and the standalone bundle still references `fonts.googleapis.com` twice (not offline-true). All confirmed. The 442 px overflow and 0/14-inputs-unnamed claims are live-DOM measurements Fable ran but I did not independently re-execute; the code structure (default `min-width:auto` on the panels, DOM-built inputs with no `<label>`/aria) is consistent with them.

**Verification note.** Screenshot capture was unavailable in this environment (tool timeout — same constraint as pass 1), so no pixel-level claims. Everything labelled *verified* below was exercised in the live DOM: real `input` events dispatched into the rendered page, element geometry measured at a true 375 px layout viewport (`documentElement.clientWidth = 375`, mobile media queries confirmed matching), accessibility properties read off live nodes, print composition taken from the stylesheet's `@media print` block. Computed styles confirm both webfonts actually load, so typography claims are grounded in the resolved font stack. Items marked *judgement* are design opinion on top of that evidence.

**Not re-reported (known from `FINDINGS.md`):** focus loss per keystroke (re-confirmed live: one dispatched keystroke → `document.activeElement` falls to `BODY`), the stale "From 1 July 2025" lede (`web/index.html:55`), the dead countdown, fixture triplication / parity-not-in-CI. These are assumed queued; several recommendations below bundle with them.

---

## 1. Executive summary

The site's *content* is the best thing about it and the *input loop* is the worst — pass 1 said that, and going deep confirms it, but the deep pass changes the picture in two ways.

First, **the mobile experience is broken outright, not merely "untightened."** At a real 375 px viewport the page acquires **442 px of horizontal overflow** — the whole page, hero to footer, drags sideways — because the calculator panels refuse to shrink below their tables' min-content width and the `.table-scroll` wrappers never engage. It's a one-line CSS fix (`min-width: 0` on the grid items), which makes it the highest payoff-per-character change in the repo. Combined with the focus-loss bug, the calculator is effectively desktop-only *and* mouse-only today; for a public page whose lay audience will substantially arrive on phones, that halves the addressable audience before the content gets a chance.

Second, **the render architecture's problems go beyond focus loss.** The all-or-nothing `render()` produces a second, subtler class of bug I verified live: edits that *don't* re-render leave derived UI stale (the sample-data badge stays up after every sample fingerprint is gone; the impact table keeps the old asset code), while edits that *do* re-render destroy focus. One rebuild-the-world loop, two opposite failure modes. The fix is the same for both — split state updates from output re-rendering — and it should also carry the accessibility repairs, because the current DOM is screen-reader-hostile in a countable way: **0 of 14 calculator inputs have an accessible name**, and the Held >12m control is a `tabIndex=-1` `<span>` no keyboard can reach.

On the strategic question I raised in pass 1 — the web should headline the **ongoing** tax — this pass delivers the concrete design (§4). The short version: the ongoing calculator is a *better* web product than the reset one (a dozen scalars and up to six member rows, no 50-row register), the hosted year-table dissolves the workbook's worst annual hazard by construction (you cannot select a year the served thresholds file doesn't contain), and the site restructure is two pages plus a landing pivot, not a rebuild. No T2 item anywhere; the stack (static HTML + vanilla JS + hand-ported kernel + parity tests) remains right for this project.

---

## 2. New verified findings (evidence layer)

### 2.1 ⚠ Mobile: page-level horizontal overflow — the calculator blows out the whole page

**Verified** at a 375 px layout viewport (mobile media queries confirmed active):

- `document.documentElement.scrollWidth = 817` vs viewport 375 → **442 px of page-level horizontal scroll**.
- Offenders: `.panel` width **792 px**; inside it `#members-table` min-content **750 px** (the `th { white-space: nowrap }` rule at `web/styles.css:600` makes "TOTAL SUPER BALANCE (TSB)" a 295 px unbreakable cell; input `min-width` does the rest).
- The designed containment never engages: every `.table-scroll` measures `clientWidth == scrollWidth` (750/750) — it isn't scrolling because its *parent panel* already grew to fit.

**Root cause:** `.calc-grid` / `.breakdown-grid` items (`web/styles.css:525-540`) keep the CSS-default `min-width: auto`, so each `.panel` grid item's minimum size is its content's min-content width. The overflow container inside the panel can't help; the grid track itself (computed 791.6 px in a 327 px container) overflows, and since nothing clips at the page level, the whole document scrolls sideways.

**Fix (one line):** `.calc-grid > .panel, .breakdown-grid > .panel { min-width: 0; }` — then `.table-scroll` engages exactly as designed. Letting `th` wrap on small screens (or shorter mobile header strings) is the polish on top. *(T0 — see W1.)*

### 2.2 Text edits don't re-render: stale badge and stale labels

**Verified sequence** in the live page: set the member TSB to $5m (numeric → triggers `render()`; badge correctly stays, sample codes still present) → rename all three asset codes to X1/X2/X3 (text inputs; `textInput` at `web/app.js:77-85` updates state but **never calls `render()`**) → **the sample-data badge remains visible although no sample fingerprint remains**, and "Most affected assets" still says **"P1"**. Any later numeric event re-renders and both correct themselves (badge hides, label becomes "X1").

So the render loop has *two* failure modes: numeric edits re-render everything (destroying focus), text edits re-render nothing (leaving the badge — a compliance-flavoured control — and result labels stale). The badge is the same defensive UX the workbook gets right; on the web it silently under-warns. Same root cause, same fix as the focus bug: one `updateOutputs()` path that every input (numeric *and* text) triggers, with table rebuilding reserved for add/remove. *(T0 — bundled into W3.)*

### 2.3 The hero card lies once the user edits anything

**Verified:** the hero card header is static copy — "Sample fund — single member, $12m balance" (`web/index.html:67`) — while its rows are re-rendered from *live state* on every render (`web/app.js:219-228`). After my edits it read "Sample fund — single member, $12m balance / If no reset $66,000". A user who scrolls back up sees their own numbers presented as the sample fund's. Either freeze the card to the genuine sample (compute once from `SAMPLE_*` and never update) or make the label live ("Your fund — 2 members, total TSB $8.2m"). Freezing is truer to its role as a hero teaser. *(T0 — W2.)*

### 2.4 Accessibility inventory (the calculator is screen-reader- and keyboard-hostile)

**Verified on live nodes:**

- **0 of 14** calculator `<input>`s have any accessible name (no `<label>`, `aria-label`, `aria-labelledby`, or `title`). A screen reader announces each as an anonymous "edit, spin button". Column `<th>`s don't help — these are DOM-built inputs inside `<td>`s with no association.
- The **Held >12m toggle** is a `<span>`, `tabIndex = -1`, no `role`, no `aria-pressed`, click-only (`web/app.js:179-187`). Keyboard users cannot change an asset's discount eligibility at all — which changes the tax answer.
- Remove-row buttons are real `<button>`s with `title="Remove row"` (adequate); the results region has `aria-live="polite"` (good); `html lang="en-AU"` (good); heading order is sane (h1 → h2 → h3).
- **Contrast (computed):** white text on `--gold #c8962e` ≈ **2.7:1** — fails WCAG AA at any text size. Used in the countdown pill number (`web/styles.css:179-188`) and the warn-card number badge (`.card-warn .card-num`). Everything else I checked passes (muted `#5c6b73` on white ≈ 5.5:1; green/red value colours ≈ 6:1+; white on teal-700 ≈ 7.6:1).

Fix shape: `aria-label` per input ("Member 2 total super balance", "P1 original cost base" — row context is available at build time in `renderMembers`/`renderAssets`); the toggle becomes a `<button aria-pressed>` styled as today; darken the gold or use dark text on it. *(T0 — W3/W5.)*

### 2.5 Print tearsheet: right concept, three composition gaps

From the `@media print` block (`web/styles.css:872-938`) plus the live print-header content:

- **The sign-convention legend is print-hidden.** `.calc-foot` — the only place "negative (green) = saving / positive (red) = the trap" is explained — is in the hidden list (`web/styles.css:890`). Colour is exactly what dies on a B&W office printer, so the client-facing PDF shows `($109,361)` and `+$8,611` with **no legend anywhere**. The workbook's Comparison tearsheet carries its convention; the web one drops it.
- **No version, no assumptions.** `.site-footer` is hidden, so `v3.4.0` appears nowhere on the printed page; nor do the thresholds/rates used. The print header (verified live) has member count, total TSB, valuation date, prepared date, and the ILLUSTRATIVE banner — good — but a tearsheet that will sit in an email thread should also carry the model version and "thresholds $3m/$10m, rates 15%/25%, 1/3 discount".
- **Inputs are unrecorded.** `.calc-grid` (members + register) is print-hidden, so the PDF shows outputs with no record of the register that produced them — unreproducible three months later. A compact print-only inputs table (or at least per-member TSBs — the per-member table prints tax only) would make the tearsheet self-documenting.
- **Genuinely good and worth protecting:** the sample-data badge and the applicability banner both *survive* into print — a tearsheet built on sample data prints its own warning. That's the workbook's defensive DNA carried over correctly.

*(T0 — W4.)*

### 2.6 Smaller verified items

- **Applicability gate behaves correctly at the edges** (verified): TSB exactly $3,000,000 → "doesn't apply" (strictly-greater, matches the law); below-threshold shows $0 / $0 / $0 with "no difference", and the banner states the election is moot. The gate is one of the site's best UX moves.
- **`web/README.md:4` contradicts the site's own law framing**: "a countdown to the 30 June 2026 **election deadline**" — the valuation-date-vs-election-deadline distinction is the thing `web/PROGRESS.md` calls its "key correctness decision (do not regress)". The regression is sitting in the folder's own README.
- **The "persistent" disclaimer bar isn't.** The HTML comment (`web/index.html:20`) says "Persistent not-advice banner", but `.disclaimer-bar` is `position: static` (verified) — it scrolls away immediately; only the nav header is sticky. Either make it sticky (it's 33 px) or fix the comment; better, give the calculator section its own inline ILLUSTRATIVE chip.
- **No favicon, zero Open Graph / Twitter tags** (verified). Advisers share links on LinkedIn and in emails; today that renders as a bare URL with no preview card. For a page whose distribution model *is* link-sharing, OG tags are disproportionately valuable.
- **Fonts are Google-hosted** (`web/index.html:11-16`), and the "self-contained" standalone bundle still requests `fonts.googleapis.com` (verified by grep of `web/standalone/div296-reset-calculator.html`). The double-click file is not offline-true (degrades to system fonts — acceptable but undermines the promise), and the hosted page makes a third-party request despite "nothing is sent anywhere" copy. Self-hosting two WOFF2 files fixes privacy, offline truthiness, and future CSP in one move.
- **Number-input ergonomics** (part verified, part judgement): values render raw (`12000000`) with no separators — deferred item #5, endorse; spinners are shown and Chrome's focused-scroll-wheel-changes-value behaviour is a live hazard on a page you scroll; no `min`/`max` anywhere (negative cost bases accepted, verified); and `numInput`'s `attrs` parameter (`web/app.js:69`) is dead code no caller uses — the fossil of constraint work that never landed. Format-on-blur, `min="0"` where the domain demands it, and deleting the dead param are one input-polish pass.
- **Version string lives in three places** (`web/app.js:24`, plus hardcoded fallbacks at `web/index.html:341` and `:364`) — harmless today, drift-bait at the next model bump.
- **Anchor clearance is 7 px** (verified: sticky header 65 px, section padding 72 px, `scroll-margin-top: 0`). Nav jumps land the section title a hair under the header — works by luck; one `scroll-margin-top` rule makes it deliberate.

---

## 3. Design / UX / IA critique (judgement layer)

**Information architecture.** The single-page flow — hero → explainer → calculator → trap → download — is right for the *reset* story, with one inversion: the **trap callout sits after the calculator** it's supposed to motivate. The user meets the machine before the reason to care; and since explainer card 3 already plants the trap, the post-calculator section reads as a repeat rather than a reveal. When the site restructures (§4), the trap material becomes the reset page's *header* framing ("why model this before lodging"), not a postscript.

**Mobile IA.** At ≤720 px the nav simply vanishes (`web/styles.css:141-145`, verified `display:none`) — no hamburger, no fallback. On a page measuring ~7,000 px tall with the calculator 2,500 px down, a phone user has no jump link to the thing the hero promises. The hero's "Run the calculator →" button covers the entry path; "Get the workbook" and "The reset trap" become scroll-only. A minimal fix (keep one "Calculator" link visible) beats building a hamburger.

**Interaction model of the calculator.** The editable-table model is right — it mirrors the workbook's register and advisers think in rows. Three judgement calls I'd change:

1. **The split % column puts a footgun where a default belongs.** Most funds' earnings split *is* their balance proportion; free-typing percentages per member (with a warning when they don't sum to 100) invites error. Offer "split pro-rata by TSB (default)" with an override; the warning disappears for the 90% case.
2. **No "reset to sample" / no scenario persistence** (deferred #4 — endorse, fold into W6): once the sample is overwritten there's no way back, and refresh loses everything. URL-hash scenarios are the difference between a toy and an instrument for an adviser modelling three funds in a row.
3. **The band transparency exists and is never shown.** `web/calcs.js` computes `band1`/`band2` per member (`memberBands`, `calcs.js:200-206`) and threads them through `computeComparison` — and `app.js` renders none of it. The workbook's proudest feature is its transparency rows; the web shows headline numbers with no "how". A per-member `<details>` expander ("$12m TSB → 58.3% in the $3m–$10m band, 16.7% above $10m") is already-computed data waiting for markup. This is also deferred item #6 — cheaper than PROGRESS thinks, because the engine already ships the numbers.

**Visual / design system.** Genuinely good at the altitude I can verify: the Fraunces/Inter pairing (both confirmed loading) gives it an editorial, un-SaaS character matching the "prepared by a professional" positioning; the workbook-echoing palette and the load-bearing green/red sign convention carried over intact; `tabular-nums` on every numeric surface; consistent card/border/shadow language. The only outright defect is the white-on-gold contrast failure (§2.4). Softer note: gold does double duty (urgency in the countdown, warning in badge/trap) and the countdown's death removes half of that justification.

**The download section talks to the wrong audience.** The workbook card's *only* acquisition path is `pip install -e .[dev]` + `python -m div296.build` (`web/index.html:346-348`). The workbook's stated audience is advisers/accountants; approximately none will build a Python artifact from source. Either attach the built `.xlsx` to GitHub Releases and link it directly, or be honest that the card is for developers. Right now the site's conversion funnel ends in a `pip` command — the biggest content-to-audience mismatch on the page.

**Trust framing.** The disclaimer posture is right, and the applicability gate actively *talking users out of* the tool is the most credibility-building element on the page. Two updates now that the law is enacted: "Confirm every figure against **the final ATO method**" reads as if the law were still pending — cite the enacted Act (Royal Assent 13 March 2026) and ATO guidance; and the footer's "Prepared by Aiden Hiew" would carry more weight with a role/credential clause and a **"law last reviewed: [date]"** stamp — the cheapest trust signal a tax page can carry, and the natural companion to pass 1's `LAW_BASIS.md` recommendation.

---

## 4. The strategic build: the ongoing-tax web calculator, designed

Pass 1's #1 recommendation, made concrete. The engine is `src/div296_calc/calcs.py` — and note how web-shaped it is: a dozen fund-level scalars, up to six member rows of five scalars each, no asset register. This is a *smaller* build than the reset calculator already shipped.

### 4.1 One site, two tools — the IA

```
/  (index.html — the headline)
├─ Hero: "Division 296 is law. What will it cost each year?"
│   chip: "First test date: 30 June 2027 (2026-27 income year)"
├─ Explainer (rebuilt for the enacted law — see §5)
├─ THE ANNUAL CALCULATOR   ← the ongoing tool, front and centre
├─ "The 2026 transition decision" teaser card → /reset.html
└─ Downloads: BOTH workbooks, audience-labelled

/reset.html  (today's page, reframed)
├─ Banner: "The reset valuation date (30 Jun 2026) has passed. The election
│   itself is made with the fund's 2026-27 return — model it before you lodge."
│   → later: "ARCHIVED — the election window closed with 2026-27 lodgment."
└─ Existing explainer + calculator + trap, unchanged below the fold
```

The reset page keeps its URL and `#calculator` anchor (links are in the wild via the standalone file); the landing page's job flips from "explain the reset" to "explain the tax, compute the annual number, and *offer* the transition page to those who still need it". The countdown dies and its slot becomes the first-test-date chip — a *future* date, which is what a hero urgency element needs to be.

**Positioning sentence for downloads** (the product thesis): *the web pages give you the answer; the workbooks are the audit trail your accountant files.* Reset workbook ↔ reset page; ongoing workbook ↔ annual calculator. Four artifacts, two jobs each.

### 4.2 The annual calculator — inputs and outputs

Direct mapping from the kernel (every input below is a field the Python dataclasses already define — no new modelling):

**Fund panel:**
- **Income year** — `<select>` enumerating *only* the years present in the served thresholds table (2026-27 at launch). Beside it a provenance line: *"Using 2026-27 thresholds: $3.0m / $10m — closing-TSB basis · thresholds last updated 13 Mar 2026."*
- **Realised income:** dividends (grossed, incl. franking) / interest / rent / other / less deductible expenses (`PooledIncome`).
- **Capital gains helper** (collapsible): gross gains held >12m, gross gains ≤12m, capital losses (current + brought-forward) → shows the s102-5 working inline: net realised CG and **unused capital loss to carry forward** (`net_capital_gain`). The workbook buries this netting in rows; the web can show the waterfall in three lines.
- **Pooled total** — displayed, signed, with "may be negative — negative years build each member's carry-forward loss" copy.

**Members panel** (up to **6** — the web should never inherit the workbook's 4-cap):
- name · opening TSB · closing TSB · share % · prior-year Div 296 loss
- **"Segregated?" toggle** revealing an override-earnings field. This dissolves the kernel's subtlest semantic — `override is None` vs `override = 0.0` — into an explicit UI state: toggle off = pooled, toggle on = this number, even zero. The workbook needs a documented convention; the web needs a switch.
- Share-sum guard mirrors `pooled_share_status`: warn when pooled shares ≠ 100%, suppressed when all members are segregated; same pro-rata-by-TSB default as §3.

**Per-member result cards** (not a table — members are the unit of assessment and the unit of conversation):
- Headline: **"Alice — Div 296 assessed: $7,875"** with status chips: `Below $3m — not liable` / `Liable` / `Loss carried forward: $X`.
- **"Show the working" expander** — the workbook's transparency rows, web-native: earnings (share × pool, or segregated) → less prior-year loss → net earnings → TSB reference (with a **"greater of opening/closing"** badge appearing automatically from 2027-28 rows — the GreaterOf semantics the workbook's Notes never explain become a visible, self-explaining label) → band1/band2 → tier 1 (15% on the $3m–$10m slice) + tier 2 (**"25% — the 15% plus the extra 10%"** — bake the footgun's correct reading into the label itself) → total.
- Fund total strip beneath.

**"Take into next year" panel** — the workbook's transcribe-two-cells-by-hand footnote becomes a designed affordance: each member's new Div 296 loss + the fund's unused capital loss, staged with a copy button and a **permalink** ("Save this scenario" — inputs URL-hash-encoded; next year, open the link, bump the year, and the carry-forwards are already on screen). No accounts, no storage, no backend — the URL is the file.

**Print:** per-fund statement and a per-member variant — the "Bob — Division 296 assessed: $19,264; payable personally or via release authority" artifact pass 1 said the workbook lacks. Carry version, income year, thresholds/rates used, prepared date, ILLUSTRATIVE banner, and the working (expanders print expanded).

### 4.3 How hosting dissolves the year-table problem — shown, not asserted

The workbook's worst annual hazard (FINDINGS §2 ⚠) is a *distribution* problem: thresholds change yearly, workbooks are copies, so every copy carries a user-maintained table with hidden columns, protection ceremony, and a silent-zero failure mode on partial rows. On the hosted page every link of that chain disappears **by construction**:

1. **Thresholds ship with the page** — a `thresholds.json` (or const in `calcs-ongoing.js`) generated from `div296_calc/assumptions.py::YEAR_TABLE` by a trivial script; one source of truth, pinned by a parity test so the JS mirror cannot drift from the Python table.
2. **The year selector enumerates that table.** An income year with no thresholds row *cannot be selected* — the `UnknownYearError` / `COUNTIF→NA()` guard class has nothing left to guard. A greyed "2027-28 — thresholds not yet published by the ATO" option communicates the roadmap.
3. **Partial rows can't exist** — rows are authored in the repo, reviewed in a PR, completeness-tested in CI (three lines), not typed into hidden Excel columns by an end user.
4. **Updates deploy centrally.** ATO publishes indexed thresholds → one commit → every user has them on next load. The page can even tell workbook users their copy is stale ("workbook v0.1 covers 2026-27 only").

The residual risk moves where it belongs: the maintainer updating one reviewed file, instead of every user performing an undocumented unprotect-unhide-fill ritual annually.

### 4.4 Engineering pattern (settled in pass 1, restated as constraints)

- `web/calcs-ongoing.js` — hand-ported mirror of `div296_calc/calcs.py` (~150–200 lines; smaller than `calcs.js`). Python stays canonical.
- Fixtures **generated** from the pytest suite (`fixtures.json`), consumed by a node parity test, **in CI from day one** (pass-1 item 4 is the prerequisite — land it for the reset engine first, inherit the pattern).
- **The Emma anchor is mandatory in the JS test**: $115,581.40, with the wrong-pairing value ($87,255.81) asserted *rejected*. Two reviewers have fallen into the slice/cumulative footgun; the third engine must carry the tripwire.
- Rebuild the render loop with targeted updates from the start — do **not** copy `app.js`'s rebuild-the-world `render()`; §2.2 is what it produces.
- Gate: the parked **hosting decision** (GitHub Pages viable; repo public + MIT — maintainer call). Interim: the standalone-bundle channel already proven for the reset page works for this page too.

Cost: days. It remains the highest-leverage item in the portfolio, and nothing in this deep pass weakened the case — mobile users *especially* are annual-number seekers, not register-modellers.

---

## 5. Content & credibility

**What's right already:** the three-card explainer is the best plain-English Div 296 writing I've seen from a non-publisher source; the accountant deep-dive earns adviser trust by showing the mechanics (s102-5 netting, per-asset allocation, an explicit "not modelled" list); the anti-avoidance note (Part IVA / TR 2008/1) is a detail most calculators omit; and the applicability gate tells people the tax *doesn't* apply to them — the most credible thing a tax tool can do.

**The big missing piece — "what changed in the final law."** The public conversation about Div 296 was formed by the 2023 draft: *unrealised* gains taxed, *unindexed* $3m threshold. The enacted Act dropped both — realised-earnings basis, indexed thresholds. A large-balance member arriving at this site probably carries draft-era fears, and the site never corrects them (its lede still carries a draft-era date). A short "The final law vs the 2023 proposal" box — realised basis only · thresholds indexed · commences 2026-27, first test 30 June 2027 · negative years carry forward — would do more for trust than any design change on this list, *and* it's the natural bridge into the ongoing calculator ("so what will it actually cost each year? →").

**Reframe the deadline.** The dead countdown isn't just a stale widget, it's now a framing error: the *live* deadline is that **the election is made with the fund's 2026-27 return** — that's the actionable date, and "model both paths before you lodge" is the correct urgency for 2026-27 lodgment season. The valuation date belongs in the mechanics copy, not the hero.

**For the lay reader, add:** how the tax is *paid* (personally or via release authority — one sentence, currently nowhere); a two-line worked example on the $12m member (§3's expander covers it); a short FAQ ("Is this on gains I haven't realised? — No, the final law…", "Will the $3m threshold rise? — Yes, indexed…", "Does my industry-fund balance count? — TSB is all your super…" — that last one is a real seam: TSB is whole-of-super while the register is this-fund-only, and only the workbook's Notes say so).

**For the adviser, add:** a citation line (the enacted Act, Royal Assent 13 March 2026) and a **"law last reviewed"** date stamp; both belong to the `LAW_BASIS.md` single-source pattern from pass 1 — the page should *render from* that file's facts, not restate them (restating is how "1 July 2025" happened).

**Weight balance:** explainer : calculator : download is a sensible pyramid. The trap gets three placements (card 3, panel hint, full section) — one more than a single page needs; the restructure resolves this by making trap material the reset page's header story.

---

## 6. Prioritised recommendations

Sequenced. W1–W5 are in-place fixes worth doing even if the strategic build never happens; W6 subsumes several content changes.

> **[T0] W1. Fix the mobile blowout (one line of CSS)**
> **Problem:** Verified 442 px page-level horizontal overflow at 375 px; `.table-scroll` never engages because `.panel` grid items keep `min-width: auto`. The calculator — and the whole page around it — is broken on phones.
> **Proposal:** `min-width: 0` on the calc/breakdown grid items; allow `th` wrap (or short mobile header strings) so the members table's min-content isn't 750 px.
> **Payoff:** The public page works on the device half its audience holds. **Cost/risk:** Trivial; re-check the impact tables after.
> **Lands in:** `web/styles.css` (~525-540, ~600); regenerate `web/standalone/`.

> **[T0] W2. Law-and-framing copy batch**
> **Problem:** Beyond the known lede date and dead countdown: the disclaimer still says "final ATO method" as if the law were pending; `web/README.md:4` calls 30 Jun 2026 the "election deadline" (regressing the folder's own documented correctness decision); the hero card shows live numbers under a static "Sample fund" label (§2.3).
> **Proposal:** One pass, sourced from a single law-facts file (pass-1 item 6): lede → enacted 2026-27 commencement; countdown slot → "election is made with the fund's 2026-27 return — model before you lodge"; disclaimers cite the enacted Act; README fixed; hero card frozen to the true sample (or its label made live).
> **Payoff:** The most public surface stops contradicting the enacted law and itself. **Cost/risk:** Trivial.
> **Lands in:** `web/index.html`, `web/app.js`, `web/README.md`; regenerate standalone.

> **[T0] W3. Rebuild the render loop once, fix four bugs at the same address**
> **Problem:** One rebuild-everything `render()` causes focus loss on numeric input (known) *and* stale derived UI on text input (new, §2.2); the same builders emit nameless inputs (0/14) and a keyboard-unreachable `<span>` toggle (§2.4).
> **Proposal:** Split *state update* from *output re-render*: every input (numeric and text) updates state and calls `updateOutputs()` (results, banners, badge, hero, print header); input tables rebuild only on add/remove. While in the builders: `aria-label` per input from row context; toggle → `<button aria-pressed>`; delete the dead `attrs` param; `min="0"` where the domain demands it.
> **Payoff:** Data entry becomes usable — mouse, keyboard, and screen reader; the sample badge stops under-warning. **Cost/risk:** Half a day; engine untouched; verify freshness via the §2.2 sequence.
> **Lands in:** `web/app.js`; regenerate standalone.

> **[T0] W4. Make the print tearsheet self-supporting**
> **Problem:** Printed PDF has signed/coloured figures but the legend is print-hidden; no model version or assumptions on the page; inputs unrecorded (§2.5).
> **Proposal:** Keep a one-line legend in print; add version + thresholds/rates to the print header; add a compact print-only inputs summary (per-member TSBs at minimum). Keep the badge/banner print-survival exactly as is.
> **Payoff:** The artifact that gets emailed around becomes interpretable and reproducible on its own. **Cost/risk:** Small, print-CSS only.
> **Lands in:** `web/styles.css` print block, `web/app.js` (`renderPrintHeader`).

> **[T0] W5. Design-system and distribution odds-and-ends**
> **Problem:** White-on-gold ≈ 2.7:1 (countdown pill, warn-card badge); no favicon; zero OG tags (bare-link sharing); Google-hosted fonts (third-party request; standalone not offline-true); raw unformatted number values; version string in three places; 7 px anchor clearance; mobile nav vanishes entirely.
> **Proposal:** Darken the gold text pairing; favicon + OG tags; self-host the two fonts; format-on-blur thousands separators; single version const; `scroll-margin-top`; keep one "Calculator" link visible on mobile.
> **Payoff:** The page looks professional in the places advisers actually meet it first — link previews, phones, printers. **Cost/risk:** Small, mechanical.
> **Lands in:** `web/index.html`, `web/styles.css`, `web/app.js`; regenerate standalone.

> **[T1] W6. Build the annual (ongoing) calculator page + restructure to "one site, two tools"** — *the strategic item, specified in §4*
> **Problem:** The evergreen annual obligation has no web presence; the site headlines a decision whose valuation date has passed; the workbook's year-table hazard is structural to offline distribution.
> **Proposal:** §4 in full: landing pivots to the enacted-law explainer + annual calculator (6 members, segregated toggle, s102-5 helper, per-member statement cards with "show the working", carry-forward staging + permalinks); reset page reframed then archived; thresholds served from a repo-reviewed table; `calcs-ongoing.js` + generated fixtures + CI parity + the Emma tripwire; targeted-update render loop from day one.
> **Payoff:** Right tax, right audience, right medium; deletes the year-table problem class for web users; gives the project a headline that appreciates rather than expires.
> **Cost/risk:** Days. Gated on the parked hosting decision (**maintainer call**); depends on the fixtures/CI pattern landing first.
> **Lands in:** `web/` (new page, `calcs-ongoing.js`, thresholds data, tests), landing/nav copy.

> **[T1] W7. Scenario permalinks + "reset to sample"** (deferred #4 — endorse, schedule with W6)
> URL-hash encoding is shared infrastructure between the two calculators; build once during W6, retrofit the reset page. Upgrades both tools from demo to instrument for an adviser running multiple funds.

> **[T2] — none.** Static files + vanilla JS + hand-ported kernels remains the right architecture at this scale. The render-loop discipline (W3) is the thing a framework would have given you; adopt the discipline, not the framework. Revisit only if a third calculator page ever makes shared state real.

---

## 7. What's already good about the web (keep these)

- **The prose.** The three-card explainer, the accountant deep-dive, and the trap callout are better plain-English Div 296 material than the professional publishers'. The §4 rebuild should *reuse* this voice.
- **The applicability gate** — verified correct at the boundary, and the copy actively releases people the tax doesn't touch. Port it to the ongoing page ("Below $3m — Div 296 doesn't apply to you; here's the indexation story").
- **Sample-data tripwires that survive into print.** The badge logic needs W3's freshness fix, but the design — a printed tearsheet warning about its own sample data — is defensive UX most tools never think of.
- **The sign convention, everywhere.** Green-saving/red-cost with parenthesised negatives, consistent with the workbook and CONTEXT.md, on `tabular-nums` numerics. Load-bearing; don't let the ongoing page invent a second convention.
- **The typography and palette.** Fraunces + Inter (verified loading) over workbook-echoing sand/teal/gold gives a credible, editorial identity that doesn't look like a SaaS template. Fix the one gold-contrast failure and leave the rest alone.
- **Client-side-only computation with the honesty to say so** ("nothing is sent anywhere") — make it fully true by self-hosting fonts, then keep saying it loudly.
- **The standalone build** — a clever distribution answer while hosting is undecided, with a module-leak guard in the build script (`web/build-standalone.mjs:36-38`). Keep it as the offline artifact even after Pages goes live; regenerate it in CI so it can't drift.
- **`PROGRESS.md` discipline** — the "key correctness decision (do not regress)" note is exactly how a solo project should encode hard-won framing; the fix it needs is that `web/README.md` didn't get the memo.
