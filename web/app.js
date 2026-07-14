/*
 * app.js — UI for the Division 296 reset calculator.
 * State -> render. All calc lives in calcs.js (the verified engine).
 */

import {
  ASSUMPTIONS,
  SAMPLE_REGISTER,
  SAMPLE_MEMBERS,
  computeComparison,
} from "./calcs.js";

// ── config ──
// 30 Jun 2026 is the reset VALUATION date (market value used for the reset
// cost base), not the election deadline — the election is made with the
// 2026-27 return. The countdown just marks days to this valuation date.
const VALUATION_DATE = new Date("2026-06-30T23:59:59+10:00"); // AEST

// Seeded sample fingerprints (mirror of inputs.py sample-data detection):
// the badge shows while any of these remain, so staff don't share a polished
// tearsheet still built on Mr Sample's figures.
const SAMPLE_CODES = new Set(["P1", "S1", "L1"]);
const SAMPLE_TSBS = new Set([12_000_000, 3_500_000]);
const VERSION = "v3.4.0";

// ── editable state (deep copies of the sample so the originals stay pristine) ──
const state = {
  members: SAMPLE_MEMBERS.map((m) => ({ ...m })),
  assets: SAMPLE_REGISTER.map((a) => ({ ...a })),
};

// ── formatting helpers ──
const fmtMoney = (n) =>
  (n < 0 ? "(" : "") +
  "$" +
  Math.abs(Math.round(n)).toLocaleString("en-AU") +
  (n < 0 ? ")" : "");

const fmtSigned = (n) => (n > 0 ? "+" : "") + fmtMoney(n);

const fmtPct = (n) => (n * 100).toFixed(n * 100 % 1 === 0 ? 0 : 1) + "%";

const signClass = (n) => (n < 0 ? "neg" : n > 0 ? "pos" : "");
const valSignClass = (n) => (n < -0.5 ? "val-neg" : n > 0.5 ? "val-pos" : "");

const el = (id) => document.getElementById(id);

// ─────────────────────────── countdown ───────────────────────────
function renderCountdown() {
  const now = new Date();
  const ms = VALUATION_DATE - now;
  const days = Math.max(0, Math.ceil(ms / 86_400_000));
  const node = el("countdown-days");
  if (ms <= 0) {
    // Valuation date is past: drop the dead day-counter and reframe the chip
    // around the live, actionable deadline — the election is made with the
    // 2026-27 return. (Fable FINDINGS_WEB W2.)
    el("countdown").classList.add("countdown--past");
    node.textContent = "";
    el("countdown").querySelector(".countdown-label").textContent =
      "The reset valuation date (30 June 2026) has passed — model both paths before you lodge your 2026-27 return.";
  } else {
    node.textContent = days;
  }
}

// ─────────────────────────── input builders ───────────────────────────
function numInput(value, onInput, attrs = "") {
  const inp = document.createElement("input");
  inp.type = "number";
  inp.value = value;
  inp.setAttribute("inputmode", "decimal");
  if (attrs) inp.setAttribute(...attrs.split("="));
  inp.addEventListener("input", (e) => {
    onInput(parseFloat(e.target.value) || 0);
    // Value edits refresh only the derived outputs — NOT the input tables.
    // A full render() rebuilds members-body/assets-body via innerHTML, which
    // destroys the very <input> being typed in and drops focus/caret after a
    // single keystroke. (Fable FINDINGS_WEB W3 render-loop / focus-loss.)
    renderOutputs();
  });
  return inp;
}

function textInput(value, onInput) {
  const inp = document.createElement("input");
  inp.type = "text";
  inp.value = value;
  inp.addEventListener("input", (e) => {
    onInput(e.target.value);
  });
  return inp;
}

function removeBtn(onClick) {
  const b = document.createElement("button");
  b.className = "row-remove";
  b.type = "button";
  b.title = "Remove row";
  b.textContent = "×";
  b.addEventListener("click", () => {
    onClick();
    render();
  });
  return b;
}

function cell(child) {
  const td = document.createElement("td");
  if (child instanceof Node) td.appendChild(child);
  else td.textContent = child;
  return td;
}

// ─────────────────────────── members table ───────────────────────────
function renderMembers() {
  const body = el("members-body");
  body.innerHTML = "";
  state.members.forEach((m, i) => {
    const tr = document.createElement("tr");
    tr.appendChild(cell(`Member ${i + 1}`));
    tr.appendChild(cell(numInput(m.tsb, (v) => (m.tsb = v))));
    // Earnings split is derived from each member's share of total balance
    // (matches the Excel workbook — CONTEXT.md "auto-derived from TSB split").
    // Read-only span with a stable id so a TSB edit can refresh it live via
    // updateDerivedSplits() WITHOUT rebuilding the input row (keeps focus).
    const splitSpan = document.createElement("span");
    splitSpan.className = "split-derived";
    splitSpan.id = `split-pct-${i}`;
    splitSpan.textContent = fmtPct(m.split_pct);
    tr.appendChild(cell(splitSpan));
    tr.appendChild(
      cell(
        state.members.length > 1
          ? removeBtn(() => state.members.splice(i, 1))
          : ""
      )
    );
    body.appendChild(tr);
  });
}

// Derive each member's earnings split from their share of total TSB, matching
// the workbook (CONTEXT.md). Mutates state.members[*].split_pct — the single
// source the engine reads. If total TSB is 0, every share is 0.
function syncSplits() {
  const total = state.members.reduce((s, m) => s + (m.tsb || 0), 0);
  state.members.forEach((m) => {
    m.split_pct = total > 0 ? (m.tsb || 0) / total : 0;
  });
}

// Refresh the read-only split-% displays from the (already synced) shares,
// WITHOUT rebuilding the input rows — so editing one member's balance updates
// every member's split live while the caret stays put. (Complements W3.)
function updateDerivedSplits() {
  state.members.forEach((m, i) => {
    const span = el(`split-pct-${i}`);
    if (span) span.textContent = fmtPct(m.split_pct);
  });
}

// Static note: the split is derived, not entered, so it always totals 100%.
function renderSplitHint() {
  const hint = el("split-hint");
  const total = state.members.reduce((s, m) => s + (m.tsb || 0), 0);
  if (total > 0) {
    hint.textContent =
      "Earnings split is derived from each member's share of total balance (totals 100%).";
    hint.style.color = "var(--muted)";
  } else {
    hint.textContent = "Enter member balances to derive the earnings split.";
    hint.style.color = "var(--muted)";
  }
}

// ─────────────────────────── assets table ───────────────────────────
function renderAssets() {
  const body = el("assets-body");
  body.innerHTML = "";
  state.assets.forEach((a, i) => {
    const tr = document.createElement("tr");
    const isLoss = a.current_market_value < a.original_cost_base;
    if (isLoss) tr.classList.add("loss-row");

    tr.appendChild(cell(textInput(a.code, (v) => (a.code = v))));
    tr.appendChild(
      cell(numInput(a.original_cost_base, (v) => (a.original_cost_base = v)))
    );
    // MV @ 30 Jun drives both the reset cost base and the loss flag here.
    tr.appendChild(
      cell(
        numInput(a.market_value_30jun2026, (v) => {
          a.market_value_30jun2026 = v;
          a.current_market_value = v; // keep loss-flag in step with MV input
        })
      )
    );
    tr.appendChild(
      cell(
        numInput(
          a.projected_sale_proceeds,
          (v) => (a.projected_sale_proceeds = v)
        )
      )
    );

    // held > 12m toggle
    const toggle = document.createElement("span");
    toggle.className = "held-toggle";
    toggle.textContent = a.held_over_12_months ? "Yes" : "No";
    toggle.style.color = a.held_over_12_months ? "var(--green)" : "var(--muted)";
    toggle.addEventListener("click", () => {
      a.held_over_12_months = !a.held_over_12_months;
      render();
    });
    tr.appendChild(cell(toggle));

    tr.appendChild(
      cell(
        state.assets.length > 1
          ? removeBtn(() => state.assets.splice(i, 1))
          : ""
      )
    );
    body.appendChild(tr);
  });
}

// ─────────────────────────── results ───────────────────────────
function renderResults() {
  const cmp = computeComparison(state.assets, state.members, ASSUMPTIONS);

  // headline cards
  el("rc-noreset").textContent = fmtMoney(cmp.noReset.headline);
  el("rc-elected").textContent = fmtMoney(cmp.elected.headline);
  const diffNode = el("rc-diff");
  diffNode.textContent = fmtSigned(cmp.headlineDifference);
  diffNode.className = "result-value " + signClass(cmp.headlineDifference);
  el("rc-diff-note").textContent =
    cmp.headlineDifference < 0
      ? "reset SAVES this much Div 296 tax"
      : cmp.headlineDifference > 0
      ? "reset COSTS this much extra"
      : "no difference";

  // per-member breakdown
  const pmBody = el("permember-body");
  pmBody.innerHTML = "";
  state.members.forEach((m, i) => {
    const no = cmp.noReset.perMember[i].tax;
    const el2 = cmp.elected.perMember[i].tax;
    const d = el2 - no;
    const tr = document.createElement("tr");
    tr.innerHTML = `<td>Member ${i + 1}</td>
      <td>${fmtMoney(no)}</td>
      <td>${fmtMoney(el2)}</td>
      <td class="${valSignClass(d)}">${fmtSigned(d)}</td>`;
    pmBody.appendChild(tr);
  });

  // most-affected assets (sorted by |difference| desc, top 10)
  const impactBody = el("asset-impact-body");
  impactBody.innerHTML = "";
  [...cmp.perAssetDiff]
    .sort((x, y) => Math.abs(y.difference) - Math.abs(x.difference))
    .slice(0, 10)
    .forEach((row) => {
      const tr = document.createElement("tr");
      if (row.trap) tr.classList.add("trap-row");
      const label =
        (row.asset.code || "—") +
        (row.trap ? ' <span class="trap-tag">TRAP</span>' : "");
      tr.innerHTML = `<td>${label}</td>
        <td>${fmtMoney(row.noResetTax)}</td>
        <td>${fmtMoney(row.electedTax)}</td>
        <td class="${valSignClass(row.difference)}">${fmtSigned(
        row.difference
      )}</td>`;
      impactBody.appendChild(tr);
    });

  // fund-level supporting figures
  el("fund-figures").innerHTML = `
    ${fundFig("Fund ordinary CGT", fmtMoney(cmp.noReset.ordCgt))}
    ${fundFig("Div 296 earnings — no reset", fmtMoney(cmp.noReset.earnings))}
    ${fundFig("Div 296 earnings — elected", fmtMoney(cmp.elected.earnings))}
    ${fundFig("Carry-forward losses", fmtMoney(cmp.carryForward))}`;
}

function fundFig(k, v) {
  return `<div class="fund-fig"><div class="k">${k}</div><div class="v">${v}</div></div>`;
}

// Hero teaser: frozen to the genuine §12 sample fund (single $12m member) and
// rendered once at load, so the "Sample fund" label stays truthful even after
// the user edits the live calculator below. (Fable FINDINGS_WEB W2 / §2.3.)
function renderHeroSample() {
  const cmp = computeComparison(SAMPLE_REGISTER, SAMPLE_MEMBERS, ASSUMPTIONS);
  el("hero-sample").innerHTML = `
    <div class="hero-row"><span class="k">If no reset</span><span class="v">${fmtMoney(
      cmp.noReset.headline
    )}</span></div>
    <div class="hero-row"><span class="k">If elected to reset</span><span class="v">${fmtMoney(
      cmp.elected.headline
    )}</span></div>
    <div class="hero-row headline"><span class="k">Difference</span><span class="v ${valSignClass(
      cmp.headlineDifference
    )}">${fmtSigned(cmp.headlineDifference)}</span></div>`;
}

// ── #2 sample-data badge ──
function renderSampleBadge() {
  const present =
    state.assets.some((a) => SAMPLE_CODES.has((a.code || "").trim())) ||
    state.members.some((m) => SAMPLE_TSBS.has(m.tsb));
  el("sample-badge").hidden = !present;
}

// ── #1 Div 296 applicability gate ──
function renderAppliesBanner() {
  const t1 = ASSUMPTIONS.threshold_1;
  const above = state.members.filter((m) => m.tsb > t1).length;
  const banner = el("applies-banner");
  if (above === 0) {
    banner.className = "applies-banner is-clear";
    banner.innerHTML = `✓ <strong>Division 296 doesn't apply to this fund.</strong> No member's
      total super balance exceeds $${t1.toLocaleString("en-AU")}, so the Div 296 figures below
      are $0 either way. The reset election makes no difference here.`;
  } else {
    banner.className = "applies-banner is-applies";
    banner.innerHTML = `● <strong>Division 296 applies.</strong> ${above} member${
      above > 1 ? "s have" : " has"
    } a total super balance above $${t1.toLocaleString(
      "en-AU"
    )} — the reset election is worth modelling.`;
  }
}

// ── #3 print tearsheet header ──
function renderPrintHeader() {
  const totalTsb = state.members.reduce((s, m) => s + m.tsb, 0);
  const n = state.members.filter((m) => m.tsb > 0).length;
  const when = new Date().toLocaleDateString("en-AU", {
    day: "numeric",
    month: "long",
    year: "numeric",
  });
  el("print-header").innerHTML = `
    <h1>Division 296 — reset comparison</h1>
    <div class="ph-meta">
      ${n} member${n === 1 ? "" : "s"} · total TSB ${fmtMoney(totalTsb)} ·
      reset valuation date 30 June 2026 · prepared ${when}
    </div>
    <div class="ph-illus">ILLUSTRATIVE — NOT FINANCIAL, TAX OR LEGAL ADVICE</div>`;
}

// ─────────────────────────── orchestration ───────────────────────────
// Derived displays only — safe to call on every keystroke because it never
// rebuilds the members-body / assets-body input tables, so the focused <input>
// survives. (Fable FINDINGS_WEB W3.)
function renderOutputs() {
  syncSplits(); // derive splits from TSB before the engine reads them
  renderResults();
  updateDerivedSplits(); // refresh the read-only split-% spans in place
  renderSplitHint();
  renderSampleBadge();
  renderAppliesBanner();
  renderPrintHeader();
}

// Full render — rebuilds the input tables too. Call only on structural changes
// (add/remove member or asset), never on a value edit.
function render() {
  syncSplits(); // so the split-% spans render with correct values on first paint
  renderMembers();
  renderAssets();
  renderOutputs();
}

// ── static chrome ──
function initChrome() {
  renderCountdown();
  renderHeroSample();
  el("download-version").textContent = `Division 296 Model — ${VERSION}`;
  el("footer-version").textContent = VERSION;

  el("add-member").addEventListener("click", () => {
    state.members.push({ tsb: 0, split_pct: 0 });
    render();
  });
  el("print-btn").addEventListener("click", () => window.print());

  el("add-asset").addEventListener("click", () => {
    state.assets.push({
      code: "NEW",
      name: "New asset",
      quantity: 1,
      original_cost_base: 0,
      current_market_value: 0,
      market_value_30jun2026: 0,
      valuation_source: "",
      projected_sale_proceeds: 0,
      held_over_12_months: true,
    });
    render();
  });
}

initChrome();
render();
