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
const DEADLINE = new Date("2026-06-30T23:59:59+10:00"); // AEST end of 30 Jun 2026
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
  const ms = DEADLINE - now;
  const days = Math.max(0, Math.ceil(ms / 86_400_000));
  const node = el("countdown-days");
  if (ms <= 0) {
    node.textContent = "0";
    el("countdown").querySelector(".countdown-label").textContent =
      "the 30 June 2026 election deadline has passed";
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
    render();
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
    tr.appendChild(
      cell(
        numInput(
          +(m.split_pct * 100).toFixed(4),
          (v) => (m.split_pct = v / 100)
        )
      )
    );
    tr.appendChild(
      cell(
        state.members.length > 1
          ? removeBtn(() => state.members.splice(i, 1))
          : ""
      )
    );
    body.appendChild(tr);
  });

  // split sanity hint
  const sum = state.members.reduce((s, m) => s + m.split_pct, 0);
  const hint = el("split-hint");
  if (Math.abs(sum - 1) > 0.005) {
    hint.innerHTML = `⚠ Earnings splits sum to <strong>${fmtPct(
      sum
    )}</strong> — they normally total 100%.`;
    hint.style.color = "var(--red)";
  } else {
    hint.textContent = "Earnings splits total 100%. ✓";
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

  // hero sample card (always shows the live numbers from current state)
  const hero = el("hero-sample");
  hero.innerHTML = `
    <div class="hero-row"><span class="k">If no reset</span><span class="v">${fmtMoney(
      cmp.noReset.headline
    )}</span></div>
    <div class="hero-row"><span class="k">If elected to reset</span><span class="v">${fmtMoney(
      cmp.elected.headline
    )}</span></div>
    <div class="hero-row headline"><span class="k">Difference</span><span class="v ${valSignClass(
      cmp.headlineDifference
    )}">${fmtSigned(cmp.headlineDifference)}</span></div>`;

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

// ─────────────────────────── orchestration ───────────────────────────
function render() {
  renderMembers();
  renderAssets();
  renderResults();
}

// ── static chrome ──
function initChrome() {
  renderCountdown();
  el("download-version").textContent = `Division 296 Model — ${VERSION}`;
  el("footer-version").textContent = VERSION;

  el("add-member").addEventListener("click", () => {
    state.members.push({ tsb: 0, split_pct: 0 });
    render();
  });
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
