// @ts-nocheck
/*
 * ui.ts — SPFx-scoped copy of web/app.js.
 *
 * Identical logic to the website, with three adaptations for running inside a
 * SharePoint web part:
 *   1. DOM lookups are scoped to a ROOT element (the web part's domElement),
 *      not document, so multiple web parts / the SharePoint page can't collide.
 *   2. Elements that live outside the calculator section on the website
 *      (countdown, hero-sample, download/footer version) are guarded — they
 *      simply no-op when absent.
 *   3. Exports initCalculator(root) instead of auto-running on load.
 *
 * Keep in lockstep with web/app.js.
 */
import {
  ASSUMPTIONS,
  SAMPLE_REGISTER,
  SAMPLE_MEMBERS,
  computeComparison,
} from "./calcs";

const VALUATION_DATE = new Date("2026-06-30T23:59:59+10:00"); // AEST
const SAMPLE_CODES = new Set(["P1", "S1", "L1"]);
const SAMPLE_TSBS = new Set([12_000_000, 3_500_000]);

// ROOT is set by initCalculator(); el() is scoped to it (1).
let ROOT: Document | HTMLElement = document;
const el = (id) => ROOT.querySelector("#" + id);

const state = {
  members: SAMPLE_MEMBERS.map((m) => ({ ...m })),
  assets: SAMPLE_REGISTER.map((a) => ({ ...a })),
};

const fmtMoney = (n) =>
  (n < 0 ? "(" : "") +
  "$" +
  Math.abs(Math.round(n)).toLocaleString("en-AU") +
  (n < 0 ? ")" : "");
const fmtSigned = (n) => (n > 0 ? "+" : "") + fmtMoney(n);
const fmtPct = (n) => (n * 100).toFixed((n * 100) % 1 === 0 ? 0 : 1) + "%";
const signClass = (n) => (n < 0 ? "neg" : n > 0 ? "pos" : "");
const valSignClass = (n) => (n < -0.5 ? "val-neg" : n > 0.5 ? "val-pos" : "");

function renderCountdown() {
  const node = el("countdown-days");
  if (!node) return; // (2) countdown lives in the hero; absent in the web part
  const ms = +VALUATION_DATE - +new Date();
  if (ms <= 0) {
    node.textContent = "0";
    el("countdown").querySelector(".countdown-label").textContent =
      "30 June 2026 (the reset valuation date) has passed";
  } else {
    node.textContent = String(Math.max(0, Math.ceil(ms / 86_400_000)));
  }
}

function numInput(value, onInput) {
  const inp = document.createElement("input");
  inp.type = "number";
  inp.value = value;
  inp.setAttribute("inputmode", "decimal");
  inp.addEventListener("input", (e) => {
    onInput(parseFloat((e.target as HTMLInputElement).value) || 0);
    render();
  });
  return inp;
}

function textInput(value, onInput) {
  const inp = document.createElement("input");
  inp.type = "text";
  inp.value = value;
  inp.addEventListener("input", (e) =>
    onInput((e.target as HTMLInputElement).value)
  );
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

function renderMembers() {
  const body = el("members-body");
  body.innerHTML = "";
  state.members.forEach((m, i) => {
    const tr = document.createElement("tr");
    tr.appendChild(cell(`Member ${i + 1}`));
    tr.appendChild(cell(numInput(m.tsb, (v) => (m.tsb = v))));
    tr.appendChild(
      cell(numInput(+(m.split_pct * 100).toFixed(4), (v) => (m.split_pct = v / 100)))
    );
    tr.appendChild(
      cell(state.members.length > 1 ? removeBtn(() => state.members.splice(i, 1)) : "")
    );
    body.appendChild(tr);
  });

  const sum = state.members.reduce((s, m) => s + m.split_pct, 0);
  const hint = el("split-hint");
  if (Math.abs(sum - 1) > 0.005) {
    hint.innerHTML = `⚠ Earnings splits sum to <strong>${fmtPct(sum)}</strong> — they normally total 100%.`;
    hint.style.color = "var(--red)";
  } else {
    hint.textContent = "Earnings splits total 100%. ✓";
    hint.style.color = "var(--muted)";
  }
}

function renderAssets() {
  const body = el("assets-body");
  body.innerHTML = "";
  state.assets.forEach((a, i) => {
    const tr = document.createElement("tr");
    if (a.current_market_value < a.original_cost_base) tr.classList.add("loss-row");

    tr.appendChild(cell(textInput(a.code, (v) => (a.code = v))));
    tr.appendChild(cell(numInput(a.original_cost_base, (v) => (a.original_cost_base = v))));
    tr.appendChild(
      cell(
        numInput(a.market_value_30jun2026, (v) => {
          a.market_value_30jun2026 = v;
          a.current_market_value = v;
        })
      )
    );
    tr.appendChild(
      cell(numInput(a.projected_sale_proceeds, (v) => (a.projected_sale_proceeds = v)))
    );

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
      cell(state.assets.length > 1 ? removeBtn(() => state.assets.splice(i, 1)) : "")
    );
    body.appendChild(tr);
  });
}

function fundFig(k, v) {
  return `<div class="fund-fig"><div class="k">${k}</div><div class="v">${v}</div></div>`;
}

function renderResults() {
  const cmp = computeComparison(state.assets, state.members, ASSUMPTIONS);

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

  const hero = el("hero-sample"); // (2) hero card absent in the web part
  if (hero) {
    hero.innerHTML = `
      <div class="hero-row"><span class="k">If no reset</span><span class="v">${fmtMoney(cmp.noReset.headline)}</span></div>
      <div class="hero-row"><span class="k">If elected to reset</span><span class="v">${fmtMoney(cmp.elected.headline)}</span></div>
      <div class="hero-row headline"><span class="k">Difference</span><span class="v ${valSignClass(cmp.headlineDifference)}">${fmtSigned(cmp.headlineDifference)}</span></div>`;
  }

  const pmBody = el("permember-body");
  pmBody.innerHTML = "";
  state.members.forEach((m, i) => {
    const no = cmp.noReset.perMember[i].tax;
    const elx = cmp.elected.perMember[i].tax;
    const d = elx - no;
    const tr = document.createElement("tr");
    tr.innerHTML = `<td>Member ${i + 1}</td><td>${fmtMoney(no)}</td><td>${fmtMoney(elx)}</td><td class="${valSignClass(d)}">${fmtSigned(d)}</td>`;
    pmBody.appendChild(tr);
  });

  const impactBody = el("asset-impact-body");
  impactBody.innerHTML = "";
  [...cmp.perAssetDiff]
    .sort((x, y) => Math.abs(y.difference) - Math.abs(x.difference))
    .slice(0, 10)
    .forEach((row) => {
      const tr = document.createElement("tr");
      if (row.trap) tr.classList.add("trap-row");
      const label = (row.asset.code || "—") + (row.trap ? ' <span class="trap-tag">TRAP</span>' : "");
      tr.innerHTML = `<td>${label}</td><td>${fmtMoney(row.noResetTax)}</td><td>${fmtMoney(row.electedTax)}</td><td class="${valSignClass(row.difference)}">${fmtSigned(row.difference)}</td>`;
      impactBody.appendChild(tr);
    });

  el("fund-figures").innerHTML = `
    ${fundFig("Fund ordinary CGT", fmtMoney(cmp.noReset.ordCgt))}
    ${fundFig("Div 296 earnings — no reset", fmtMoney(cmp.noReset.earnings))}
    ${fundFig("Div 296 earnings — elected", fmtMoney(cmp.elected.earnings))}
    ${fundFig("Carry-forward losses", fmtMoney(cmp.carryForward))}`;
}

function renderSampleBadge() {
  const present =
    state.assets.some((a) => SAMPLE_CODES.has((a.code || "").trim())) ||
    state.members.some((m) => SAMPLE_TSBS.has(m.tsb));
  el("sample-badge").hidden = !present;
}

function renderAppliesBanner() {
  const t1 = ASSUMPTIONS.threshold_1;
  const above = state.members.filter((m) => m.tsb > t1).length;
  const banner = el("applies-banner");
  if (above === 0) {
    banner.className = "applies-banner is-clear";
    banner.innerHTML = `✓ <strong>Division 296 doesn't apply to this fund.</strong> No member's total super balance exceeds $${t1.toLocaleString("en-AU")}, so the Div 296 figures below are $0 either way. The reset election makes no difference here.`;
  } else {
    banner.className = "applies-banner is-applies";
    banner.innerHTML = `● <strong>Division 296 applies.</strong> ${above} member${above > 1 ? "s have" : " has"} a total super balance above $${t1.toLocaleString("en-AU")} — the reset election is worth modelling.`;
  }
}

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
    <div class="ph-meta">${n} member${n === 1 ? "" : "s"} · total TSB ${fmtMoney(totalTsb)} · reset valuation date 30 June 2026 · prepared ${when}</div>
    <div class="ph-illus">ILLUSTRATIVE — NOT FINANCIAL, TAX OR LEGAL ADVICE</div>`;
}

function render() {
  renderMembers();
  renderAssets();
  renderResults();
  renderSampleBadge();
  renderAppliesBanner();
  renderPrintHeader();
}

function initChrome() {
  renderCountdown();
  // (2) download/footer version chips live outside the calculator section
  const dv = el("download-version");
  if (dv) dv.textContent = "Division 296 Model — v3.4.0";
  const fv = el("footer-version");
  if (fv) fv.textContent = "v3.4.0";

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

/** Entry point called by the web part with its domElement as root. */
export function initCalculator(root: HTMLElement): void {
  ROOT = root;
  initChrome();
  render();
}
