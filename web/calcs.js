/*
 * calcs.js — browser port of src/div296/calcs.py
 *
 * Pure functions, no DOM. This is a 1:1 mirror of the Python calc engine so
 * the website reproduces the exact §12 acceptance numbers the Excel workbook
 * and pytest suite assert. If calcs.py changes, change this in lockstep — the
 * Python module remains the source of truth (see README "Locked design
 * decisions"). tests/parity.test.js (run under node) pins the §12 numbers.
 *
 * Locked decisions mirrored here (see CONTEXT.md glossary + calcs.py docstring):
 *  - Ordinary CGT at fund level per s102-5 (losses to non-discount gains first,
 *    then 1/3 discount on remaining long-held portion).
 *  - Div 296 fund earnings = MAX(0, sum of adjusted gains) — intra-year netting,
 *    floored at zero.
 *  - Per-asset Div 296 tax = pro-rata of the member-attributed headline,
 *    positive-gain assets only (loss assets bear $0).
 *  - Pension phase NOT modelled (100% accumulation, 15% fund CGT rate).
 */

// ---- Assumptions: single source of truth (mirror of assumptions.py) ----
export const ASSUMPTIONS = Object.freeze({
  rate_tier1: 0.15, // $3m–$10m band
  rate_tier2: 0.25, // above $10m
  threshold_1: 3_000_000,
  threshold_2: 10_000_000,
  indexation_increment_1: 150_000,
  indexation_increment_2: 500_000,
  discount_rate: 1.0 / 3.0, // 1/3 SMSF CGT discount
  fund_cgt_rate: 0.15,
  asset_register_rows: 50,
  member_count: 4,
});

/**
 * @typedef {Object} Asset
 * @property {string} code
 * @property {string} name
 * @property {number} quantity
 * @property {number} original_cost_base
 * @property {number} current_market_value
 * @property {number} market_value_30jun2026
 * @property {string} valuation_source
 * @property {number} projected_sale_proceeds
 * @property {boolean} held_over_12_months
 */

/**
 * @typedef {Object} Member
 * @property {number} tsb
 * @property {number} split_pct   0.0–1.0
 */

// ---- per-asset gain calculations -----------------------------------------

function applyDiscount(rawGain, asset, discountRate) {
  // Discount applies to gains only, never to losses; iff held > 12 months.
  if (rawGain <= 0) return rawGain;
  if (asset.held_over_12_months) return rawGain * (1 - discountRate);
  return rawGain;
}

export function ordinaryRawGain(asset) {
  return asset.projected_sale_proceeds - asset.original_cost_base;
}

export function ordinaryTaxableGain(asset, discountRate) {
  return applyDiscount(ordinaryRawGain(asset), asset, discountRate);
}

export function div296CostBase(asset, resetOn) {
  return resetOn ? asset.market_value_30jun2026 : asset.original_cost_base;
}

export function div296RawGain(asset, resetOn) {
  return asset.projected_sale_proceeds - div296CostBase(asset, resetOn);
}

export function div296AdjustedGain(asset, resetOn, discountRate) {
  return applyDiscount(div296RawGain(asset, resetOn), asset, discountRate);
}

// ---- per-asset ordinary CGT — STANDALONE DIAGNOSTIC VIEW ONLY ------------

export function ordinaryCgt(asset, discountRate, fundCgtRate) {
  return Math.max(0, ordinaryTaxableGain(asset, discountRate)) * fundCgtRate;
}

export function carryForwardLoss(asset) {
  return Math.max(0, -ordinaryRawGain(asset));
}

// ---- fund-level ordinary CGT (s102-5 method statement) ------------------

export function ordinaryCgtFund(assets, discountRate, fundCgtRate) {
  const discountGains = assets
    .filter((a) => a.held_over_12_months)
    .reduce((s, a) => s + Math.max(0, ordinaryRawGain(a)), 0);
  const nondiscountGains = assets
    .filter((a) => !a.held_over_12_months)
    .reduce((s, a) => s + Math.max(0, ordinaryRawGain(a)), 0);
  const grossLosses = assets.reduce(
    (s, a) => s + Math.max(0, -ordinaryRawGain(a)),
    0
  );

  // Apply losses to non-discount gains first (taxpayer-favourable).
  const ndAfter = Math.max(0, nondiscountGains - grossLosses);
  const lossesRemaining = Math.max(0, grossLosses - nondiscountGains);
  const dAfter = Math.max(0, discountGains - lossesRemaining);

  const netTaxable = ndAfter + dAfter * (1 - discountRate);
  return netTaxable * fundCgtRate;
}

export function carryForwardLossFund(assets) {
  const grossGains = assets.reduce(
    (s, a) => s + Math.max(0, ordinaryRawGain(a)),
    0
  );
  const grossLosses = assets.reduce(
    (s, a) => s + Math.max(0, -ordinaryRawGain(a)),
    0
  );
  return Math.max(0, grossLosses - grossGains);
}

// ---- Div 296 fund earnings (intra-year netting, fund-level floor) --------

export function div296FundEarnings(assets, resetOn, discountRate) {
  return Math.max(
    0,
    assets.reduce((s, a) => s + div296AdjustedGain(a, resetOn, discountRate), 0)
  );
}

// ---- per-member Div 296 tax ---------------------------------------------

export function div296TaxForMember(
  earnings,
  member,
  threshold1,
  threshold2,
  rateTier1,
  rateTier2
) {
  const e = earnings * member.split_pct;
  if (e <= 0 || member.tsb <= 0) return 0;
  const band1 =
    Math.max(0, Math.min(member.tsb, threshold2) - threshold1) / member.tsb;
  const band2 = Math.max(0, member.tsb - threshold2) / member.tsb;
  return e * band1 * rateTier1 + e * band2 * rateTier2;
}

export function div296HeadlineTax(
  assets,
  members,
  resetOn,
  discountRate,
  threshold1,
  threshold2,
  rateTier1,
  rateTier2
) {
  const earnings = div296FundEarnings(assets, resetOn, discountRate);
  return members.reduce(
    (s, m) =>
      s +
      div296TaxForMember(
        earnings,
        m,
        threshold1,
        threshold2,
        rateTier1,
        rateTier2
      ),
    0
  );
}

// ---- per-asset Div 296 tax — pro-rata of headline -----------------------

export function perAssetDiv296Tax(
  asset,
  allAssets,
  headlineTax,
  resetOn,
  discountRate
) {
  const myGain = Math.max(0, div296AdjustedGain(asset, resetOn, discountRate));
  const total = allAssets.reduce(
    (s, a) => s + Math.max(0, div296AdjustedGain(a, resetOn, discountRate)),
    0
  );
  if (total <= 0) return 0;
  return (myGain / total) * headlineTax;
}

// ---- per-member band helpers (for transparency display) -----------------

export function memberBands(member, threshold1, threshold2) {
  if (member.tsb <= 0) return { band1: 0, band2: 0 };
  const band1 =
    Math.max(0, Math.min(member.tsb, threshold2) - threshold1) / member.tsb;
  const band2 = Math.max(0, member.tsb - threshold2) / member.tsb;
  return { band1, band2 };
}

// ---- §12 canonical sample register (mirror of tests/test_calcs.py) ------

export const SAMPLE_REGISTER = [
  {
    code: "P1",
    name: "Commercial property",
    quantity: 1,
    original_cost_base: 800_000,
    current_market_value: 2_400_000,
    market_value_30jun2026: 2_400_000,
    valuation_source: "Independent val, 30/06/26",
    projected_sale_proceeds: 2_600_000,
    held_over_12_months: true,
  },
  {
    code: "S1",
    name: "Listed shares parcel",
    quantity: 5_000,
    original_cost_base: 300_000,
    current_market_value: 520_000,
    market_value_30jun2026: 520_000,
    valuation_source: "ASX close 30/06/26",
    projected_sale_proceeds: 600_000,
    held_over_12_months: true,
  },
  {
    code: "L1",
    name: "Loss-making holding",
    quantity: 2_000,
    original_cost_base: 500_000,
    current_market_value: 100_000,
    market_value_30jun2026: 100_000,
    valuation_source: "Independent val, 30/06/26",
    projected_sale_proceeds: 200_000,
    held_over_12_months: true,
  },
];

export const SAMPLE_MEMBERS = [{ tsb: 12_000_000, split_pct: 1.0 }];

/**
 * Compute the full comparison both ways. Returns the headline figures plus
 * per-member and per-asset breakdowns for both scenarios. Difference is signed
 * (elected − no-reset) per the CONTEXT.md glossary convention.
 */
export function computeComparison(assets, members, a = ASSUMPTIONS) {
  const scenario = (resetOn) => {
    const earnings = div296FundEarnings(assets, resetOn, a.discount_rate);
    const headline = div296HeadlineTax(
      assets,
      members,
      resetOn,
      a.discount_rate,
      a.threshold_1,
      a.threshold_2,
      a.rate_tier1,
      a.rate_tier2
    );
    const ordCgt = ordinaryCgtFund(assets, a.discount_rate, a.fund_cgt_rate);
    const perMember = members.map((m) => ({
      member: m,
      tax: div296TaxForMember(
        earnings,
        m,
        a.threshold_1,
        a.threshold_2,
        a.rate_tier1,
        a.rate_tier2
      ),
      ...memberBands(m, a.threshold_1, a.threshold_2),
    }));
    const perAsset = assets.map((as) => ({
      asset: as,
      adjustedGain: div296AdjustedGain(as, resetOn, a.discount_rate),
      tax: perAssetDiv296Tax(as, assets, headline, resetOn, a.discount_rate),
    }));
    return { earnings, headline, ordCgt, perMember, perAsset };
  };

  const noReset = scenario(false);
  const elected = scenario(true);

  // Per-asset signed difference (elected − no-reset) for the "top affected" panel.
  const perAssetDiff = assets.map((as, i) => ({
    asset: as,
    noResetTax: noReset.perAsset[i].tax,
    electedTax: elected.perAsset[i].tax,
    difference: elected.perAsset[i].tax - noReset.perAsset[i].tax,
    // "reset trap": an asset in unrealised loss whose reset CREATES a Div 296 gain
    trap:
      as.current_market_value < as.original_cost_base &&
      elected.perAsset[i].adjustedGain > 0,
  }));

  return {
    assumptions: a,
    noReset,
    elected,
    perAssetDiff,
    headlineDifference: elected.headline - noReset.headline,
    carryForward: carryForwardLossFund(assets),
  };
}
