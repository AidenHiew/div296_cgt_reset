/*
 * parity.test.js — verifies calcs.js reproduces the §12 acceptance numbers
 * that tests/test_calcs.py (the Python source of truth) asserts.
 *
 * Run with:  node web/tests/parity.test.js
 * No test framework — plain asserts, exits non-zero on first failure.
 */

import assert from "node:assert/strict";
import {
  ASSUMPTIONS as A,
  SAMPLE_REGISTER,
  SAMPLE_MEMBERS,
  ordinaryTaxableGain,
  ordinaryCgtFund,
  div296AdjustedGain,
  div296FundEarnings,
  div296HeadlineTax,
  perAssetDiv296Tax,
  carryForwardLossFund,
  computeComparison,
} from "../calcs.js";

const r = (x) => Math.round(x);
const [PROPERTY, SHARES, LOSS] = SAMPLE_REGISTER;
const REGISTER = SAMPLE_REGISTER;
const MEMBERS = SAMPLE_MEMBERS;

let passed = 0;
function check(label, actual, expected) {
  assert.equal(actual, expected, `${label}: got ${actual}, want ${expected}`);
  passed++;
}

// --- ordinary taxable gains ---
check("P1 ord taxable gain", r(ordinaryTaxableGain(PROPERTY, A.discount_rate)), 1_200_000);
check("S1 ord taxable gain", r(ordinaryTaxableGain(SHARES, A.discount_rate)), 200_000);
check("L1 ord taxable gain", r(ordinaryTaxableGain(LOSS, A.discount_rate)), -300_000);

// --- fund ordinary CGT (s102-5 netting) ---
check("fund ord CGT", r(ordinaryCgtFund(REGISTER, A.discount_rate, A.fund_cgt_rate)), 180_000);

// --- reset ON (elected) ---
check("P1 div296 adj gain (reset)", r(div296AdjustedGain(PROPERTY, true, A.discount_rate)), 133_333);
check("S1 div296 adj gain (reset)", r(div296AdjustedGain(SHARES, true, A.discount_rate)), 53_333);
check("L1 div296 adj gain (reset, the trap)", r(div296AdjustedGain(LOSS, true, A.discount_rate)), 66_667);
check("div296 earnings (reset)", r(div296FundEarnings(REGISTER, true, A.discount_rate)), 253_333);

const headlineOn = div296HeadlineTax(
  REGISTER, MEMBERS, true, A.discount_rate,
  A.threshold_1, A.threshold_2, A.rate_tier1, A.rate_tier2
);
check("headline tax (reset)", r(headlineOn), 32_722);
check("P1 per-asset div296 (reset)", r(perAssetDiv296Tax(PROPERTY, REGISTER, headlineOn, true, A.discount_rate)), 17_222);
check("S1 per-asset div296 (reset)", r(perAssetDiv296Tax(SHARES, REGISTER, headlineOn, true, A.discount_rate)), 6_889);
check("L1 per-asset div296 (reset)", r(perAssetDiv296Tax(LOSS, REGISTER, headlineOn, true, A.discount_rate)), 8_611);
check("fund carry-forward (reset register)", r(carryForwardLossFund(REGISTER)), 0);

// --- reset OFF (no election) ---
check("P1 div296 adj gain (no reset)", r(div296AdjustedGain(PROPERTY, false, A.discount_rate)), 1_200_000);
check("S1 div296 adj gain (no reset)", r(div296AdjustedGain(SHARES, false, A.discount_rate)), 200_000);
check("L1 div296 adj gain (no reset)", r(div296AdjustedGain(LOSS, false, A.discount_rate)), -300_000);
check("div296 earnings (no reset, netted)", r(div296FundEarnings(REGISTER, false, A.discount_rate)), 1_100_000);

const headlineOff = div296HeadlineTax(
  REGISTER, MEMBERS, false, A.discount_rate,
  A.threshold_1, A.threshold_2, A.rate_tier1, A.rate_tier2
);
check("headline tax (no reset)", r(headlineOff), 142_083);

// --- net effect ---
check("net effect of electing reset", r(headlineOff - headlineOn), 109_361);

// --- end-to-end computeComparison ---
const cmp = computeComparison(REGISTER, MEMBERS);
check("compare: no-reset headline", r(cmp.noReset.headline), 142_083);
check("compare: elected headline", r(cmp.elected.headline), 32_722);
check("compare: signed difference", r(cmp.headlineDifference), -109_361);
check("compare: L1 flagged as trap", String(cmp.perAssetDiff[2].trap), "true");

console.log(`\n✓ all ${passed} parity checks passed — calcs.js matches §12 acceptance numbers\n`);
