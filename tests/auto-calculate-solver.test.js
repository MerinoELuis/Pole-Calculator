"use strict";

const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");

function parseHeight(value) {
  if (typeof value === "number" && Number.isFinite(value)) return Math.round(value);
  const raw = String(value ?? "").trim();
  if (!raw) return null;
  if (/^-?\d+(?:\.\d+)?$/.test(raw)) return Math.round(Number(raw));
  const sign = raw.startsWith("-") ? -1 : 1;
  const clean = raw.replace(/^-/, "");
  const feet = Number(clean.match(/(\d+)\s*'/)?.[1] || 0);
  const inches = Number(clean.match(/(\d+)\s*"/)?.[1] || 0);
  return sign * (feet * 12 + inches);
}

function formatHeight(value) {
  const rounded = Math.round(Number(value));
  const sign = rounded < 0 ? "-" : "";
  const absolute = Math.abs(rounded);
  const feet = Math.floor(absolute / 12);
  const inches = absolute % 12;
  return `${sign}${feet}'${inches ? `${inches}\"` : ""}`;
}

const state = {
  settings: {
    position: "TOP_COMM",
    commClearance: "12\"",
    boltClearance: "4\"",
    projectProfile: "INTEC"
  },
  poles: {},
  spans: {},
  spanSides: {},
  spanComms: {}
};

const window = {
  HeightUtils: { parseHeight, formatHeight },
  AppStore: {
    getState: () => state,
    getSpanCommsForPole: () => [],
    getConnectedSpans: () => [],
    getSpanSide: () => null,
    getSpanSidesForPole: () => [],
    keyForSpanComm: (...parts) => parts.join("__")
  },
  Calculations: {
    autoCalculateMovements() {},
    recalculateAll() {},
    isPofComm: () => false,
    commOwnerLabel: row => row.owner || ""
  }
};

const sourcePath = path.join(__dirname, "..", "js", "auto-calculate-solver.js");
vm.runInNewContext(fs.readFileSync(sourcePath, "utf8"), { window, console, Date, JSON, Math, Number, Object, Array, Set, Map, String, RegExp }, { filename: sourcePath });
const solver = window.AutoCalculateSolver;

const poleSafeMidspanBad = {
  poleViolationCount: 0,
  poleViolationInches: 0,
  midspanViolationCount: 1,
  midspanViolationInches: 8,
  movedCommCount: 3,
  totalMovementInches: 24,
  proposedDistanceFromIdeal: 0
};
const poleBadMidspanSafe = {
  poleViolationCount: 1,
  poleViolationInches: 1,
  midspanViolationCount: 0,
  midspanViolationInches: 0,
  movedCommCount: 1,
  totalMovementInches: 4,
  proposedDistanceFromIdeal: 0
};
assert.ok(
  solver.compareAnalyses(poleSafeMidspanBad, poleBadMidspanSafe) < 0,
  "A pole-compliant result must always beat a result that only fixes Midspan."
);
assert.equal(solver.statusForAnalysis(poleSafeMidspanBad), "BEST_AVAILABLE");
assert.equal(solver.statusForAnalysis(poleBadMidspanSafe), "CRITICAL");

const smallerMidspanProblem = {
  ...poleSafeMidspanBad,
  midspanViolationInches: 4,
  movedCommCount: 4,
  totalMovementInches: 30
};
assert.ok(
  solver.compareAnalyses(smallerMidspanProblem, poleSafeMidspanBad) < 0,
  "After the pole is compliant, a smaller Midspan violation must beat fewer movements."
);

const safe = {
  ...poleSafeMidspanBad,
  midspanViolationCount: 0,
  midspanViolationInches: 0,
  movedCommCount: 5,
  totalMovementInches: 36
};
assert.equal(solver.statusForAnalysis(safe), "SAFE");
assert.ok(solver.compareAnalyses(safe, smallerMidspanProblem) < 0);

const groups = [
  { key: "catv", ownerToken: "catv", existingInches: 252, effectiveInches: 252, locked: false },
  { key: "ctl", ownerToken: "ctl", existingInches: 240, effectiveInches: 240, locked: false }
];

const topCandidates = solver.candidateHeights({ groups, maxPole: 288, mode: "TOP_COMM", currentProposed: [], state });
assert.ok(topCandidates.includes(264), "TOP COMM must include Top Comm + Comm-comm as the ideal candidate.");
const lowCandidates = solver.candidateHeights({ groups, maxPole: 288, mode: "LOW_COMM", currentProposed: [], state });
assert.ok(lowCandidates.includes(228), "LOW COMM must include Low Comm - Comm-comm as the ideal candidate.");

const topPlan = solver.buildStackPlan(groups, 258, "TOP_COMM", 288, state);
assert.deepEqual(
  JSON.parse(JSON.stringify(topPlan.map(item => item.targetInches))),
  [246, 234],
  "TOP COMM must build a downward stack below Proposed."
);

const lowPlan = solver.buildStackPlan(groups, 234, "LOW_COMM", 288, state);
assert.deepEqual(
  JSON.parse(JSON.stringify(lowPlan.map(item => item.targetInches))),
  [246, 258],
  "LOW COMM must build an upward stack above Proposed."
);

const lockedGroups = [
  { key: "locked", ownerToken: "catv", existingInches: 252, effectiveInches: 248, locked: true, lockedInches: 248 },
  { key: "other", ownerToken: "ctl", existingInches: 240, effectiveInches: 240, locked: false }
];
const lockedPlan = solver.buildStackPlan(lockedGroups, 258, "TOP_COMM", 288, state);
assert.equal(lockedPlan[0].targetInches, 248, "A user-entered HOA Change must remain fixed.");

const environmentCorrectionPlan = solver.buildStackPlan([
  { key: "catv-env", ownerToken: "catv", existingInches: 208, effectiveInches: 208, minimumInches: 226, maximumInches: null, locked: false },
  { key: "ctl-env", ownerToken: "ctl", existingInches: 184, effectiveInches: 184, minimumInches: 218, maximumInches: null, locked: false }
], 242, "TOP_COMM", 328, state);
assert.deepEqual(
  JSON.parse(JSON.stringify(environmentCorrectionPlan.map(item => item.targetInches))),
  [230, 218],
  "An undersized Proposed candidate must keep the stack inside its available upper envelope."
);

const environmentCorrectionFitPlan = solver.buildStackPlan([
  { key: "catv-env", ownerToken: "catv", existingInches: 208, effectiveInches: 208, minimumInches: 226, maximumInches: null, locked: false },
  { key: "ctl-env", ownerToken: "ctl", existingInches: 184, effectiveInches: 184, minimumInches: 252, maximumInches: null, locked: false }
], 276, "TOP_COMM", 328, state);
assert.deepEqual(
  JSON.parse(JSON.stringify(environmentCorrectionFitPlan.map(item => item.targetInches))),
  [264, 252],
  "TOP COMM must propagate a lower comm's Environment MS correction upward through the pole stack."
);
const environmentCorrectionCandidates = solver.candidateHeights({
  groups: [
    { key: "catv-env", ownerToken: "catv", existingInches: 208, effectiveInches: 208, minimumInches: 226, maximumInches: null, locked: false },
    { key: "ctl-env", ownerToken: "ctl", existingInches: 184, effectiveInches: 184, minimumInches: 252, maximumInches: null, locked: false }
  ],
  maxPole: 328,
  mode: "TOP_COMM",
  currentProposed: [220],
  state
});
assert.ok(
  environmentCorrectionCandidates.includes(276),
  "The required Proposed height must survive candidate truncation when a lower comm forces the stack upward."
);

assert.match(solver.resultMessage(poleSafeMidspanBad, "BEST_AVAILABLE"), /Pole clearances satisfied/i);
assert.match(solver.resultMessage(poleBadMidspanSafe, "CRITICAL"), /No pole-compliant aerial arrangement/i);

state.poles.P1 = { poleId: "P1" };
state.poles.P2 = { poleId: "P2" };
state.spans.S1 = { spanId: "S1", fromPole: "P1", toPole: "P2", type: "Fore Span" };
state.spans.S2 = { spanId: "S2", fromPole: "P2", toPole: "P1", type: "Back Span" };
state.spanComms.remote = {
  spanId: "S2",
  poleId: "P2",
  owner: "CATV",
  flaggingStatus: "PROBLEM",
  flaggingMessage: `Pole: HOA Change 23' exceeds max 22'. Environment: 14'6\" < 15'.`,
  clearanceMSStatus: "OK",
  clearanceMSMessage: "OK"
};
window.AppStore.getConnectedSpans = poleId => Object.values(state.spans).filter(span => span.fromPole === poleId || span.toPole === poleId);
window.AppStore.getSpanCommsForPole = poleId => Object.values(state.spanComms).filter(row => row.poleId === poleId);
const networkIssues = solver.collectIssues("P1");
assert.equal(networkIssues.poleViolationCount, 0, "A remote endpoint pole violation must be scored when its own pole is solved.");
assert.equal(networkIssues.midspanViolationCount, 1, "A reciprocal physical-span Midspan violation must still affect the current pole arrangement.");

console.log("best-available Auto Calculate solver tests passed");
