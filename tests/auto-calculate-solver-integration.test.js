"use strict";

const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");

function parseHeight(value) {
  if (typeof value === "number" && Number.isFinite(value)) return Math.round(value);
  const raw = String(value ?? "").trim();
  if (!raw) return null;
  const sign = raw.startsWith("-") ? -1 : 1;
  const clean = raw.replace(/^-/, "");
  const feet = Number(clean.match(/(\d+)\s*'/)?.[1] || 0);
  const inches = Number(clean.match(/(\d+)\s*"/)?.[1] || 0);
  return sign * (feet * 12 + inches);
}

function formatHeight(value) {
  const rounded = Math.round(Number(value));
  const feet = Math.floor(Math.abs(rounded) / 12);
  const inches = Math.abs(rounded) % 12;
  return `${rounded < 0 ? "-" : ""}${feet}'${inches ? `${inches}\"` : ""}`;
}

let state = {
  settings: { position: "TOP_COMM", commClearance: "12\"", boltClearance: "4\"", projectProfile: "INTEC" },
  poles: { P1: { poleId: "P1", maxCommHeight: "22'", metadata: {} }, P2: { poleId: "P2", metadata: {} } },
  spans: { S1: { spanId: "S1", fromPole: "P1", toPole: "P2", type: "Fore Span" } },
  spanSides: { S1__P1: { spanId: "S1", poleId: "P1", proposedHOA: "" } },
  spanComms: {
    S1__P1__CATV__: {
      spanId: "S1", poleId: "P1", owner: "CATV", existingHOA: "21'6\"", existingHOAChange: "", midspan: "15'",
      flaggingStatus: "OK", flaggingMessage: "OK", clearanceMSStatus: "OK", clearanceMSMessage: "OK"
    }
  }
};
const initialState = JSON.stringify(state);

const sideKey = (spanId, poleId) => `${spanId}__${poleId}`;
const commKey = (spanId, poleId, owner, wireId = "") => `${spanId}__${poleId}__${owner}__${wireId}`;
const AppStore = {
  getState: () => state,
  setState(next) { state = JSON.parse(JSON.stringify(next)); return state; },
  getPole: poleId => state.poles[poleId] || null,
  getSpanSide: (spanId, poleId) => state.spanSides[sideKey(spanId, poleId)] || null,
  getSpanSidesForPole: poleId => Object.values(state.spanSides).filter(side => side.poleId === poleId),
  getSpanCommsForPole: poleId => Object.values(state.spanComms).filter(row => row.poleId === poleId),
  getConnectedSpans: poleId => Object.values(state.spans).filter(span => span.fromPole === poleId || span.toPole === poleId),
  keyForSpanSide: sideKey,
  keyForSpanComm: commKey,
  upsertSpanSide(side) { state.spanSides[sideKey(side.spanId, side.poleId)] = { ...side }; return state.spanSides[sideKey(side.spanId, side.poleId)]; },
  upsertSpanComm(row) { state.spanComms[commKey(row.spanId, row.poleId, row.owner, row.wireId || "")] = { ...row }; return state.spanComms[commKey(row.spanId, row.poleId, row.owner, row.wireId || "")]; }
};

function recalculate() {
  const row = state.spanComms.S1__P1__CATV__;
  const side = state.spanSides.S1__P1;
  const effective = parseHeight(row.existingHOAChange || row.existingHOA);
  const proposed = parseHeight(side.proposedHOA);
  const midspan = 180 + Math.round((effective - 258) / 2);
  row.flaggingStatus = midspan < 180 ? "PROBLEM" : "OK";
  row.flaggingMessage = midspan < 180 ? `Environment: ${formatHeight(midspan)} < 15'.` : "OK";
  row.clearanceMSStatus = "OK";
  row.clearanceMSMessage = "OK";
  const poleProblem = proposed !== null && (proposed > 264 || proposed - effective < 12);
  side.proposedFlaggingStatus = poleProblem ? "PROBLEM" : "OK";
  side.proposedFlaggingMessage = poleProblem ? `Proposed ${formatHeight(proposed)} does not respect Pole · Comm-comm 12\" against CATV ${formatHeight(effective)}.` : "OK";
  side.clearanceMSStatus = "OK";
  side.clearanceMSMessage = "OK";
}

const Calculations = {
  autoCalculateMovements() {}, recalculateAll: recalculate, recalculateSpansForPole: recalculate,
  spanHasRealMidspan: () => true, isPofComm: () => false, commOwnerLabel: row => row.owner,
  updateExistingHOAChange() {}, updateSpanSideField() {}
};
const window = { AppStore, Calculations, HeightUtils: { parseHeight, formatHeight } };
const sourcePath = path.join(__dirname, "..", "js", "auto-calculate-solver.js");
vm.runInNewContext(fs.readFileSync(sourcePath, "utf8"), { window, console, Date, JSON, Math, Number, Object, Array, Set, Map, String, RegExp }, { filename: sourcePath });

(async () => {
  const candidateProgress = [];
  const solved = await window.AutoCalculateSolver.solvePole("P1", "TOP_COMM", {
    onCandidateProgress: detail => candidateProgress.push(detail)
  });
  assert.equal(solved.status, "BEST_AVAILABLE");
  assert.equal(state.spanSides.S1__P1.proposedHOA, "22'");
  assert.equal(state.spanComms.S1__P1__CATV__.existingHOAChange, "21'");
  assert.equal(state.poles.P1.metadata.autoCalculateResult.poleViolationCount, 0);
  assert.equal(state.poles.P1.metadata.autoCalculateResult.midspanViolationCount, 1);
  assert.match(state.poles.P1.metadata.autoCalculateResult.recommendation, /UG or PCO/);
  assert.ok(candidateProgress.length > 0, "candidate progress should be reported while evaluating a pole");
  assert.equal(candidateProgress.at(-1).candidateIndex, candidateProgress.at(-1).candidateCount);

  state = JSON.parse(initialState);
  const runProgress = [];
  const summary = await window.AutoCalculateSolver.autoCalculateMovements({
    onProgress: detail => runProgress.push(detail)
  });
  assert.equal(summary.disabled, false);
  assert.equal(runProgress[0].phase, "starting");
  assert.equal(runProgress.at(-1).phase, "complete");
  assert.equal(runProgress.at(-1).progress, 100);

  console.log("best-available Auto Calculate integration tests passed");
})().catch(error => {
  console.error(error);
  process.exitCode = 1;
});
