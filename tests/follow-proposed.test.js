"use strict";

const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");

const window = {
  localStorage: { getItem: () => null, setItem: () => {} },
  MRLogic: { generateAllMR: () => {}, generateMRForPole: () => {} },
  Validations: { validateAll: () => {}, validatePole: () => {}, validateSpan: () => {} }
};
const sandbox = { window, console, Date, Set, Map, Math, JSON, Object, Array, String, Number, Boolean, RegExp };

for (const file of ["domain/height-math.js", "core/project-config.js", "domain/pole-resolution.js", "core/state-compat.js", "core/store.js", "domain/comm-movements.js", "domain/midspan-engine.js", "domain/clearance-rules.js", "rules/auto-calculate-solver.js"]) {
  const modulePath = path.join(__dirname, "..", "js", file);
  vm.runInNewContext(fs.readFileSync(modulePath, "utf8"), sandbox, { filename: modulePath });
}

const S = window.AppStore;
const solver = window.AutoCalculateSolver;
const H = window.HeightUtils;

function setup(profile = "METRONET") {
  S.resetState();
  S.applyProjectProfile(profile);
  S.upsertPole(S.createPole({ poleId: "P1", maxCommHeight: "25'" }));
  S.upsertPole(S.createPole({ poleId: "P2" }));
  S.upsertSpan(S.createSpan("S1", "P1", "P2", "E", "", { type: "Fore Span", lengthDisplay: "100'" }));
}

setup();
S.upsertSpanComm(S.createSpanComm({ spanId: "S1", poleId: "P1", owner: "Telco", existingHOA: "20'4\"" }));
S.upsertSpanSide(S.createSpanSide({ spanId: "S1", poleId: "P1", proposedHOA: "20'10\"" }));
let result = solver.applyProposedCommTarget("P1", H.parseHeight("20'10\""));
assert.equal(result.applied, true);
assert.equal(S.getSpanComm("S1", "P1", "Telco", "").existingHOAChange, "19'10\"");

// MidAm uses the 6-inch bolt spacing when the desired target is too close to
// another existing bolt. The source bolt at 20' therefore pushes the moved
// comm down to 19'6" instead of leaving it two inches away.
setup();
S.upsertSpanComm(S.createSpanComm({ spanId: "S1", poleId: "P1", owner: "Telco", existingHOA: "20'" }));
result = solver.applyProposedCommTarget("P1", H.parseHeight("20'10\""));
assert.equal(S.getSpanComm("S1", "P1", "Telco", "").existingHOAChange, "19'6\"");

// A second Proposed edit recalculates from imported HOA rather than treating a
// previous manual comm edit as a permanent lock.
const row = S.getSpanComm("S1", "P1", "Telco", "");
S.upsertSpanComm({ ...row, existingHOAChange: "18'" });
solver.applyProposedCommTarget("P1", H.parseHeight("21'10\""));
assert.equal(S.getSpanComm("S1", "P1", "Telco", "").existingHOAChange, "20'10\"");

// Follow Proposed rebuilds the stack from the configured minimums instead of
// preserving a larger imported gap. The first target stays one foot below the
// Proposed and the second seeks the configured Comm-comm clearance even when
// both rows belong to the same owner.
setup();
S.upsertSpanComm(S.createSpanComm({ spanId: "S1", poleId: "P1", owner: "Telco", wireId: "A", existingHOA: "20'4\"" }));
S.upsertSpanComm(S.createSpanComm({ spanId: "S1", poleId: "P1", owner: "Telco", wireId: "B", existingHOA: `18'` }));
solver.applyProposedCommTarget("P1", H.parseHeight("20'10\""));
assert.equal(S.getSpanComm("S1", "P1", "Telco", "A").existingHOAChange, "19'10\"");
assert.equal(S.getSpanComm("S1", "P1", "Telco", "B").existingHOAChange, "18'10\"");

// A same-owner pair that was originally about one foot apart must remain one
// foot apart after Follow Proposed (not collapse to the six-inch bolt rule).
setup();
S.upsertSpanComm(S.createSpanComm({ spanId: "S1", poleId: "P1", owner: "CATV", wireId: "C", existingHOA: `22'4"` }));
S.upsertSpanComm(S.createSpanComm({ spanId: "S1", poleId: "P1", owner: "Telco", wireId: "A", existingHOA: `20'8"` }));
S.upsertSpanComm(S.createSpanComm({ spanId: "S1", poleId: "P1", owner: "Telco", wireId: "B", existingHOA: `19'10"` }));
solver.applyProposedCommTarget("P1", H.parseHeight("21'2\""));
assert.equal(S.getSpanComm("S1", "P1", "CATV", "C").existingHOAChange, "19'10\"");
assert.equal(S.getSpanComm("S1", "P1", "Telco", "A").existingHOAChange, "18'10\"");
assert.equal(S.getSpanComm("S1", "P1", "Telco", "B").existingHOAChange, "17'10\"");

S.updateSetting("position", "LOW_COMM");
assert.equal(solver.applyProposedCommTarget("P1", H.parseHeight("20'10\"")).disabled, true);

console.log("follow-proposed tests passed");
