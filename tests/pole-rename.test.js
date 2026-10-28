"use strict";

const assert = require("node:assert/strict");

global.window = global;
global.localStorage = { getItem: () => null, setItem: () => {} };
global.Calculations = { recalculateAll: () => {} };

require("../js/domain/height-math.js");
require("../js/core/project-config.js");
require("../js/domain/pole-resolution.js");
require("../js/core/state-compat.js");
require("../js/core/store-identity.js");
require("../js/core/store.js");

const S = AppStore;
S.resetState();
S.upsertPole(S.createPole({ poleId: "P1", lowPower: "30'" }));
S.upsertPole(S.createPole({ poleId: "P2", lowPower: "30'" }));
S.upsertSpan(S.createSpan("S1", "P1", "P2", "E", "", { type: "Fore Span" }));
S.upsertSpanSide({ spanId: "S1", poleId: "P1", proposedHOA: "20'" });
S.upsertSpanComm({ spanId: "S1", poleId: "P1", owner: "CTL", ownerCanonical: "CTL", wireId: "W1", remotePoleId: "P2", existingHOA: "18'" });
S.upsertSpanComm({ spanId: "S1", poleId: "P2", owner: "CTL", ownerCanonical: "CTL", wireId: "W1", remotePoleId: "P1", existingHOA: "18'" });
S.addSpanPower({ spanId: "S1", poleId: "P1", wireId: "POWER-1", label: "Primary", attachmentHeight: "29'" });

const state = S.getState();
state.makeReadyReferences.push(S.createMakeReadyReference({ poleId: "P1", makeReadyNotes: "Pole equipment reference" }));
state.poleClassChecks.push({ poleId: "P1", status: "OK" });
state.movements.push({ poleId: "P1", remotePoleId: "P2", fromPole: "P1", toPole: "P2" });
state.mr.push({ poleId: "P1", text: "Raise CTL." });
state.warnings.push({ poleId: "P1", spanId: "S1", code: "TEST" });
state.selectedPoleId = "P1";
state.ui.hiddenPoleIds = ["P1"];
state.ui.revealedUnknownPoleIds = ["P1"];

const renamed = S.renamePole("P1", "P1-RENAMED");
assert.deepEqual(JSON.parse(JSON.stringify(renamed)), {
  ok: true,
  changed: true,
  previousPoleId: "P1",
  poleId: "P1-RENAMED"
});
assert.equal(S.getPole("P1"), null, "the old pole ID should no longer exist in the graph");
assert.equal(S.getPole("P1-RENAMED")?.poleId, "P1-RENAMED", "the pole map key and pole record should use the new ID");
assert.equal(S.getSpan("S1").fromPole, "P1-RENAMED", "span endpoints should be remapped");
assert.equal(S.getSpanSide("S1", "P1-RENAMED")?.proposedHOA, "20'", "SpanSide keys and values should be remapped");
assert.equal(S.getSpanComm("S1", "P1-RENAMED", "CTL", "W1")?.remotePoleId, "P2", "the renamed endpoint's comm row should remain addressable");
assert.equal(S.getSpanComm("S1", "P2", "CTL", "W1")?.remotePoleId, "P1-RENAMED", "remote pole references should be remapped");
assert.equal(S.getSpanPowerForPole("P1-RENAMED")[0]?.wireId, "POWER-1", "Power rows should be re-keyed and remain addressable");
assert.equal(state.makeReadyReferences[0].poleId, "P1-RENAMED");
assert.equal(state.poleClassChecks[0].poleId, "P1-RENAMED");
assert.equal(state.movements[0].remotePoleId, "P2");
assert.equal(state.movements[0].fromPole, "P1-RENAMED");
assert.equal(state.movements[0].toPole, "P2");
assert.equal(state.mr[0].poleId, "P1-RENAMED");
assert.equal(state.warnings[0].poleId, "P1-RENAMED");
assert.equal(state.selectedPoleId, "P1-RENAMED");
assert.equal(state.poleIdAliases.P1, "P1-RENAMED", "the imported ID should remain an alias for future Update Data reconciliation");
assert.deepEqual(state.ui.hiddenPoleIds, ["P1-RENAMED"]);
assert.deepEqual(state.ui.revealedUnknownPoleIds, ["P1-RENAMED"]);

assert.equal(S.validatePoleRename("P2", "P1-RENAMED").ok, false, "a duplicate real pole ID should be rejected");
assert.equal(S.validatePoleRename("P2", "P1-RENAMED STEEL").ok, false, "a canonical duplicate pole ID should be rejected");
assert.equal(S.validatePoleRename("P2", "  ").ok, false, "an empty pole ID should be rejected");
assert.equal(S.getPole("P2")?.poleId, "P2", "invalid rename attempts must not change the graph");

S.renamePole("P1-RENAMED", "P1-FINAL");
assert.equal(state.poleIdAliases.P1, "P1-FINAL", "repeated renames should point the original imported ID at the latest ID");
assert.equal(state.poleIdAliases["P1-RENAMED"], "P1-FINAL", "each previous ID should resolve to the latest ID");
const normalized = S.normalizeState(JSON.parse(JSON.stringify(state)));
assert.equal(normalized.poles["P1-FINAL"]?.poleId, "P1-FINAL", "the renamed ID should survive save/load normalization");
assert.equal(normalized.poleIdAliases.P1, "P1-FINAL", "the source ID mapping should survive save/load normalization");
console.log("Pole rename tests passed.");
