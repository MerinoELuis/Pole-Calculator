"use strict";

const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");

const window = { localStorage: { getItem: () => null, setItem: () => {} } };
const sandbox = { window, console, Date, Set, Map, Math, JSON, Object, Array, String, Number, Boolean, RegExp };

for (const file of ["height-utils.js", "project-config.js", "state.js"]) {
  const modulePath = path.join(__dirname, "..", "js", file);
  vm.runInNewContext(fs.readFileSync(modulePath, "utf8"), sandbox, { filename: modulePath });
}

const S = window.AppStore;
S.resetState();
S.upsertPole(S.createPole({ poleId: "P1 STEEL" }));
S.upsertPole(S.createPole({ poleId: "P2" }));
S.upsertPole(S.createPole({ poleId: "P3" }));
S.upsertSpan(S.createSpan("S1", "P1 STEEL", "P2", "E", "", { type: "Fore Span" }));
S.upsertSpan(S.createSpan("S2", "P2", "P3", "E", "", { type: "Fore Span" }));
S.upsertSpanComm({ spanId: "S1", poleId: "P1 STEEL", owner: "CATV", wireId: "W1", existingHOA: "20'" });
S.upsertSpanComm({ spanId: "S1", poleId: "P2", owner: "CATV", wireId: "W1", existingHOA: "21'" });
S.upsertSpanComm({ spanId: "S2", poleId: "P2", owner: "CTL", wireId: "W2", existingHOA: "19'" });
S.upsertSpanSide({ spanId: "S1", poleId: "P1 STEEL", proposedHOA: "22'" });
S.addSpanPower({ spanId: "S1", poleId: "P1 STEEL", wireId: "PW1", midspan: "25'" });

const result = S.removePole("P1 STEEL");
assert.deepEqual(Array.from(result.spanIds), ["S1"], "only spans connected to the deleted pole should be removed");
assert.equal(S.getPole("P1 STEEL"), null, "the selected pole should be deleted");
assert.equal(S.getSpan("S1"), null, "a connected span should be deleted");
assert.ok(S.getSpan("S2"), "an unrelated span should remain");
assert.equal(S.getSpanCommsForSpan("S1").length, 0, "both endpoints of a deleted span should be removed");
assert.equal(S.getSpanCommsForSpan("S2").length, 1, "unrelated comm rows should remain");
assert.deepEqual(Array.from(S.getState().deletedPoleIds), ["P1"], "the canonical pole identity should be retained for Update Data");

const saved = JSON.parse(JSON.stringify(S.getState()));
saved.ui.hiddenPoleIds = ["P2", "P2", "missing"];
S.setState(saved);
assert.deepEqual(Array.from(S.getState().ui.hiddenPoleIds), ["P2"], "hidden poles should persist as valid unique IDs");
assert.deepEqual(Array.from(S.getState().deletedPoleIds), ["P1"], "deleted pole tombstones should survive Save and Load");

console.log("pole-controls.test.js passed");
