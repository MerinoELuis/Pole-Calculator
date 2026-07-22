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
S.upsertPole(S.createPole({ poleId: "P1" }));
S.upsertPole(S.createPole({ poleId: "P2" }));
S.upsertSpan(S.createSpan("S1", "P1", "P2", "E", "", { type: "Fore Span" }));
S.upsertComm("P1", "CATV", "20'", "", { wireId: "W1" });
S.upsertComm("P1", "CTL", "19'", "", { wireId: "W2" });
S.upsertComm("P2", "CATV", "21'", "", { wireId: "W1" });
S.upsertSpanComm({ spanId: "S1", poleId: "P1", owner: "CATV", wireId: "W1", existingHOA: "20'", midspan: "17'" });
S.upsertSpanComm({ spanId: "S1", poleId: "P1", owner: "CTL", wireId: "W2", existingHOA: "19'", midspan: "16'" });
S.upsertSpanComm({ spanId: "S1", poleId: "P2", owner: "CATV", wireId: "W1", existingHOA: "21'" });

const removed = S.removeSpanCommsForPole("P1");
assert.equal(removed.length, 2, "Delete Comms must remove every span relationship owned by the selected pole");
assert.equal(S.getSpanCommsForPole("P1").length, 0, "the selected pole must remain without comm rows");
assert.equal(S.getPole("P1").comms.length, 0, "the selected pole comm catalog must also be cleared");
assert.equal(S.getSpanCommsForPole("P2").length, 1, "the opposite endpoint must not be deleted");

const saved = JSON.parse(JSON.stringify(S.getState()));
saved.ui.hiddenCommPoleIds = ["P1", "P1", ""];
S.setState(saved);
assert.equal(S.getSpanCommsForPole("P1").length, 0, "loading a saved job must not recreate intentionally deleted endpoint comms");
assert.deepEqual(Array.from(S.getState().ui.hiddenCommPoleIds), ["P1"], "hidden comm sections must persist as a normalized unique list");

console.log("comm-section-controls.test.js passed");
