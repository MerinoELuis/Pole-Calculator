"use strict";

const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");

const sandbox = { window: {}, console };
for (const file of ["domain/height-math.js", "core/project-config.js", "domain/pole-resolution.js", "core/state-compat.js", "core/store.js"]) {
  const statePath = path.join(__dirname, "..", "js", file);
  vm.runInNewContext(fs.readFileSync(statePath, "utf8"), sandbox, { filename: statePath });
}

const store = sandbox.window.AppStore;
const created = store.upsertSpanSide({
  spanId: "S1",
  poleId: "P1",
  proposedHOA: "22'",
  autoCalcProposedStatus: "AUTO",
  autoCalcProposedMode: "TOP_COMM"
});
assert.equal(created.autoCalcProposedStatus, "AUTO");
assert.equal(created.autoCalcProposedMode, "TOP_COMM");

store.setState({
  spanSides: {
    S1__P1: {
      spanId: "S1",
      poleId: "P1",
      proposedHOA: "22'",
      autoCalcProposedStatus: "AUTO",
      autoCalcProposedMode: "TOP_COMM"
    }
  }
});
assert.equal(store.getState().spanSides.S1__P1.autoCalcProposedStatus, "AUTO");
assert.equal(store.getState().spanSides.S1__P1.autoCalcProposedMode, "TOP_COMM");

store.upsertSpanSide({
  ...store.getState().spanSides.S1__P1,
  proposedHOA: "21'8\"",
  autoCalcProposedStatus: "",
  autoCalcProposedMode: ""
});
assert.equal(store.getState().spanSides.S1__P1.autoCalcProposedStatus, "");
assert.equal(store.getState().spanSides.S1__P1.autoCalcProposedMode, "");

console.log("Auto Calculate Proposed source state tests passed");
