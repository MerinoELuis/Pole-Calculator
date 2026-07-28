"use strict";

const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");

let state = { spanSides: {} };
const keyForSpanSide = (spanId, poleId) => `${spanId}__${poleId}`;
const store = {
  keyForSpanSide,
  getState: () => state,
  upsertSpanSide(data = {}) {
    const normalized = {
      spanId: String(data.spanId || "").trim(),
      poleId: String(data.poleId || "").trim(),
      proposedHOA: String(data.proposedHOA || "").trim()
    };
    state.spanSides[keyForSpanSide(normalized.spanId, normalized.poleId)] = normalized;
    return normalized;
  },
  setState(nextState = {}) {
    state = {
      spanSides: Object.fromEntries(
        Object.values(nextState.spanSides || {}).map(side => {
          const normalized = {
            spanId: String(side.spanId || "").trim(),
            poleId: String(side.poleId || "").trim(),
            proposedHOA: String(side.proposedHOA || "").trim()
          };
          return [keyForSpanSide(normalized.spanId, normalized.poleId), normalized];
        })
      )
    };
    return state;
  }
};

const window = { AppStore: store };
const sourcePath = path.join(__dirname, "..", "js", "auto-calculate-source-compat.js");
vm.runInNewContext(fs.readFileSync(sourcePath, "utf8"), { window, String, Object, Map }, { filename: sourcePath });

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
assert.equal(state.spanSides.S1__P1.autoCalcProposedStatus, "AUTO");
assert.equal(state.spanSides.S1__P1.autoCalcProposedMode, "TOP_COMM");

store.upsertSpanSide({
  ...state.spanSides.S1__P1,
  proposedHOA: "21'8\"",
  autoCalcProposedStatus: "",
  autoCalcProposedMode: ""
});
assert.equal(state.spanSides.S1__P1.autoCalcProposedStatus, "");
assert.equal(state.spanSides.S1__P1.autoCalcProposedMode, "");

console.log("Auto Calculate source compatibility tests passed");
