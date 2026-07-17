"use strict";

const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");

const window = {
  localStorage: { getItem: () => null, setItem: () => {} },
  Validations: { validateAll: () => {}, validatePole: () => {}, validateSpan: () => {} }
};
const sandbox = { window, console, Date, Set, Map, Math, JSON, Object, Array, String, Number, Boolean, RegExp };

for (const file of ["height-utils.js", "project-config.js", "state.js", "mr-logic.js", "calculations.js"]) {
  const modulePath = path.join(__dirname, "..", "js", file);
  vm.runInNewContext(fs.readFileSync(modulePath, "utf8"), sandbox, { filename: modulePath });
}

const S = window.AppStore;
const C = window.Calculations;

function mrText(poleId) {
  return S.getState().mr.find(item => item.poleId === poleId)?.text || "";
}

S.resetState();
S.applyProjectProfile("METRONET");
S.upsertPole(S.createPole({
  poleId: "TRANSFORMER",
  lowPower: "24'",
  metadata: {
    lowPowerBaseline: "24'",
    powerEquipment: [{
      category: "TRANSFORMER",
      attachmentHeight: "30'",
      dripLoopHeight: "24'",
      actionActive: true,
      actionHeight: "25'6\""
    }]
  }
}));
C.recalculateAll();
assert.equal(S.getPole("TRANSFORMER").lowPower, "25'6\"", "redressing the limiting drip loop must update Low Power");
assert.match(mrText("TRANSFORMER"), /POWER REDRESS TRANSFORMER DRIP LOOP TO HOA 25'6"\./);

S.updatePowerEquipmentField("TRANSFORMER", 0, "actionActive", false);
C.recalculateAll();
assert.equal(S.getPole("TRANSFORMER").lowPower, "24'", "disabling redress must restore imported Low Power");
assert.equal(mrText("TRANSFORMER"), "", "disabled equipment work must leave no MR instruction");

S.upsertPole(S.createPole({
  poleId: "RISER",
  lowPower: "23'6\"",
  metadata: {
    lowPowerBaseline: "23'6\"",
    powerEquipment: [{
      category: "RISER",
      attachmentHeight: "23'6\"",
      actionActive: true,
      actionHeight: "26'6\""
    }]
  }
}));
C.recalculateAll();
assert.equal(S.getPole("RISER").lowPower, "26'6\"", "raising a limiting Power Riser must update Low Power");
assert.match(mrText("RISER"), /AT HOA 23'6" RAISE POWER RISER TO HOA 26'6" DUE TO CLEARANCES\./);

S.updatePowerEquipmentField("RISER", 0, "actionHeight", "23'");
C.recalculateAll();
assert.equal(S.getPole("RISER").lowPower, "23'6\"", "a Riser target that does not raise the attachment must be ignored");
assert.equal(mrText("RISER"), "", "an invalid Riser raise must not generate MR");

S.upsertPole(S.createPole({
  poleId: "STREETLIGHT",
  lowPower: "25'",
  metadata: {
    lowPowerBaseline: "25'",
    powerEquipment: [{
      category: "STREETLIGHT",
      attachmentHeight: "26'",
      actionActive: true,
      actionHeight: ""
    }]
  }
}));
C.recalculateAll();
assert.equal(S.getPole("STREETLIGHT").lowPower, "25'", "grounding has no vertical target and must not invent a Low Power height");
assert.match(mrText("STREETLIGHT"), /MNT GROUND STREETLIGHT/);

console.log("Power Equipment action tests passed.");
