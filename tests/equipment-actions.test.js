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
    powerEquipment: [{
      category: "STREETLIGHT",
      attachmentHeight: "26'",
      bottomHeight: "21'",
      dripLoopHeight: "21'",
      actionActive: false,
      actionHeight: ""
    }]
  }
}));
C.recalculateAll();
assert.equal(S.getPole("STREETLIGHT").metadata.powerEquipment[0].actionActive, true, "MidAm must automatically ground every Streetlight");
assert.equal(S.getPole("STREETLIGHT").metadata.lowPowerBaseline, "25'", "old Save files must gain a Low Power baseline without losing mandatory Ground");
assert.equal(S.getPole("STREETLIGHT").lowPower, "25'", "grounding has no vertical target and must not invent a Low Power height");
assert.equal(S.getPole("STREETLIGHT").maxCommHeight, "19'4\"", "mandatory grounding must retain bracket and drip-loop clearances");
assert.match(mrText("STREETLIGHT"), /MNT GROUND STREETLIGHT/);

S.resetState();
S.applyProjectProfile("INTEC");
S.upsertPole(S.createPole({
  poleId: "INTEC-EQUIPMENT",
  lowPower: "23'",
  metadata: {
    lowPowerBaseline: "23'",
    powerEquipment: [
      { category: "TRANSFORMER", dripLoopHeight: "23'", actionActive: true, actionHeight: "25'8\"" },
      { category: "STREETLIGHT", attachmentHeight: "24'", actionActive: true, actionHeight: "" },
      { category: "RISER", attachmentHeight: "27'", actionActive: true, actionHeight: "28'2\"" }
    ]
  }
}));
C.recalculateAll();
const intecEquipmentMR = mrText("INTEC-EQUIPMENT");
assert.match(intecEquipmentMR, /Secure transformer drip loop to HOA 25'8"\./);
assert.match(intecEquipmentMR, /Install flex conduit to STLT circuit\. bond STLT housing to pole GRND\/NEUT\./);
assert.match(intecEquipmentMR, /Raise APS riser from HOA 27' to HOA 28'2"\./);
assert.equal(
  C.getPowerEquipmentCeiling({ category: "STREETLIGHT", attachmentHeight: "24'2\"", actionActive: true }),
  "23'2\"",
  "an INTEC grounded Streetlight bracket must keep 12 inches to comm"
);
assert.equal(
  C.getPowerEquipmentCeiling({ category: "STREETLIGHT", attachmentHeight: "24'2\"", actionActive: false }),
  "20'10\"",
  "an INTEC ungrounded Streetlight bracket must keep 40 inches to comm"
);
assert.equal(
  C.getPowerEquipmentCeiling({ category: "STREETLIGHT", attachmentHeight: "24'2\"", dripLoopHeight: "22'", actionActive: true }),
  "18'8\"",
  "an INTEC Streetlight drip loop must still keep 40 inches after the bracket is grounded"
);
assert.equal(
  C.getPowerEquipmentCeiling({ category: "TRANSFORMER", attachmentHeight: "28'", bottomHeight: "25'" }),
  "21'8\"",
  "an INTEC Transformer must keep 40 inches below its Bottom Height"
);

console.log("Power Equipment action tests passed.");
