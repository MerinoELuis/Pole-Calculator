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

for (const file of ["height-utils.js", "project-config.js", "state.js", "calculations.js"]) {
  const modulePath = path.join(__dirname, "..", "js", file);
  vm.runInNewContext(fs.readFileSync(modulePath, "utf8"), sandbox, { filename: modulePath });
}

const S = window.AppStore;
const C = window.Calculations;
S.resetState();
S.applyProjectProfile("METRONET");
S.upsertPole(S.createPole({ poleId: "P1", lowPower: "30'" }));
S.upsertPole(S.createPole({
  poleId: "P2",
  lowPower: "30'",
  metadata: {
    powerEquipment: [
      { category: "STREETLIGHT", bottomHeight: "25'", dripLoopHeight: "24'6\"" },
      { category: "TRANSFORMER", attachmentHeight: "28'" }
    ],
    midAmConstraints: {
      streetlights: [{ bottomHeight: "25'", dripLoopHeight: "24'6\"" }],
      powerGuys: []
    }
  }
}));
S.upsertSpan(S.createSpan("FORE", "P1", "P2", "S", "", { type: "Fore Span", rawType: "Fore Span" }));
S.upsertSpan(S.createSpan("BACK", "P2", "P1", "N", "", { type: "Back Span", rawType: "Back Span" }));

S.upsertComm("P1", "COMMUNICATION > Fiber", "19'", "", { wireId: "FORE-WIRE" });
S.upsertComm("P2", "COMMUNICATION > Fiber", "20'2\"", "", { wireId: "BACK-WIRE" });
S.upsertSpanComm(S.createSpanComm({
  spanId: "FORE",
  poleId: "P1",
  owner: "COMMUNICATION > Fiber",
  ownerBase: "COMMUNICATION > Fiber",
  wireId: "FORE-WIRE",
  existingHOA: "19'"
}));
S.upsertSpanComm(S.createSpanComm({
  spanId: "BACK",
  poleId: "P2",
  owner: "COMMUNICATION > Fiber",
  ownerBase: "COMMUNICATION > Fiber",
  wireId: "BACK-WIRE",
  existingHOA: "20'2\"",
  midspan: "18'10\"",
  ocalcMS: "18'10\""
}));

C.recalculateAll();
let back = S.getSpanComm("BACK", "P2", "COMMUNICATION > Fiber", "BACK-WIRE");
assert.equal(S.getPole("P2").maxCommHeight, "23'4\"", "the lower MidAm streetlight ceiling must control Max Height on Pole");
assert.equal(C.isCalculatedBackspanComm(back), true, "MidAm Back Span with an imported midspan must be calculated");
assert.equal(back.calculatedMidspan, "18'10\"", "unchanged endpoints must preserve the imported baseline");

// Different directed Span/Wire IDs still describe the same physical endpoint
// pair. Moving only the Fore endpoint down 12 inches changes the Back MS by 6.
S.upsertSpanComm({
  ...S.getSpanComm("FORE", "P1", "COMMUNICATION > Fiber", "FORE-WIRE"),
  existingHOAChange: "18'"
});
C.recalculateSpansForPole("P1");
back = S.getSpanComm("BACK", "P2", "COMMUNICATION > Fiber", "BACK-WIRE");
assert.equal(back.remoteHOA, "18'", "Back Span must expose the effective HOA at the reciprocal pole");
assert.equal(back.calculatedMidspan, "18'4\"", "a 12-inch remote movement must change the MidAm Back Span by 6 inches");

S.upsertSpanComm({ ...back, existingHOAChange: "19'2\"" });
C.recalculateSpansForPole("P2");
back = S.getSpanComm("BACK", "P2", "COMMUNICATION > Fiber", "BACK-WIRE");
assert.equal(back.calculatedMidspan, "17'10\"", "movements at both endpoints must each contribute half their change exactly once");

S.resetState();
S.applyProjectProfile("INTEC");
S.upsertPole(S.createPole({
  poleId: "INTEC-EQUIPMENT",
  lowPower: "30'",
  metadata: {
    powerEquipment: [{ category: "TRANSFORMER", attachmentHeight: "25'" }]
  }
}));
C.recalculateAll();
assert.equal(S.getPole("INTEC-EQUIPMENT").maxCommHeight, "21'8\"", "INTEC Power Equipment must constrain Max Height on Pole using Power-comms clearance");

console.log("MidAm Back Span calculation tests passed.");
