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

for (const file of ["domain/height-math.js", "core/project-config.js", "domain/pole-resolution.js", "core/state-compat.js", "core/store.js", "domain/comm-movements.js", "domain/midspan-engine.js", "domain/clearance-rules.js"]) {
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
      { category: "STREETLIGHT", attachmentHeight: "25'", bottomHeight: "", dripLoopHeight: "" },
      { category: "TRANSFORMER", attachmentHeight: "28'" }
    ],
    midAmConstraints: {
      streetlights: [{ attachmentHeight: "25'", bottomHeight: "", dripLoopHeight: "" }],
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
S.upsertPole(S.createPole({ poleId: "INTEC-P1", lowPower: "30'" }));
S.upsertPole(S.createPole({ poleId: "INTEC-P2", lowPower: "30'" }));
S.upsertSpan(S.createSpan("INTEC-BACK", "INTEC-P2", "INTEC-P1", "N", "", { type: "Back Span", rawType: "Back Span" }));
S.upsertSpanComm(S.createSpanComm({
  spanId: "INTEC-BACK",
  poleId: "INTEC-P2",
  owner: "COMMUNICATION > CATV",
  wireId: "INTEC-BACK-WIRE",
  existingHOA: "20'",
  midspan: "18'10\""
}));
// Saves created before INTEC Back Span calculation existed contain an explicit
// false. Loading that state must migrate to the current project rule.
const legacyIntecState = JSON.parse(JSON.stringify(S.getState()));
legacyIntecState.settings.calculateBackspanMidspan = false;
S.setState(legacyIntecState);
assert.equal(S.getState().settings.calculateBackspanMidspan, true, "legacy INTEC saves must enable current Back Span calculation rules");
C.recalculateAll();
let intecBack = S.getSpanComm("INTEC-BACK", "INTEC-P2", "COMMUNICATION > CATV", "INTEC-BACK-WIRE");
assert.equal(C.isCalculatedBackspanComm(intecBack), true, "INTEC Back Span with its own imported midspan must calculate");
assert.equal(intecBack.calculatedMidspan, "18'10\"", "INTEC must preserve the imported Back Span baseline before movements");

S.upsertSpanComm({ ...intecBack, existingHOAChange: "19'2\"" });
C.recalculateSpansForPole("INTEC-P2");
intecBack = S.getSpanComm("INTEC-BACK", "INTEC-P2", "COMMUNICATION > CATV", "INTEC-BACK-WIRE");
assert.equal(intecBack.calculatedMidspan, "18'5\"", "a 10-inch INTEC endpoint movement must apply half, or 5 inches, to its Back Span midspan");

S.upsertPole(S.createPole({
  poleId: "INTEC-EQUIPMENT",
  lowPower: "30'",
  metadata: {
    powerEquipment: [{ category: "TRANSFORMER", attachmentHeight: "25'" }]
  }
}));
C.recalculateAll();
assert.equal(S.getPole("INTEC-EQUIPMENT").maxCommHeight, "21'8\"", "INTEC Power Equipment must constrain Max Height on Pole using Power-comms clearance");

// MidAm can store the measured cable midspan only on the Fore relationship
// while the editable movements are on the reciprocal REF relationship. The
// reference must use that same P14-P15 physical span, never P15-P16.
S.resetState();
S.applyProjectProfile("METRONET");
S.updateSetting("proposedOwner", "MidAm");
S.upsertPole(S.createPole({ poleId: "P14", lowPower: "30'" }));
S.upsertPole(S.createPole({ poleId: "P15", lowPower: "30'" }));
S.upsertPole(S.createPole({ poleId: "P16", lowPower: "30'" }));
S.upsertSpan(S.createSpan("P14-P15-FORE", "P14", "P15", "S", "", { type: "Fore Span", rawType: "Fore Span" }));
S.upsertSpan(S.createSpan("P15-P14-BACK", "P15", "P14", "N", "", { type: "Back Span", rawType: "Back Span" }));
S.upsertSpan(S.createSpan("P15-P16-FORE", "P15", "P16", "S", "", { type: "Fore Span", rawType: "Fore Span" }));

const catv = "COMMUNICATION > CATV";
const telco = "COMMUNICATION > Telco";
S.upsertSpanComm(S.createSpanComm({
  spanId: "P14-P15-FORE", poleId: "P14", owner: catv, wireId: "P14-CATV",
  existingHOA: "19'10\"", midspan: "16'3\""
}));
S.upsertSpanComm(S.createSpanComm({
  spanId: "P14-P15-FORE", poleId: "P14", owner: telco, wireId: "P14-TELCO",
  existingHOA: "18'6\"", midspan: "15'2\""
}));
S.upsertSpanComm(S.createSpanComm({
  spanId: "P15-P14-BACK", poleId: "P14", owner: catv, wireId: "P14-BACK-CATV",
  existingHOA: "19'10\""
}));
S.upsertSpanComm(S.createSpanComm({
  spanId: "P15-P14-BACK", poleId: "P14", owner: telco, wireId: "P14-BACK-TELCO",
  existingHOA: "18'6\""
}));
S.upsertSpanComm(S.createSpanComm({
  spanId: "P15-P14-BACK", poleId: "P15", owner: catv, wireId: "P15-BACK-CATV",
  existingHOA: "18'6\"", existingHOAChange: "20'6\""
}));
S.upsertSpanComm(S.createSpanComm({
  spanId: "P15-P14-BACK", poleId: "P15", owner: telco, wireId: "P15-BACK-TELCO",
  existingHOA: "16'8\"", existingHOAChange: "19'"
}));
S.upsertSpanComm(S.createSpanComm({
  // The same cable identity continues through P15, so this deliberately
  // reuses P14-CATV. Remote matching must still stay on the P14-P15 pair.
  spanId: "P15-P16-FORE", poleId: "P15", owner: catv, wireId: "P14-CATV",
  existingHOA: "18'10\"", existingHOAChange: "21'", midspan: "16'8\""
}));
S.upsertSpanComm(S.createSpanComm({
  spanId: "P15-P16-FORE", poleId: "P16", owner: catv, wireId: "P16-P15-CATV",
  existingHOA: "20'10\"", existingHOAChange: "18'8\"", midspan: "16'8\""
}));

C.recalculateAll();
const referenceCatv = S.getSpanComm("P15-P14-BACK", "P15", catv, "P15-BACK-CATV");
const referenceTelco = S.getSpanComm("P15-P14-BACK", "P15", telco, "P15-BACK-TELCO");
assert.equal(referenceCatv.calculatedMidspan, "17'3\"", "P15-P14 CATV must use the P14-P15 measured midspan and its own movement");
assert.equal(referenceTelco.calculatedMidspan, "16'4\"", "P15-P14 Telco must use the P14-P15 measured midspan and its own movement");
assert.equal(C.displayMidspanForComm(referenceCatv), "", "a reciprocal Back Span must not display the source Midspan at the receiving pole");
assert.equal(C.displayMidspanForComm(referenceTelco), "", "a reciprocal Back Span Telco must not display the source Midspan at the receiving pole");
assert.equal(referenceCatv.flaggingStatus, "OK", "reciprocal calculated rows must not duplicate Fore Span violations");
assert.equal(S.getSpanComm("P14-P15-FORE", "P14", catv, "P14-CATV").calculatedMidspan, "17'3\"", "P14-P15 CATV must use its physical reciprocal instead of P15-P16 when wireId is reused");
assert.equal(S.getSpanComm("P14-P15-FORE", "P14", telco, "P14-TELCO").calculatedMidspan, "16'4\"", "P14-P15 Telco must use its physical reciprocal movement");
assert.equal(S.getSpanComm("P15-P16-FORE", "P15", catv, "P14-CATV").calculatedMidspan, "16'8\"", "P15-P16 must keep its own cable midspan calculation");

// Endpoint rows with the same wireId are one physical cable, not two comms.
// The endpoint without an imported Midspan mirrors the source calculation and
// must not generate a false zero-separation violation.
S.resetState();
S.applyProjectProfile("METRONET");
S.updateSetting("proposedOwner", "MidAm");
S.upsertPole(S.createPole({ poleId: "P77", lowPower: "23'6\"", maxCommHeight: "20'2\"" }));
S.upsertPole(S.createPole({ poleId: "P16", lowPower: "23'6\"", maxCommHeight: "20'2\"" }));
S.upsertSpan(S.createSpan("P77-P16", "P77", "P16", "N", "", {
  type: "Fore Span", rawType: "Fore Span", environmentClearance: "15'6\"", midspanMaxCommHeight: "21'7\""
}));
const sharedWire = "SHARED-CATV-WIRE";
S.upsertSpanComm(S.createSpanComm({
  spanId: "P77-P16", poleId: "P77", owner: catv, wireId: sharedWire,
  existingHOA: "20'10\"", existingHOAChange: "18'10\"", midspan: "19'1\"", ocalcMS: "19'1\""
}));
S.upsertSpanComm(S.createSpanComm({
  spanId: "P77-P16", poleId: "P16", owner: catv, wireId: sharedWire,
  existingHOA: "20'6\"", existingHOAChange: "18'2\""
}));
C.recalculateAll();
const p77Catv = S.getSpanComm("P77-P16", "P77", catv, sharedWire);
const p16Catv = S.getSpanComm("P77-P16", "P16", catv, sharedWire);
assert.equal(p77Catv.calculatedMidspan, "16'11\"", "shared cable MS must include both endpoint movements");
assert.equal(p16Catv.calculatedMidspan, "16'11\"", "the other endpoint must mirror the shared cable MS");
assert.equal(p77Catv.flaggingStatus, "OK", "a shared cable must not flag zero MS separation against itself");
assert.equal(p16Catv.flaggingStatus, "OK", "the mirrored endpoint must not flag zero MS separation against itself");
assert.equal(p16Catv.midspan, "", "the mirrored endpoint must keep its imported Midspan blank");
assert.equal(C.displayMidspanForComm(p16Catv), "", "the mirrored endpoint must not display the source Midspan as a second cable");

// A transferred attachment's old HOA belongs to the previous pole. It must
// not be used as a bolt point on the receiving pole, while the new HOA still
// participates in normal clearances against other attachments.
S.resetState();
S.applyProjectProfile("METRONET");
S.updateSetting("proposedOwner", "MidAm");
S.upsertPole(S.createPole({ poleId: "TRANSFER-POLE", lowPower: "25'", maxCommHeight: "22'" }));
S.upsertSpan(S.createSpan("TRANSFER-SPAN", "TRANSFER-POLE", "REMOTE", "N", "", {
  type: "Fore Span", rawType: "Fore Span", environmentClearance: "15'6\"", midspanMaxCommHeight: "21'7\""
}));
S.upsertSpanComm(S.createSpanComm({
  spanId: "TRANSFER-SPAN", poleId: "TRANSFER-POLE", owner: catv, wireId: "TRANSFER-WIRE",
  existingHOA: "20'6\"", existingHOAChange: "20'4\"", transferToNewPole: true
}));
C.recalculateAll();
const transferred = S.getSpanComm("TRANSFER-SPAN", "TRANSFER-POLE", catv, "TRANSFER-WIRE");
assert.equal(transferred.flaggingStatus, "OK", "a transferred comm old HOA must not create a bolt-bolt violation on the new pole");

console.log("Back Span calculation tests passed.");
