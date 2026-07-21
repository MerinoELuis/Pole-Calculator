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

function seedSpan(profile, spanId, lengthDisplay, proposedHOA) {
  S.resetState();
  S.applyProjectProfile(profile);
  S.upsertPole(S.createPole({ poleId: "P1", lowPower: "35'" }));
  S.upsertPole(S.createPole({ poleId: "P2", lowPower: "35'" }));
  S.upsertSpan(S.createSpan(spanId, "P1", "P2", "E", "", {
    type: "Fore Span",
    rawType: "Fore Span",
    lengthDisplay
  }));
  S.upsertSpanSide(S.createSpanSide({ spanId, poleId: "P1", proposedHOA }));
}

seedSpan("INTEC", "WITH-MS", "150'", "22'");
S.upsertSpanComm(S.createSpanComm({
  spanId: "WITH-MS",
  poleId: "P1",
  owner: "COMMUNICATION > CATV",
  existingHOA: "20'",
  midspan: "17'"
}));
S.upsertSpan(S.createSpan("UNRELATED", "P1", "P3", "N", "", { type: "Other", rawType: "Other", lengthDisplay: "100'" }));
S.upsertSpanComm(S.createSpanComm({
  spanId: "UNRELATED",
  poleId: "P1",
  owner: "COMMUNICATION > Fiber",
  existingHOA: "30'",
  midspan: "29'"
}));
let side = S.getSpanSide("WITH-MS", "P1");
let span = S.getSpan("WITH-MS");
assert.equal(C.calculateProposedMidspanBase(side, span), 18 * 12, "Wecom Proposed MS must be 12 inches above the highest comm midspan");
C.calculateSpanSideMidspan("WITH-MS", "P1");
assert.equal(S.getSpanSide("WITH-MS", "P1").msProposed, "18'", "the automatic Wecom base must be persisted for display and validation");

S.upsertSpan(S.createSpan("POWER-ONLY", "P1", "P2", "E", "", {
  type: "Fore Span",
  rawType: "Fore Span",
  lengthDisplay: "147'4\""
}));
S.addSpanPower(S.createSpanPower({
  spanId: "POWER-ONLY",
  poleId: "P1",
  wireId: "NEUTRAL-1",
  size: "Neutral",
  midspan: "25'10\""
}));
assert.equal(C.spanHasRealMidspan("POWER-ONLY"), true, "an INTEC Fore Span with only a Power midspan must be eligible for Proposed by Span");

S.upsertSpan(S.createSpan("EMPTY-FORE", "P1", "P2", "E", "", {
  type: "Fore Span",
  rawType: "Fore Span",
  lengthDisplay: "140'"
}));
assert.equal(C.spanHasRealMidspan("EMPTY-FORE"), false, "a truly empty Fore Span must remain hidden under the INTEC profile");

seedSpan("METRONET", "NO-MS", "148'5\"", "22'");
side = S.getSpanSide("NO-MS", "P1");
span = S.getSpan("NO-MS");
assert.equal(C.getEstimatedSagInches(span), 18, "148 feet 5 inches must round to the 150-foot sag bucket");
assert.equal(C.calculateProposedMidspanBase(side, span), 20 * 12 + 6, "MidAm Proposed MS without a comm midspan must subtract 18 inches of estimated sag");
C.calculateSpanSideMidspan("NO-MS", "P1");
assert.equal(S.getSpanSide("NO-MS", "P1").msProposed, "20'6\"", "the estimated MidAm base must be persisted for display and validation");

[
  ["100'", 12],
  ["150'", 18],
  ["200'", 24],
  ["250'", 30]
].forEach(([lengthDisplay, expectedSag]) => {
  assert.equal(C.getEstimatedSagInches({ lengthDisplay }), expectedSag, `${lengthDisplay} must produce ${expectedSag} inches of sag`);
});

S.upsertSpanSide({ ...side, ocalcMS: "19.25" });
side = S.getSpanSide("NO-MS", "P1");
assert.equal(C.calculateProposedMidspanBase(side, span), 19 * 12 + 3, "manual O-CALC MS must override the automatic Proposed calculation");

seedSpan("METRONET", "FORE", "101'6\"", "20'");
S.upsertSpan(S.createSpan("BACK", "P2", "P1", "W", "", {
  type: "Back Span",
  rawType: "Back Span",
  lengthDisplay: "101'6\""
}));
S.upsertSpanComm(S.createSpanComm({
  spanId: "FORE",
  poleId: "P1",
  owner: "COMMUNICATION > Fiber",
  existingHOA: "19'"
}));
S.upsertSpanComm(S.createSpanComm({
  spanId: "BACK",
  poleId: "P2",
  owner: "COMMUNICATION > Fiber",
  existingHOA: "20'2\"",
  midspan: "18'10\""
}));
side = S.getSpanSide("FORE", "P1");
span = S.getSpan("FORE");
assert.equal(C.calculateProposedMidspanBase(side, span), 19 * 12 + 10, "MidAm must use the reciprocal same-connection midspan before falling back to span sag");

console.log("Proposed midspan fallback tests passed.");
