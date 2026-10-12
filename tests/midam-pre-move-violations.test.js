"use strict";

const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");

const window = {
  localStorage: { getItem: () => null, setItem: () => {} },
  MRLogic: { generateMRForPole: () => {}, generateAllMR: () => {} },
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
S.upsertPole(S.createPole({ poleId: "P1", maxCommHeight: "20'" }));
S.upsertPole(S.createPole({ poleId: "P2" }));
S.upsertSpan(S.createSpan("FORE", "P1", "P2", "E", "", {
  type: "Fore Span", rawType: "Fore Span", midspanMaxCommHeight: "18'"
}));
S.upsertSpan(S.createSpan("BACK", "P1", "P2", "W", "", {
  type: "Back Span", rawType: "Back Span"
}));
S.upsertSpanComm(S.createSpanComm({
  spanId: "FORE", poleId: "P1", owner: "COMMUNICATION > Telco", ownerBase: "COMMUNICATION > Telco",
  existingHOA: "21'", midspan: "19'"
}));
S.upsertSpanComm(S.createSpanComm({
  spanId: "FORE", poleId: "P1", owner: "COMMUNICATION > CATV", ownerBase: "COMMUNICATION > CATV",
  existingHOA: "20'6\"", midspan: "18'2\""
}));
S.upsertSpanComm(S.createSpanComm({
  spanId: "BACK", poleId: "P1", owner: "COMMUNICATION > Telco", ownerBase: "COMMUNICATION > Telco",
  existingHOA: "21'"
}));

const violations = C.getMidamPreMoveViolations("P1");
assert.ok(violations.some(item => /On-pole violation:.*Telco.*21'/.test(item.message)), "must describe the imported on-pole violation");
assert.ok(violations.some(item => /Midspan violation:.*Telco.*19'.*E/.test(item.message)), "must describe the Fore Span midspan violation and direction");
assert.ok(violations.every(item => item.scope === "pole" || item.scope === "midspan"));
assert.ok(!violations.some(item => /separation 0/.test(item.message)), "reference rows must not create duplicate on-pole violations");

S.applyProjectProfile("INTEC");
assert.equal(C.getMidamPreMoveViolations("P1").length, 0, "the section must be empty outside MidAm");

console.log("MidAm pre-move violation tests passed.");
