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
S.applyProjectProfile("OLSSON_OPPD");
S.upsertPole(S.createPole({ poleId: "P1", lowPower: "30'" }));
S.upsertPole(S.createPole({ poleId: "P2", lowPower: "30'" }));
S.upsertSpan(S.createSpan("S1", "P1", "P2", "E", "", {
  type: "Fore Span",
  midspanLowPower: "25'",
  sourceMidspanLowPower: "25'"
}));
S.addSpanPower(S.createSpanPower({
  spanId: "S1",
  poleId: "P1",
  label: "Secondary",
  attachmentHeight: "30'",
  midspan: "26'",
  wireId: "POWER-1"
}));

C.recalculateSpan("S1");
assert.equal(S.getSpan("S1").midspanLowPower, "25'", "Olsson must keep the lower Span low-power midspan when power is higher");

const powerKey = Object.keys(S.getState().spanPower)[0];
S.getState().spanPower[powerKey].midspan = "24'";
C.recalculateSpan("S1");
assert.equal(S.getSpan("S1").midspanLowPower, "24'", "Olsson must use a lower imported power midspan");

S.getState().spanPower[powerKey].midspan = "26'";
C.recalculateSpan("S1");
assert.equal(S.getSpan("S1").midspanLowPower, "25'", "Olsson must retain the original Span low-power baseline after edits");

console.log("Olsson low-power midspan tests passed.");
