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
S.applyProjectProfile("INTEC");
S.upsertPole(S.createPole({ poleId: "P1", poleInsetHeight: "18'" }));
S.upsertPole(S.createPole({ poleId: "P2" }));
S.upsertSpan(S.createSpan("S1", "P1", "P2", "E", "", { type: "Fore Span", lengthDisplay: "150'" }));
S.upsertSpanSide(S.createSpanSide({ spanId: "S1", poleId: "P1", proposedHOA: "20'", proposedHOAChange: "19'" }));
S.addSpanPower(S.createSpanPower({ spanId: "S1", poleId: "P1", wireId: "N1", size: "Neutral", midspan: "30'" }));
S.addSpanPower(S.createSpanPower({ spanId: "S1", poleId: "P1", wireId: "P1", size: "Primary", midspan: "32'" }));

const details = C.getPoleInsetDetails("P1");
assert.equal(details.maxHeightAtMidspan, "26'8\"", "Pole Inset must use the lower 40/43-inch power-midspan ceiling");
assert.equal(details.selectedHeight, "18'", "Pole Inset must expose the manually selected HOA");
assert.equal(details.endDropToPole, "-2'", "Pole Inset must show the End Drop to the local pole");
assert.equal(details.endDropToNextPole, "-1'", "Pole Inset must show the End Drop from the next pole toward the inset");
assert.equal(details.spanId, "S1", "Pole Inset must identify the proposed span used for End Drops");

S.upsertPole({ ...S.getPole("P1"), poleInsetHeight: "" });
assert.equal(C.getPoleInsetDetails("P1").endDropToPole, "", "End Drop to Pole must stay blank without a manual inset HOA");
assert.equal(C.getPoleInsetDetails("P1").endDropToNextPole, "", "End Drop to Next Pole must stay blank without a manual inset HOA");

S.resetState();
S.applyProjectProfile("INTEC");
S.upsertPole(S.createPole({ poleId: "P3" }));
S.upsertPole(S.createPole({ poleId: "P2" }));
S.upsertPole(S.createPole({ poleId: "P4" }));
S.upsertSpan(S.createSpan("LOWER", "P2", "P3", "N", "", { type: "Fore Span", lengthDisplay: "150'" }));
S.upsertSpan(S.createSpan("SELECTED", "P3", "P4", "S", "", { type: "Fore Span", lengthDisplay: "150'" }));
S.upsertSpanSide(S.createSpanSide({ spanId: "SELECTED", poleId: "P3", proposedHOA: "23'10\"" }));
S.addSpanPower(S.createSpanPower({ spanId: "LOWER", poleId: "P3", wireId: "LOWER-N", size: "Neutral", midspan: "20'11\"" }));
S.addSpanPower(S.createSpanPower({ spanId: "SELECTED", poleId: "P3", wireId: "SELECTED-N", size: "Neutral", midspan: "23'4\"" }));
assert.equal(C.getPoleInsetDetails("P3").maxHeightAtMidspan, "20'", "Pole Inset must use the selected proposed span power midspan");

S.resetState();
S.applyProjectProfile("INTEC");
S.upsertPole(S.createPole({ poleId: "EMPTY" }));
assert.equal(C.getPoleInsetDetails("EMPTY").maxHeightAtMidspan, "", "Missing Power Midspan must not fabricate a Pole Inset ceiling");

console.log("pole-inset.test.js passed");
