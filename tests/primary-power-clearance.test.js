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

function seedPole({ poleId, lowPower, primaryHeight = "" }) {
  S.upsertPole(S.createPole({ poleId, lowPower }));
  if (primaryHeight) {
    S.addSpanPower(S.createSpanPower({
      spanId: `${poleId}-SPAN`,
      poleId,
      label: "Primary",
      attachmentHeight: primaryHeight,
      midspan: "25'",
      wireId: `${poleId}-PRIMARY`
    }));
  }
}

S.resetState();
S.applyProjectProfile("INTEC");
seedPole({ poleId: "P1", lowPower: "31'10\"", primaryHeight: "31'10\"" });
S.upsertSpan(S.createSpan("P1-SPAN", "P1", "P2", "E", "", { type: "Fore Span" }));
S.upsertPole(S.createPole({ poleId: "P2", lowPower: "31'10\"" }));
C.recalculateAll();
assert.equal(
  S.getPole("P1").maxCommHeight,
  "28'3\"",
  "all comms must remain 43 inches below a Primary when Primary is the limiting power height"
);

S.upsertPole({
  ...S.getPole("P1"),
  lowPower: "30'",
  metadata: { ...S.getPole("P1").metadata, lowPowerBaseline: "30'" }
});
C.recalculateAll();
assert.equal(
  S.getPole("P1").maxCommHeight,
  "26'8\"",
  "the lower generic Low Power ceiling must still control when it is stricter"
);

S.upsertPole(S.createPole({ poleId: "P3", lowPower: "31'10\"" }));
C.recalculateAll();
assert.equal(
  S.getPole("P3").maxCommHeight,
  "28'6\"",
  "without a Primary row the existing 40-inch Low Power rule must remain"
);

console.log("Primary power clearance tests passed.");
