"use strict";

const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");

const appPath = path.join(__dirname, "..", "js", "app.js");
const appSource = fs.readFileSync(appPath, "utf8").replace(
  'document.addEventListener("DOMContentLoaded", init);',
  "global.__mergeImportedUpdate = mergeImportedUpdate;"
);
const sandbox = {
  window: {
    AppStore: {},
    HeightUtils: {},
    ExcelImport: { recalculatePoleClassCheck: row => row }
  },
  console,
  Date,
  Map,
  Set,
  JSON
};
vm.runInNewContext(appSource, sandbox, { filename: appPath });
const merge = sandbox.window.__mergeImportedUpdate;

const previous = {
  poles: { P1: { poleId: "P1", lowPower: "30'", poleHeight: "40'", notes: "Saved note", ugActive: false, pcoActive: false } },
  spans: {
    S1: { spanId: "S1", fromPole: "P1", toPole: "P2", lengthDisplay: "100'", environment: "STREET" },
    S2: { spanId: "S2", fromPole: "P1", toPole: "P3", lengthDisplay: "80'", environment: "ALLEY" }
  },
  spanSides: { S1__P1: { spanId: "S1", poleId: "P1", proposedHOA: "22'", proposedHOAChange: "", ocalcMS: "" } },
  spanComms: {
    C1: { spanId: "S1", poleId: "P1", owner: "CATV", ownerBase: "CATV", wireId: "W1", existingHOA: "20'", midspan: "16'", size: "Old size" },
    C2: { spanId: "S2", poleId: "P1", owner: "CTL", ownerBase: "CTL", wireId: "W2", existingHOA: "19'", midspan: "15'", size: "Telco" }
  },
  spanPower: { PW1: { spanId: "S1", poleId: "P1", owner: "APS", wireId: "PW1", size: "Primary", attachmentHeight: "35'", midspan: "29'" } },
  makeReadyReferences: [],
  poleClassChecks: [],
  settings: { projectProfile: "INTEC", fiberSizes: {} },
  ui: {},
  excelReviewSource: { collection: { rows: [{ Id: "OLD" }] } }
};

const imported = {
  poles: { P1: { poleId: "P1", lowPower: "", poleHeight: "45'", notes: "", ugActive: false, pcoActive: false } },
  spans: { S1: { spanId: "S1", fromPole: "P1", toPole: "P2", lengthDisplay: "", environment: "ALLEY" } },
  spanSides: { S1__P1: { spanId: "S1", poleId: "P1", proposedHOA: "", proposedHOAChange: "", ocalcMS: "" } },
  spanComms: { C1: { spanId: "S1", poleId: "P1", owner: "CATV", ownerBase: "CATV", wireId: "W1", existingHOA: "", midspan: "", size: "New size" } },
  spanPower: { PW1: { spanId: "S1", poleId: "P1", owner: "APS", wireId: "PW1", size: "Primary", attachmentHeight: "", midspan: "" } },
  makeReadyReferences: [],
  poleClassChecks: [],
  settings: { projectProfile: "INTEC", fiberSizes: {} },
  ui: {},
  excelReviewSource: { collection: { rows: [{ Id: "NEW", "Low Power Attachment.display": "" }] } }
};

const result = merge(previous, imported);
assert.equal(result.poles.P1.lowPower, "30'", "blank updated pole value must retain the prior value");
assert.equal(result.poles.P1.poleHeight, "45'", "non-empty updated pole value must win");
assert.equal(result.spans.S1.lengthDisplay, "100'", "blank updated span value must retain the prior value");
assert.equal(result.spans.S1.environment, "ALLEY", "non-empty updated span value must win");
assert.equal(result.spans.S2, undefined, "a span omitted by Update Data must not be recreated");
assert.equal(result.spanSides.S1__P1.proposedHOA, "22'", "saved Proposed must be retained");
assert.equal(result.spanComms.C1.existingHOA, "20'", "blank Existing HOA must retain the prior value");
assert.equal(result.spanComms.C1.midspan, "16'", "blank imported midspan must retain the prior value");
assert.equal(result.spanComms.C1.size, "New size", "new non-empty comm data must win");
assert.equal(result.spanComms.C2, undefined, "an omitted comm without user work must not create an empty row");
assert.equal(result.spanPower.PW1.midspan, "29'", "blank power midspan must retain the prior value");
assert.equal(result.excelReviewSource.collection.rows[0].Id, "NEW", "raw review source must remain the new workbook snapshot");
assert.ok(result.updateDiagnostics.blankValuesPreserved > 0);

console.log("Update Data non-destructive merge tests passed.");
