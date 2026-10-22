"use strict";

const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");

const sandbox = { window: {} };
for (const file of ["domain/height-math.js", "import/import-normalizers.js", "import/excel-import.js"]) {
  const modulePath = path.join(__dirname, "..", "js", file);
  vm.runInNewContext(fs.readFileSync(modulePath, "utf8"), sandbox, { filename: modulePath });
}

const ExcelImport = sandbox.window.ExcelImport;
assert.equal(ExcelImport.hasBirthmarkValue("1"), true);
assert.equal(ExcelImport.hasBirthmarkValue("photo_001.jpg"), true);
assert.equal(ExcelImport.hasBirthmarkValue("0"), false);
assert.equal(ExcelImport.hasBirthmarkValue(""), false);

sandbox.window.AppStore = {
  getState: () => ({
    settings: { projectProfile: "METRONET", metronetWI: "MIDAM", proposedOwner: "MidAm" }
  })
};
const importedRow = ExcelImport.buildPoleClassChecks([{
  Id: "P-001",
  "Tip.display": "28'",
  Circumference: "35",
  Type: "Southern Pine > 4 > 35",
  "Birthmark Photos": 1
}])[0];
assert.equal(importedRow.birthmarkExists, true, "Birthmark Photos should activate the Pole Type Check toggle");
assert.equal(importedRow.birthmarkActive, true, "an imported Birthmark should start active");
assert.equal(importedRow.birthmarkClassOverride, true, "MidAm birthmark should override a class mismatch");

const manuallyDisabledBirthmark = ExcelImport.recalculatePoleClassCheck({
  ...importedRow,
  birthmarkActive: false
});
assert.equal(manuallyDisabledBirthmark.birthmarkExists, true, "manual disable must retain the imported Birthmark fact");
assert.equal(manuallyDisabledBirthmark.birthmarkActive, false, "Birthmark can be manually disabled");
assert.match(manuallyDisabledBirthmark.status, /Class mismatch/, "disabling Birthmark must restore the class check");

const manuallyEnabledBirthmark = ExcelImport.recalculatePoleClassCheck({
  tip: "32'",
  importedType: "Southern Pine > 4 > 40",
  circumference: "",
  projectProfile: "METRONET",
  metronetWI: "MIDAM",
  birthmarkExists: false,
  birthmarkActive: true
});
assert.equal(manuallyEnabledBirthmark.birthmarkActive, true, "Birthmark can be manually enabled without an imported value");
assert.equal(manuallyEnabledBirthmark.status, "OK", "manual Birthmark authority must apply to the class review");

const birthmarkWithoutCircumference = ExcelImport.recalculatePoleClassCheck({
  tip: "32'",
  importedType: "Southern Pine > 4 > 40",
  circumference: "",
  projectProfile: "METRONET",
  metronetWI: "MIDAM",
  birthmarkExists: true
});
assert.equal(birthmarkWithoutCircumference.status, "OK", "Birthmark must make class valid even without imported circumference");

const mismatchWithoutBirthmark = ExcelImport.recalculatePoleClassCheck({
  tip: "28'",
  importedType: "Southern Pine > 4 > 35",
  circumference: "35",
  projectProfile: "METRONET",
  metronetWI: "MIDAM",
  birthmarkExists: false
});
assert.match(mismatchWithoutBirthmark.status, /Class mismatch/);

const mismatchWithBirthmark = ExcelImport.recalculatePoleClassCheck({
  tip: "28'",
  importedType: "Southern Pine > 4 > 35",
  circumference: "35",
  projectProfile: "METRONET",
  metronetWI: "MIDAM",
  birthmark: "1"
});
assert.equal(mismatchWithBirthmark.birthmarkExists, true);
assert.equal(mismatchWithBirthmark.birthmarkClassOverride, true);
assert.equal(mismatchWithBirthmark.status, "OK");

const comedMismatch = ExcelImport.recalculatePoleClassCheck({
  tip: "28'",
  importedType: "Southern Pine > 4 > 35",
  circumference: "35",
  projectProfile: "METRONET",
  metronetWI: "COMED",
  birthmark: "1"
});
assert.match(comedMismatch.status, /Class mismatch/);

console.log("Birthmark Pole Type Check tests passed.");
