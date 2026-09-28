"use strict";

const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");

const sandbox = { window: {} };
for (const file of ["height-utils.js", "excel-import.js"]) {
  const modulePath = path.join(__dirname, "..", "js", file);
  vm.runInNewContext(fs.readFileSync(modulePath, "utf8"), sandbox, { filename: modulePath });
}

const ExcelImport = sandbox.window.ExcelImport;
assert.deepEqual(Array.from(ExcelImport.OPPD_POLE_TYPES), [
  "30-6", "30-4", "35-6", "35-4", "35-2", "40-6", "40-4", "40-2",
  "45-4", "45-2", "50-4", "50-2", "55-4", "55-2", "60-4", "60-2"
]);

const roundedClass = ExcelImport.recalculatePoleClassCheck({
  tip: "28'",
  importedType: "Southern Pine > 3 > 35",
  projectProfile: "OLSSON_OPPD",
  circumference: "35"
});
assert.equal(roundedClass.calculatedHeight, 35);
assert.equal(roundedClass.calculatedClass, "3");
assert.equal(roundedClass.recommendedClass, "4", "OPPD must round class 3 down to class 4");
assert.equal(roundedClass.expectedType, "Southern Pine > 4 > 35");

const matchingRoundedType = ExcelImport.recalculatePoleClassCheck({
  tip: "28'",
  importedType: "Southern Pine > 4 > 35",
  projectProfile: "OLSSON_OPPD",
  circumference: "35"
});
assert.equal(matchingRoundedType.status, "OK", "an imported type equal to the OPPD expected type must be OK");

const roundedHeight = ExcelImport.recalculatePoleClassCheck({
  tip: "32'",
  importedType: "Southern Pine > 1 > 40",
  projectProfile: "OLSSON_OPPD",
  circumference: "42"
});
assert.equal(roundedHeight.calculatedHeight, 40);
assert.equal(roundedHeight.recommendedClass, "2", "OPPD must round class 1 down to class 2");
assert.equal(roundedHeight.expectedType, "Southern Pine > 2 > 40");

const typeOnly = ExcelImport.recalculatePoleClassCheck({
  tip: "32'",
  importedType: "Southern Pine > 4 > 40",
  projectProfile: "OLSSON_OPPD",
  circumference: ""
});
assert.equal(typeOnly.status, "OK", "Olsson must accept a valid Type when circumference is unavailable");

const invalidTypeOnly = ExcelImport.recalculatePoleClassCheck({
  tip: "32'",
  importedType: "Southern Pine > 5 > 40",
  projectProfile: "OLSSON_OPPD",
  circumference: ""
});
assert.equal(invalidTypeOnly.status, "OK", "Olsson must ignore class when circumference is unavailable");

console.log("OPPD pole class recommendation tests passed.");
