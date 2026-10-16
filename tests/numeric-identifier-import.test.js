"use strict";

const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");

const window = {
  localStorage: { getItem: () => null, setItem: () => {} },
  Calculations: { recalculateAll: () => {} }
};
const sandbox = { window, console, Date, Set, Map, Math, JSON, Object, Array, String, Number, Boolean, RegExp };
for (const file of ["height-utils.js", "project-config.js", "state.js", "excel-import.js"]) {
  const modulePath = path.join(__dirname, "..", "js", file);
  vm.runInNewContext(fs.readFileSync(modulePath, "utf8"), sandbox, { filename: modulePath });
}

const state = window.ExcelImport.importOriginalWorkbook({
  SheetNames: ["Collection", "Span", "Span.Wire"],
  Sheets: {
    Collection: [
      ["Id", "Sequence", "Owner", "Lowest Power.display"],
      [1, 1, "UTILITY > MidAm", "27'"]
    ],
    Span: [
      ["Id", "Span Id", "Type"],
      [1, "S1", "Fore Span"]
    ],
    "Span.Wire": [
      ["Id", "Span Id", "Owner", "Attachment Height.display"],
      [1, "S1", "COMMUNICATION > CATV", "20'"]
    ]
  }
}, "EXCEL_IAIE037_2026-10-16.xlsx");

assert.ok(state.poles["001"], "numeric Collection Id 1 must be restored to pole Id 001");
assert.equal(state.poles["001"].sequence, "001", "numeric Sequence 1 must be restored to 001");
assert.equal(state.spans.S1.fromPole, "001", "numeric span source Id must match the padded Collection Id");
assert.equal(state.excelReviewSource.collection.rows[0].Id, "001", "review data must retain the padded Id");
assert.equal(state.excelReviewSource.collection.rows[0].Sequence, "001", "review data must retain the padded Sequence");

console.log("Numeric project identifier import tests passed.");
