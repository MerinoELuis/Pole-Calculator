"use strict";

const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");

const window = {};
const sandbox = { window, String, Array, Object, Number, Set, RegExp, Math };
const modulePath = path.join(__dirname, "..", "js", "import", "import-normalizers.js");
vm.runInNewContext(fs.readFileSync(modulePath, "utf8"), sandbox, { filename: modulePath });

const heights = {
  parseHeight(value) {
    const match = String(value || "").match(/^(\d+)'(?:\s*(\d+)")?$/);
    return match ? Number(match[1]) * 12 + Number(match[2] || 0) : null;
  },
  formatHeight(inches) {
    return `${Math.floor(inches / 12)}'${inches % 12 ? ` ${inches % 12}"` : ""}`;
  }
};
const normalizers = window.ImportNormalizers.create(heights);

const rows = normalizers.rowsToObjects([
  ["Id", "Sequence", "Mid Span Height.display"],
  [1, 1, "18' 6\""]
]);
normalizers.normalizeNumericProjectIdentifiers(rows);
assert.equal(rows[0].Id, "001");
assert.equal(rows[0].Sequence, "001");
assert.equal(normalizers.normalizeHeaderName("Mid Span Height.display"), "midspanheightdisplay");
assert.equal(normalizers.pick({ Id: "P001" }, ["Midspan"], { minPartialLength: 3 }), "");
assert.equal(normalizers.heightFromRow(rows[0], ["Mid Span Height.display"], ["Mid Span Height"]), "18' 6\"");
const workbook = { SheetNames: ["Pole.Collection"], Sheets: { "Pole.Collection": [] } };
assert.equal(normalizers.findSheet(workbook, ["Pole Collection"]), workbook.Sheets["Pole.Collection"]);

console.log("import normalization tests passed");
