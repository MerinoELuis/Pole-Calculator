"use strict";

const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");

function makeRow(values) {
  const row = { children: [] };
  row.children = values.map(value => ({
    textContent: value,
    parent: row,
    remove() {
      const index = this.parent.children.indexOf(this);
      if (index >= 0) this.parent.children.splice(index, 1);
    }
  }));
  return row;
}

function makeTable(headerValues, bodyRows) {
  const headerRow = makeRow(headerValues);
  const rows = [headerRow, ...bodyRows.map(makeRow)];
  return {
    rows,
    querySelectorAll(selector) {
      if (selector === "thead th") return headerRow.children;
      if (selector === "tr") return rows;
      return [];
    }
  };
}

const window = {
  document: null,
  console,
  setTimeout,
  queueMicrotask
};

const sourcePath = path.join(__dirname, "..", "js", "comm-table-columns.js");
vm.runInNewContext(fs.readFileSync(sourcePath, "utf8"), { window }, { filename: sourcePath });

const api = window.CommTableColumns;
assert.ok(api);
assert.equal(
  api.findColumnIndex([
    { textContent: "Owner/Comm" },
    { textContent: "Existing HOA" },
    { textContent: "HOA Change" },
    { textContent: " Other   Pole HOA " },
    { textContent: "Span" }
  ]),
  3
);

const table = makeTable(
  ["Owner/Comm", "Existing HOA", "HOA Change", "Other Pole HOA", "Span", "Midspan"],
  [
    ["CATV", "18'", "19'", "20'", "P1 → P2", "15'"],
    ["CTL", "17'", "18'", "19'", "P1 → P3", "14'6\""]
  ]
);

assert.equal(api.removeColumnFromTable(table, 3), true);
assert.deepEqual(table.rows[0].children.map(cell => cell.textContent), [
  "Owner/Comm", "Existing HOA", "HOA Change", "Span", "Midspan"
]);
assert.deepEqual(table.rows[1].children.map(cell => cell.textContent), [
  "CATV", "18'", "19'", "P1 → P2", "15'"
]);
assert.deepEqual(table.rows[2].children.map(cell => cell.textContent), [
  "CTL", "17'", "18'", "P1 → P3", "14'6\""
]);

const noTarget = makeTable(["Owner/Comm", "Existing HOA", "Span"], [["CATV", "18'", "P1 → P2"]]);
assert.equal(api.findColumnIndex(noTarget.querySelectorAll("thead th")), -1);
assert.equal(api.removeColumnFromTable(noTarget, -1), false);

console.log("comm table column tests passed");
