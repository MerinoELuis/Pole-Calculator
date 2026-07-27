"use strict";

const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");

function makeRow(values, keys = []) {
  const row = { children: [] };
  row.children = values.map((value, index) => ({
    textContent: value,
    dataset: keys[index] ? { column: keys[index] } : {},
    parent: row,
    remove() {
      const position = this.parent.children.indexOf(this);
      if (position >= 0) this.parent.children.splice(position, 1);
    }
  }));
  return row;
}

function makeTable(headerValues, bodyRows, keys = []) {
  const headerRow = makeRow(headerValues, keys);
  const rows = [headerRow, ...bodyRows.map(row => makeRow(row, keys))];
  return {
    rows,
    querySelectorAll(selector) {
      if (selector === "thead th") return headerRow.children;
      if (selector === "tr") return rows;
      return [];
    }
  };
}

const window = { document: null };
const sourcePath = path.join(__dirname, "..", "js", "ui", "comm-table.js");
vm.runInNewContext(fs.readFileSync(sourcePath, "utf8"), { window }, { filename: sourcePath });
const api = window.CommTableUI;

assert.equal(api.findColumnIndex([
  { textContent: "Owner/Comm", dataset: { column: "owner-comm" } },
  { textContent: "Renamed label", dataset: { column: "other-pole-hoa" } },
  { textContent: "Span", dataset: { column: "span" } }
]), 1, "stable data-column must take priority over visible text");

assert.equal(api.findColumnIndex([
  { textContent: "Owner/Comm", dataset: {} },
  { textContent: " Other   Pole HOA ", dataset: {} }
]), 1, "legacy label remains a safe fallback");

const keys = ["owner-comm", "existing-hoa", "hoa-change", "other-pole-hoa", "span", "midspan"];
const table = makeTable(
  ["Owner/Comm", "Existing HOA", "HOA Change", "Other Pole HOA", "Span", "Midspan"],
  [
    ["CATV", "18'", "19'", "20'", "P1 → P2", "15'"],
    ["CTL", "17'", "18'", "19'", "P1 → P3", "14'6\""]
  ],
  keys
);

assert.equal(api.removeColumnFromTable(table, api.findColumnIndex(table.querySelectorAll("thead th"))), true);
assert.deepEqual(table.rows[0].children.map(cell => cell.dataset.column), [
  "owner-comm", "existing-hoa", "hoa-change", "span", "midspan"
]);
assert.equal(api.removeColumnFromTable(table, -1), false);

console.log("comm table UI tests passed");
