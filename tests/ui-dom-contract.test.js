"use strict";

const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");

const window = { document: null };
const sourcePath = path.join(__dirname, "..", "js", "ui", "dom-contract.js");
vm.runInNewContext(fs.readFileSync(sourcePath, "utf8"), { window }, { filename: sourcePath });
const api = window.UiDomContract;

assert.equal(api.normalizeLabel("  Other   Pole HOA "), "other pole hoa");
assert.equal(api.columnKey("OWNER/COMM"), "owner-comm");
assert.equal(api.columnKey("Environment Clearance"), "environment-clearance");

const spanRows = [
  { dataset: { spanId: "S1" } },
  { dataset: { span: "S2" } }
];
const midspanRows = [
  { dataset: { spanId: "S2" } },
  { dataset: { spanId: "S1" } }
];
const tableRow = {
  querySelectorAll(selector) {
    if (selector === ".comm-span-list .comm-span-row") return spanRows;
    if (selector === ".comm-midspan-list[data-midspan-list] .comm-midspan-value") return midspanRows;
    if (selector === ".comm-midspan-list .comm-midspan-value") return [];
    return [];
  }
};
const pairs = api.pairCommRows(tableRow);
assert.equal(pairs.length, 2);
assert.equal(pairs[0].spanId, "S1");
assert.equal(pairs[0].midspanRow, midspanRows[1]);
assert.equal(pairs[1].spanId, "S2");
assert.equal(pairs[1].midspanRow, midspanRows[0]);

console.log("DOM contract tests passed");
