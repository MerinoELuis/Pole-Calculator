"use strict";

const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");

function classList(initial = []) {
  const values = new Set(initial);
  return {
    add: value => values.add(value),
    remove: value => values.delete(value),
    [Symbol.iterator]: function* () { yield* values; },
    values: () => Array.from(values)
  };
}

const window = { document: null, UiDomContract: null };
const sourcePath = path.join(__dirname, "..", "js", "ui", "span-colors.js");
vm.runInNewContext(fs.readFileSync(sourcePath, "utf8"), { window, Set, Map, Number, String, Array }, { filename: sourcePath });
const api = window.SpanColorUI;

const spans = {
  "A-hidden": { spanId: "A-hidden", spanIndex: "0" },
  "A-1": { spanId: "A-1", spanIndex: "1" },
  "A-2": { spanId: "A-2", spanIndex: "2" },
  "B-1": { spanId: "B-1", spanIndex: "1" },
  "B-2": { spanId: "B-2", spanIndex: "2" },
  "manual-B": { spanId: "manual-B", sourceSpanId: "B-2", spanIndex: "99" }
};
const connected = {
  P01: [spans["A-hidden"], spans["A-1"], spans["A-2"]],
  P02: [spans["B-1"], spans["B-2"]]
};
const store = {
  getSpan: id => spans[id] || null,
  getConnectedSpans: poleId => connected[poleId] || []
};

assert.deepEqual(Array.from(api.sortVisibleSpanIds(store, "P01", ["A-2", "A-1"])), ["A-1", "A-2"]);
assert.equal(api.classForVisibleIndex(0), "span-color-0");
assert.equal(api.classForVisibleIndex(1), "span-color-1");
assert.equal(api.classForVisibleIndex(5), "span-color-5");
assert.equal(api.classForVisibleIndex(9), "span-color-9");
assert.equal(api.classForVisibleIndex(10), "span-color-0");
assert.equal(api.classForVisibleIndex(api.sortVisibleSpanIds(store, "P02", ["manual-B", "B-1"]).indexOf("B-1")), "span-color-0");
assert.equal(api.physicalSpanId(store, "manual-B"), "B-2");

const stable = { dataset: { spanId: "S-blue" } };
assert.equal(api.firstSpanId(stable), "S-blue", "stable data-span-id must be authoritative");

const target = { classList: classList(["span-color-4", "other-class"]) };
api.resetCard({
  dataset: { poleId: "P01" },
  querySelectorAll(selector) {
    if (selector === ".span-proposed-table tbody tr") return [];
    if (selector === ".comm-movement-table tbody > tr") return [];
    return [];
  }
}, store);
assert.deepEqual(target.classList.values(), ["span-color-4", "other-class"], "unrelated elements are untouched");

console.log("span color UI tests passed");
