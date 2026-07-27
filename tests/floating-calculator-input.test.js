"use strict";

const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");

const source = fs.readFileSync(
  path.join(__dirname, "../js/floating-calculator.js"),
  "utf8"
);

const window = { document: null };
vm.runInNewContext(source, { window, console, String, Number });

const calculator = window.FloatingCalculator;

assert.equal(
  calculator.normalizeInchQuotes("12'6'' + 3'8''"),
  "12'6\" + 3'8\"",
  "Two apostrophes must be normalized to one inches quote."
);

assert.equal(
  calculator.normalizeInchQuotes("12'6\" + 3'8\""),
  "12'6\" + 3'8\"",
  "Existing inches quotes must remain unchanged."
);

assert.equal(
  calculator.normalizeInchQuotes("12'6'"),
  "12'6'",
  "A single apostrophe must not be changed."
);

const original = "12'6'' + 3'8''";
let selection = null;
const input = {
  value: original,
  selectionStart: original.length,
  selectionEnd: original.length,
  setSelectionRange(start, end) {
    selection = { start, end };
  }
};

assert.equal(
  calculator.normalizeExpressionInput(input),
  true,
  "Manual input containing two apostrophes must be normalized."
);
assert.equal(input.value, "12'6\" + 3'8\"");
assert.deepEqual(selection, {
  start: input.value.length,
  end: input.value.length
});

assert.equal(
  calculator.normalizeExpressionInput(input),
  false,
  "Already normalized input must not be changed again."
);

console.log("floating calculator input normalization tests passed");
