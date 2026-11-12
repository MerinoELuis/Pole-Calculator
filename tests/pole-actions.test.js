"use strict";

const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");

const modulePath = path.join(__dirname, "..", "js", "ui", "controls.js");
const window = { Calculations: { recalculateSpansForPole() {} } };
vm.runInNewContext(fs.readFileSync(modulePath, "utf8"), { window }, { filename: modulePath });

let pole = { poleId: "P1", ugActive: false, pcoActive: false };
let undoCount = 0;
let recalculationCount = 0;
const listeners = { toggle: {}, match: {} };
const button = {
  dataset: { pole: "P1" },
  addEventListener(eventName, callback) { listeners.toggle[eventName] = callback; }
};
const matchButton = {
  dataset: { pole: "P1" },
  addEventListener(eventName, callback) { listeners.match[eventName] = callback; }
};
let matchPole = "";
const context = {
  S: {
    getPole: () => pole,
    getConnectedSpans: () => [],
    upsertPole(nextPole) { pole = nextPole; }
  },
  poleResolutionOf: () => "AERIAL",
  editCommPairing(poleId) { matchPole = poleId; },
  recordUndoSnapshot() { undoCount += 1; },
  renderAffectedPoles() {},
  toast() {}
};
window.Calculations.recalculateSpansForPole = () => { recalculationCount += 1; };
const actions = window.PoleActions.create(context);
actions.bindLocalActions({
  querySelectorAll(selector) {
    if (selector === "[data-toggle-ug]") return [button];
    if (selector === "[data-match-pole]") return [matchButton];
    return [];
  }
});

assert.equal(typeof listeners.toggle.click, "function", "pole actions should bind the generated UG control");
listeners.toggle.click();
assert.equal(pole.ugActive, true, "the extracted action should activate UG on the selected pole");
assert.equal(undoCount, 1, "the extracted action should keep undo history behavior");
assert.equal(recalculationCount, 1, "the extracted action should recalculate after changing state");
assert.equal(typeof listeners.match.click, "function", "the pole-header Match icon should be bound");
listeners.match.click();
assert.equal(matchPole, "P1", "the Match icon should open cable matching for its pole");

console.log("pole action binding tests passed");
