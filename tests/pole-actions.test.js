"use strict";

const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");

const modulePath = path.join(__dirname, "..", "js", "ui", "controls.js");
const window = {
  Calculations: { recalculateSpansForPole() {}, recalculateSpan() {} },
  MRLogic: { generateMRForUGChange: () => ["P1", "P2"] }
};
vm.runInNewContext(fs.readFileSync(modulePath, "utf8"), { window }, { filename: modulePath });

let pole = { poleId: "P1", ugActive: false, pcoActive: false };
let undoCount = 0;
let recalculationCount = 0;
const listeners = { toggle: {}, match: {} };
const spanUGListeners = {};
const toggleUGAllListeners = {};
const span = { spanId: "S1", fromPole: "P1", toPole: "P2", ugActive: false, ugReason: "" };
const button = {
  dataset: { pole: "P1" },
  addEventListener(eventName, callback) { listeners.toggle[eventName] = callback; }
};
const matchButton = {
  dataset: { pole: "P1" },
  addEventListener(eventName, callback) { listeners.match[eventName] = callback; }
};
const spanUGCheckbox = {
  dataset: { pole: "P1", span: "S1" },
  checked: true,
  addEventListener(eventName, callback) { spanUGListeners[eventName] = callback; }
};
const toggleUGAllButton = {
  dataset: { pole: "P1" },
  addEventListener(eventName, callback) { toggleUGAllListeners[eventName] = callback; }
};
let matchPole = "";
let dirtyCount = 0;
let renderedPoleIds = [];
let spanMRRefreshCount = 0;
window.MRLogic.generateMRForUGChange = () => {
  spanMRRefreshCount += 1;
  return ["P1", "P2"];
};
const context = {
  S: {
    getPole: () => pole,
    getSpan: () => span,
    getConnectedSpans: () => [span],
    getOtherPoleId: () => "P2",
    updateSpanField(spanId, field, value) { span[field] = value; },
    upsertPole(nextPole) { pole = nextPole; }
  },
  poleResolutionOf: selectedPole => selectedPole?.ugActive ? "UG" : "AERIAL",
  editCommPairing(poleId) { matchPole = poleId; },
  recordUndoSnapshot() { undoCount += 1; },
  renderAffectedPoles(poleIds) { renderedPoleIds = poleIds; },
  markDirty() { dirtyCount += 1; },
  toast() {}
};
window.Calculations.recalculateSpansForPole = () => { recalculationCount += 1; };
window.Calculations.recalculateSpan = () => { recalculationCount += 1; };
const actions = window.PoleActions.create(context);
actions.bindLocalActions({
  querySelectorAll(selector) {
    if (selector === "[data-toggle-ug]") return [button];
    if (selector === "[data-toggle-span-ug]") return [spanUGCheckbox];
    if (selector === "[data-toggle-ug-all]") return [toggleUGAllButton];
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
spanUGListeners.change();
assert.equal(span.ugActive, true, "the span checkbox should independently activate a UG transfer");
assert.equal(pole.ugActive, false, "selecting a specific span should keep the pole itself aerial");
assert.equal(span.ugReason, "CLEARANCE VIOLATION", "a new span-level UG transfer should receive the default reason");
assert.equal(spanMRRefreshCount, 2, "UG scope edits should regenerate connected Make Ready");
assert.deepEqual(Array.from(renderedPoleIds), ["P1", "P2"], "the UG pole and its adjacent pole should refresh together");
assert.equal(dirtyCount, 2, "UG scope edits should mark the job dirty");
toggleUGAllListeners.click();
assert.equal(pole.ugActive, true, "ALL should switch a span-only selection back to full-pole UG");
assert.equal(pole.ugSpanId, "ALL", "the whole-pole mode should preserve the ALL scope");
assert.equal(span.ugActive, false, "switching to ALL should clear span-only UG selections");
assert.equal(spanMRRefreshCount, 3, "switching to ALL should refresh connected Make Ready");
assert.equal(dirtyCount, 3, "switching to ALL should mark the job dirty");

console.log("pole action binding tests passed");
