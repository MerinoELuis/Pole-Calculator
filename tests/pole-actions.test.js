"use strict";

const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");

const modulePath = path.join(__dirname, "..", "js", "ui", "controls.js");
const window = {
  Calculations: { recalculateSpansForPole() {}, recalculateSpan() {} },
  MRLogic: {
    generateMRForUGChange: () => ["P1", "P2"],
    getUGSpanMRTemplate: (_poleId, spanIds) => Array.from(spanIds)
      .map(spanId => `${spanId} to go UG due to [insert reason].`)
      .join("\n")
  }
};
vm.runInNewContext(fs.readFileSync(modulePath, "utf8"), { window }, { filename: modulePath });

let pole = { poleId: "P1", ugActive: false, pcoActive: false };
let undoCount = 0;
let recalculationCount = 0;
const listeners = { toggle: {}, match: {}, add: {}, remove: {}, cancel: {} };
const ugScopeListeners = [];
const span = { spanId: "S1", fromPole: "P1", toPole: "P2", ugActive: false, ugReason: "", ugSourcePoleId: "" };
const span2 = { spanId: "S2", fromPole: "P1", toPole: "P3", ugActive: false, ugReason: "", ugSourcePoleId: "" };
const button = {
  dataset: { pole: "P1" },
  addEventListener(eventName, callback) { listeners.toggle[eventName] = callback; }
};
const matchButton = {
  dataset: { pole: "P1" },
  addEventListener(eventName, callback) { listeners.match[eventName] = callback; }
};
const firstUGScope = {
  dataset: { pole: "P1", currentSpan: "ALL" },
  value: "S1",
  addEventListener(eventName, callback) { ugScopeListeners[0] = callback; }
};
const addUGScope = {
  dataset: { pole: "P1", currentSpan: "" },
  value: "S2",
  addEventListener(eventName, callback) { ugScopeListeners[1] = callback; }
};
const allUGScope = {
  dataset: { pole: "P1", currentSpan: "S2" },
  value: "ALL",
  addEventListener(eventName, callback) { ugScopeListeners[2] = callback; }
};
let focusedAddSelect = false;
const addUGRow = {
  hidden: true,
  querySelector() { return addUGScope; }
};
const ugSection = { querySelector() { return addUGRow; } };
const addUGButton = {
  closest() { return ugSection; },
  addEventListener(eventName, callback) { listeners.add[eventName] = callback; }
};
const removeUGButton = {
  dataset: { pole: "P1", span: "S1" },
  addEventListener(eventName, callback) { listeners.remove[eventName] = callback; }
};
const cancelUGButton = {
  closest() { return addUGRow; },
  addEventListener(eventName, callback) { listeners.cancel[eventName] = callback; }
};
addUGScope.focus = () => { focusedAddSelect = true; };
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
    getConnectedSpans: () => [span, span2],
    getSpan(spanId) { return ({ S1: span, S2: span2 })[spanId] || null; },
    getOtherPoleId: () => "P2",
    updateSpanField(spanId, field, value) { ({ S1: span, S2: span2 })[spanId][field] = value; },
    updatePoleField(_poleId, field, value) { pole[field] = value; },
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
    if (selector === "[data-ug-span-select]") return [firstUGScope, addUGScope, allUGScope];
    if (selector === "[data-add-ug-span]") return [addUGButton];
    if (selector === "[data-remove-ug-span]") return [removeUGButton];
    if (selector === "[data-cancel-ug-span]") return [cancelUGButton];
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
ugScopeListeners[0]();
assert.equal(span.ugActive, true, "choosing a connected dropdown option should activate only that UG span");
assert.equal(pole.ugActive, false, "selecting a specific span should keep the pole itself aerial");
assert.equal(span.ugSourcePoleId, "P1", "the pole that selected a span owns the shared editable Make Ready box");
assert.match(pole.ugSpanMRText, /S1 to go UG due to \[insert reason\]\./, "the first selected span should appear in the shared Make Ready box");
assert.equal(spanMRRefreshCount, 2, "UG scope edits should regenerate connected Make Ready");
assert.deepEqual(Array.from(renderedPoleIds), ["P1", "P2"], "the UG pole and its adjacent pole should refresh together");
assert.equal(dirtyCount, 2, "UG scope edits should mark the job dirty");
listeners.add.click();
assert.equal(addUGRow.hidden, false, "the plus icon should reveal a selector for another UG span");
assert.equal(focusedAddSelect, true, "adding another span should focus its selector");
ugScopeListeners[1]();
assert.equal(span2.ugActive, true, "an additional dropdown should allow a second physical span to go UG");
assert.equal(pole.ugActive, false, "multiple span selections should leave the pole aerial");
assert.match(pole.ugSpanMRText, /S1 to go UG due to \[insert reason\]\.[\s\S]*S2 to go UG due to \[insert reason\]\./, "both selected spans should share one Make Ready box");
listeners.remove.click();
assert.equal(span.ugActive, false, "the minus icon should remove its selected physical span");
assert.equal(span2.ugActive, true, "removing one span should leave other UG span selections active");
assert.doesNotMatch(pole.ugSpanMRText, /S1 to go UG/, "removing a span should remove only its line from the shared Make Ready box");
ugScopeListeners[2]();
assert.equal(pole.ugActive, true, "choosing ALL should switch span selections back to full-pole UG");
assert.equal(pole.ugSpanId, "ALL", "the whole-pole mode should preserve the ALL scope");
assert.equal(span.ugActive, false, "switching to ALL should clear the first span-only UG selection");
assert.equal(span2.ugActive, false, "switching to ALL should clear every span-only UG selection");
assert.equal(spanMRRefreshCount, 5, "switching to ALL should refresh connected Make Ready");
assert.equal(dirtyCount, 5, "switching to ALL should mark the job dirty");

console.log("pole action binding tests passed");
