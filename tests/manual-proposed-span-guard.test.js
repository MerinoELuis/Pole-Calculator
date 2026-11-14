"use strict";

const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");

const window = {
  localStorage: { getItem: () => null, setItem: () => {} },
  ExcelReviewRenderer: { create: () => ({}) },
  Calculations: { recalculateSpan: () => {} }
};
const sandbox = {
  window,
  console,
  Date,
  Set,
  Map,
  Math,
  JSON,
  Object,
  Array,
  String,
  Number,
  Boolean,
  RegExp,
  CSS: { escape: value => String(value) }
};

for (const file of [
  "domain/height-math.js",
  "core/project-config.js",
  "domain/pole-resolution.js",
  "core/state-compat.js",
  "core/store-identity.js",
  "core/store.js",
  "ui/poles/power-equipment-renderer.js",
  "ui/settings-renderer.js",
  "ui/pole-renderer.js",
  "ui/controls.js"
]) {
  const modulePath = path.join(__dirname, "..", "js", file);
  vm.runInNewContext(fs.readFileSync(modulePath, "utf8"), sandbox, { filename: modulePath });
}

const S = window.AppStore;
S.resetState();
for (const poleId of ["P1", "P2", "P3"]) S.upsertPole(S.createPole({ poleId }));
S.upsertSpan(S.createSpan("S12", "P1", "P2", "E", "", { type: "Fore Span" }));
S.upsertSpan(S.createSpan("S31", "P3", "P1", "W", "", { type: "Back Span" }));
S.upsertSpanSide(S.createSpanSide({
  spanId: "S12",
  poleId: "P1",
  proposedHOA: "22'",
  isManualProposed: true
}));
S.upsertSpanSide(S.createSpanSide({
  spanId: "S31",
  poleId: "P1",
  proposedHOAChange: "21'"
}));

assert.equal(S.hasProposedSpanForPolePair("P1", "P2"), true,
  "the source pole should recognize its existing proposal to the same target");
assert.equal(S.hasProposedSpanForPolePair("P2", "P1"), false,
  "the other endpoint keeps an independent local proposal slot");
assert.equal(S.hasProposedSpanForPolePair("P1", "P1"), false,
  "a pole cannot be a proposed target for itself");
assert.equal(S.hasProposedSpanForPolePair("P1", "P3"), false,
  "Next Pole Proposed alone must not count as a proposal at the current pole");

const renderer = window.PoleWorkspaceRenderer.create({ S, escapeHtml: value => String(value) });
const starter = renderer.renderManualProposedSpanStarter("P1", { compact: true });
const p2Option = starter.match(/<option value="P2"[^>]*>/)?.[0] || "";
const p3Option = starter.match(/<option value="P3"[^>]*>/)?.[0] || "";
assert.match(p2Option, /disabled/, "a target with an existing proposal should be disabled in the selector");
assert.doesNotMatch(p3Option, /disabled/, "a target without a proposal should remain selectable");

let undoCount = 0;
const messages = [];
const select = { value: "P2" };
const button = {
  dataset: { pole: "P1" },
  listeners: {},
  addEventListener(type, callback) { this.listeners[type] = callback; }
};
const root = {
  querySelector: () => select,
  querySelectorAll: selector => selector === "[data-add-proposed-span]" ? [button] : []
};
const noop = () => {};
const controls = window.PoleActions.create({
  S,
  H: window.HeightUtils,
  addCommToPole: noop,
  clearCommMidspan: noop,
  confirmInApp: async () => false,
  connectedSpansSorted: poleId => S.getConnectedSpans(poleId),
  deleteCommGroup: noop,
  deleteCommSpan: noop,
  deletePole: noop,
  deletePowerEquipment: noop,
  deleteSpanPower: noop,
  editCommSpans: noop,
  formatExistingViolationsText: () => "",
  groupedCommsForPole: () => [],
  hidePole: noop,
  markDirty: noop,
  poleIdsForSpan: () => [],
  poleResolutionOf: () => "AERIAL",
  recordUndoSnapshot: () => { undoCount += 1; },
  renderAffectedPoles: noop,
  shortSpanLabel: span => span?.spanId || "",
  toast: (message, type) => messages.push({ message, type }),
  updateCommGroupField: noop
});
controls.bindLocalActions(root);

const spanCountBeforeDuplicate = S.getConnectedSpans("P1").length;
button.listeners.click();
assert.equal(S.getConnectedSpans("P1").length, spanCountBeforeDuplicate,
  "attempting to add a second proposal to the same target must not create another span");
assert.equal(undoCount, 0, "a rejected duplicate must not create an undo snapshot");
assert.match(messages.at(-1)?.message || "", /already exists/i,
  "the duplicate attempt should explain that the existing proposal can be edited");

select.value = "P3";
button.listeners.click();
assert.equal(S.getConnectedSpans("P1").length, spanCountBeforeDuplicate,
  "a new proposal should reuse a physical span even when its endpoints are stored in reverse order");
assert.equal(S.getSpanSide("S31", "P1")?.isManualProposed, true,
  "the new proposal should be stored on the current pole side of the existing physical span");
assert.equal(S.getSpanSide("S31", "P1")?.proposedHOAChange, "21'",
  "adding the local proposal must preserve the independent Next Pole Proposed value");
assert.equal(undoCount, 1, "a valid proposal should create an undo snapshot");

select.value = "P1";
button.listeners.click();
assert.equal(S.getConnectedSpans("P1").length, spanCountBeforeDuplicate,
  "attempting a self-span must not add a physical span");
assert.match(messages.at(-1)?.message || "", /cannot connect a pole to itself/i,
  "the self-span attempt should explain why it was rejected");

console.log("Manual Proposed span uniqueness tests passed.");
