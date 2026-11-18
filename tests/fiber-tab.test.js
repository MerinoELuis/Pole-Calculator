"use strict";

const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");

const window = {};
const modulePath = path.join(__dirname, "..", "js", "ui", "fiber-tab-renderer.js");
vm.runInNewContext(fs.readFileSync(modulePath, "utf8"), { window, console }, { filename: modulePath });
const FiberTabRenderer = window.FiberTabRenderer;

assert.deepEqual(
  Array.from(FiberTabRenderer.splitFiberComponents("6.6M 48CT Fiber (S/N) / MST (E/W) + 24CT Fiber (NW)")),
  ["6.6M 48CT Fiber (S/N)", "MST (E/W)", "24CT Fiber (NW)"]
);
assert.deepEqual(
  JSON.parse(JSON.stringify(FiberTabRenderer.parseFiberComponent("MST 8 Port (E/W)"))),
  { raw: "MST 8 Port (E/W)", fiberType: "MST", directions: ["E", "W"], mstPorts: "8" }
);

const baseState = {
  settings: {
    fiberSizes: { "24CT Fiber": "0.44", MST: "0.50", "48CT Fiber": "0.46" },
    fiberPlanDefaults: { mstPorts: "16" }
  },
  poles: { P1: {}, P2: {}, P3: {} },
  spans: {
    northFore: { spanId: "northFore", fromPole: "P1", toPole: "P2", bearingDegrees: 0, type: "Fore Span", spanIndex: 1 },
    northBack: { spanId: "northBack", fromPole: "P2", toPole: "P1", bearingDegrees: 180, type: "Back Span", spanIndex: 2 },
    east: { spanId: "east", fromPole: "P1", toPole: "P3", bearingDegrees: 90, type: "Fore Span", spanIndex: 3 }
  },
  makeReadyReferences: [
    {
      makeReadyId: "MR-1", makeReadyIndex: "1", poleId: "P1", collectionId: "COL-1",
      attachmentSizeRaw: "6.6M 24CT Fiber (N) / MST 8 Port (E)",
      attachmentType: "Overlash"
    },
    {
      makeReadyId: "MR-2", makeReadyIndex: "2", poleId: "P1", collectionId: "COL-1",
      attachmentSizeRaw: "6.6M 48CT Fiber (SW)", attachmentType: "New"
    }
  ],
  fiberAssignments: []
};

const initialRows = FiberTabRenderer.buildRows(baseState);
assert.equal(initialRows.length, 3, "fiber components are split into pole/span planning rows");
const initial24 = initialRows.find(row => row.fiberType === "24CT Fiber");
const initialMst = initialRows.find(row => row.fiberType === "MST");
const initial48 = initialRows.find(row => row.fiberType === "48CT Fiber");
assert.equal(initial24.spanId, "northFore", "reciprocal Fore/Back rows represent one physical span; Fore is preferred");
assert.equal(initial24.diameter, "0.44");
assert.equal(initial24.installType, "OVERLASH");
assert.equal(initialMst.mstPorts, "8", "an imported MST port count takes precedence over the bulk default");
assert.equal(initialMst.spanId, "east");
assert.equal(initialMst.installType, "OVERLASH");
assert.equal(initial48.spanId, "", "a direction without a matching span remains unassigned for manual selection");
assert.equal(initial48.installType, "NEW");

const outOfOrderPolesState = {
  settings: baseState.settings,
  poles: { P1: {}, P10: {}, P2: {}, P3: {} },
  spans: {},
  makeReadyReferences: ["P1", "P10", "P2", "P3"].map((poleId, index) => ({
    makeReadyId: `MR-${index + 1}`,
    poleId,
    attachmentSizeRaw: "24CT Fiber",
    attachmentType: "New"
  })),
  fiberAssignments: []
};
assert.deepEqual(
  Array.from(FiberTabRenderer.buildRows(outOfOrderPolesState), row => row.poleId),
  ["P1", "P2", "P3", "P10"],
  "fiber assignments use natural pole order even when imported pole keys put P10 before P2"
);

const editedState = {
  ...baseState,
  fiberAssignments: [{
    id: initial24.id,
    poleId: "P1",
    spanId: "east",
    diameter: "0.45",
    installType: "NEW"
  }]
};
const editedRows = FiberTabRenderer.buildRows(editedState);
const edited24 = editedRows.find(row => row.fiberType === "24CT Fiber");
assert.equal(edited24.id, initial24.id, "row identity remains stable when its selected span changes");
assert.equal(edited24.spanId, "east");
assert.equal(edited24.spanDirection, "E");
assert.equal(edited24.diameter, "0.44", "the global fiber diameter remains authoritative over legacy row overrides");
assert.equal(edited24.installType, "NEW");

const noImportedFiberState = {
  ...baseState,
  makeReadyReferences: [],
  fiberAssignments: [{
    id: "manual-fiber:1:1", manual: true, poleId: "P1", spanId: "east",
    fiberType: "MST", diameter: "0.50", mstPorts: "16", installType: "OVERLASH"
  }]
};
const manuallyAddedRows = FiberTabRenderer.buildRows(noImportedFiberState);
assert.equal(manuallyAddedRows.length, 1, "manual rows keep the Fiber view usable with no imported fibers");
assert.equal(manuallyAddedRows[0].sourceLabel, "Manual assignment");
assert.equal(manuallyAddedRows[0].spanDirection, "E");
assert.equal(manuallyAddedRows[0].mstPorts, "16");
assert.equal(manuallyAddedRows[0].installType, "OVERLASH");

const manualWithoutExcelType = FiberTabRenderer.buildRows({
  ...noImportedFiberState,
  fiberAssignments: [{ id: "manual-fiber:1:2", manual: true, poleId: "P1", spanId: "east", fiberType: "24CT Fiber" }]
});
assert.equal(manualWithoutExcelType[0].installType, "", "New/Overlash is not inherited from a global default for manual fibers");

function renderFiberView(state) {
  const container = { innerHTML: "", addEventListener() {} };
  FiberTabRenderer.create({
    els: { fiberPlanResults: container },
    S: { getState: () => state },
    recordUndoSnapshot() {},
    markDirty() {},
    toast() {}
  }).render();
  return container.innerHTML;
}

assert.doesNotMatch(renderFiberView({ ...noImportedFiberState, fiberAssignments: [] }), /data-fiber-default="mstPorts"/);
assert.match(renderFiberView(noImportedFiberState), /data-fiber-default="mstPorts"/, "global MST ports control appears when an MST row exists");
assert.doesNotMatch(renderFiberView(noImportedFiberState), /data-fiber-field="diameter"/, "individual spans do not override the global diameter");
const groupedRowsHtml = renderFiberView(baseState);
assert.match(groupedRowsHtml, /class="fiber-pole-cell" rowspan="3"><strong>P1<\/strong><\/td>/, "consecutive rows for one pole share a row-spanning pole cell");
assert.equal((groupedRowsHtml.match(/class="fiber-pole-cell"/g) || []).length, 1, "the repeated pole label is rendered once per pole group");
assert.match(groupedRowsHtml, /class="fiber-actions-cell"/, "Actions remains a per-fiber row so manual fibers can be removed individually");
assert.ok(groupedRowsHtml.indexOf("<th>INSTALL</th><th>DIRECTION</th><th>SPAN</th>") >= 0, "Direction follows Install and precedes Span");
const orderedPoleOptions = renderFiberView(outOfOrderPolesState).match(/<label class="fiber-default-field"><span>Pole<\/span><select class="input" data-fiber-new="poleId">([^]*?)<\/select><\/label>/)?.[1] || "";
assert.ok(
  orderedPoleOptions.indexOf('value="P1"') < orderedPoleOptions.indexOf('value="P2"')
    && orderedPoleOptions.indexOf('value="P2"') < orderedPoleOptions.indexOf('value="P3"')
    && orderedPoleOptions.indexOf('value="P3"') < orderedPoleOptions.indexOf('value="P10"'),
  "the add-fiber pole selector also uses natural numeric order"
);

console.log("fiber tab tests passed");
