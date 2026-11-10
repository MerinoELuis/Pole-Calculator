"use strict";

const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");

const state = {
  settings: { projectProfile: "METRONET", position: "TOP_COMM" },
  poles: { P1: {} }
};
let hasProposedFiber = true;
const window = {
  CompactAutoProposed: {
    hasProposedFiber: () => hasProposedFiber,
    allowsOverlashExportWithoutProposedFiber: currentState =>
      String(currentState.settings.projectProfile).toUpperCase() === "OLSSON_OPPD"
      && String(currentState.settings.position).toUpperCase() === "OVERLASH",
    isExportProjectSupported: currentState => ["INTEC", "OLSSON_OPPD"].includes(String(currentState.settings.projectProfile).toUpperCase())
  }
};
const modulePath = path.join(__dirname, "..", "js", "ui", "pole-renderer.js");
const reviewRendererPath = path.join(__dirname, "..", "js", "ui", "excel-review-renderer.js");
vm.runInNewContext(fs.readFileSync(reviewRendererPath, "utf8"), { window }, { filename: reviewRendererPath });
vm.runInNewContext(fs.readFileSync(modulePath, "utf8"), { window }, { filename: modulePath });

const button = () => ({ disabled: false, textContent: "", title: "", classList: { toggle() {} } });
const autoCalculateBtn = button();
const exportProposedJsonBtn = button();
const renderer = window.PoleWorkspaceRenderer.create({
  S: { getState: () => state },
  els: { autoCalculateBtn, exportProposedJsonBtn },
  autoCalculateRunning: false
});

renderer.updateAutoCalculateButtonState();
assert.equal(exportProposedJsonBtn.disabled, true, "MetroNet must disable AutoProposed export until plugin support exists");
assert.match(exportProposedJsonBtn.title, /INTEC and Olsson OPPD only/i);

state.settings.projectProfile = "OLSSON_OPPD";
renderer.updateAutoCalculateButtonState();
assert.equal(exportProposedJsonBtn.disabled, false, "Olsson OPPD may export when pole data and proposed fiber are present");
assert.equal(exportProposedJsonBtn.title, "Export the proposed fiber data.");

state.settings.projectProfile = "INTEC";
renderer.updateAutoCalculateButtonState();
assert.equal(exportProposedJsonBtn.disabled, false, "INTEC may export when pole data and proposed fiber are present");
assert.equal(exportProposedJsonBtn.title, "Export the proposed fiber data.");

hasProposedFiber = false;
renderer.updateAutoCalculateButtonState();
assert.equal(exportProposedJsonBtn.disabled, true, "INTEC must still require a proposed fiber");

state.settings.projectProfile = "OLSSON_OPPD";
state.settings.position = "TOP_COMM";
renderer.updateAutoCalculateButtonState();
assert.equal(exportProposedJsonBtn.disabled, true, "regular Olsson OPPD export must still require a proposed fiber");
state.settings.position = "OVERLASH";
renderer.updateAutoCalculateButtonState();
assert.equal(exportProposedJsonBtn.disabled, false, "Olsson OPPD overlash can export without a new fiber proposal");
assert.equal(exportProposedJsonBtn.title, "Export Olsson OPPD overlash data without a newly proposed fiber.");

console.log("Proposed export project gating tests passed");
