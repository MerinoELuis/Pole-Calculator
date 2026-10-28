"use strict";

const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");

const state = {
  settings: { projectProfile: "METRONET", position: "TOP_COMM" },
  poles: { P1: {} }
};
const window = {
  CompactAutoProposed: {
    hasProposedFiber: () => true,
    isExportProjectSupported: currentState => String(currentState.settings.projectProfile).toUpperCase() === "INTEC"
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
assert.match(exportProposedJsonBtn.title, /INTEC only/i);

state.settings.projectProfile = "INTEC";
renderer.updateAutoCalculateButtonState();
assert.equal(exportProposedJsonBtn.disabled, false, "INTEC may export when pole data and proposed fiber are present");
assert.equal(exportProposedJsonBtn.title, "Export the proposed fiber data.");

console.log("Proposed export project gating tests passed");
