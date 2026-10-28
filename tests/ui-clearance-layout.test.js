"use strict";

const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");

const state = {
  settings: {
    projectProfile: "METRONET",
    proposedOwner: "MidAm",
    metronetWI: "MIDAM"
  },
  makeReadyReferences: []
};
const window = {
  ProjectProfiles: {
    PROFILES: { METRONET: { id: "METRONET", label: "Metronet" } },
    getProfileSettings: () => ({ proposedOwners: ["MidAm"] })
  }
};
const modulePath = path.join(__dirname, "..", "js", "ui", "pole-renderer.js");
const reviewRendererPath = path.join(__dirname, "..", "js", "ui", "excel-review-renderer.js");
vm.runInNewContext(fs.readFileSync(reviewRendererPath, "utf8"), { window }, { filename: reviewRendererPath });
vm.runInNewContext(fs.readFileSync(modulePath, "utf8"), { window }, { filename: modulePath });

const clearanceSettings = { innerHTML: "" };
const renderer = window.PoleWorkspaceRenderer.create({
  els: { clearanceSettings },
  S: { getState: () => state },
  escapeHtml: value => String(value ?? ""),
  wireEditableEvents: () => {}
});
renderer.renderClearanceSettings();

const groups = [...clearanceSettings.innerHTML.matchAll(/<div class="settings-grid clearance-grid">([\s\S]*?)<\/div>/g)]
  .map(match => match[1]);
assert.equal(groups.length, 3, "Metronet MidAm clearances should be split into pole, midspan, and streetlight rows");
assert.match(groups[0], /data-field="polePowerCommsClearance"/);
assert.match(groups[0], /data-field="primaryPowerCommsClearance"/);
assert.match(groups[0], /data-field="boltClearance"/);
assert.doesNotMatch(groups[0], /data-field="midspan|streetlight/);
assert.match(groups[1], /data-field="midspanPowerCommClearance"/);
assert.match(groups[1], /data-field="midspanPrimaryPowerCommClearance"/);
assert.match(groups[1], /data-field="midspanCommCommClearance"/);
assert.doesNotMatch(groups[1], /data-field="pole|streetlight/);
assert.ok(
  groups[2].indexOf('data-field="streetlightBracketCommClearance"')
    < groups[2].indexOf('data-field="streetlightDripLoopCommClearance"'),
  "Streetlight bracket clearance must appear before drip-loop clearance"
);
assert.match(groups[2], /Streetlight bracket-comm/);
assert.match(groups[2], /Streetlight drip loop-comm/);
assert.doesNotMatch(groups[2], /Pole · Streetlight/);

state.settings.projectProfile = "INTEC";
renderer.renderClearanceSettings();
const intecGroups = [...clearanceSettings.innerHTML.matchAll(/<div class="settings-grid clearance-grid">([\s\S]*?)<\/div>/g)]
  .map(match => match[1]);
assert.match(intecGroups[2], /data-field="streetlightBracketCommClearance"/);
assert.match(intecGroups[2], /data-field="streetlightGroundedBracketCommClearance"/);
assert.match(intecGroups[2], /Streetlight bracket-comm \(Grounding\)/);
assert.doesNotMatch(intecGroups[2], /Streetlight drip loop-comm/, "INTEC drip-loop clearance remains the standard Power clearance");

console.log("Clearance settings layout tests passed");
