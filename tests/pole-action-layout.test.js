"use strict";

const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");

const pole = {
  poleId: "P1",
  poleInsetActive: true,
  riserActive: true,
  ancActive: true,
  ancDistance: "15'",
  ancDirection: "N",
  dgHoa: "19'",
  metadata: {
    powerEquipment: [],
    proposedANCs: [{ ancDistance: "15'", ancDirection: "E", dgHoa: "18'", manualFields: [] }]
  }
};
const state = { settings: { projectProfile: "INTEC" }, mr: [] };
const window = {
  ExcelReviewRenderer: { create: () => ({}) },
  PowerEquipmentRenderer: { create: () => ({ renderPowerTable() {}, renderCollapseEyeIcon() { return ""; }, renderEquipmentTable() {} }) },
  AnchorRenderer: { create: () => ({ renderAnchorSection() { return ""; } }) },
  PoleSettingsRenderer: { create: () => ({}) },
  MRLogic: {
    isANCEnabled: () => true,
    isRiserAvailable: () => true,
    isRiserEnabled: () => true,
    getUGSpanOptions: () => [
      { spanId: "SPAN-W", relation: "Backspan", direction: "W", otherPoleId: "P2", label: "Backspan W · P2" },
      { spanId: "SPAN-E", relation: "Forespan", direction: "E", otherPoleId: "P3", label: "Forespan E · P3" }
    ],
    getEffectiveANCDistance: () => "15'",
    getEffectiveANCDGHOA: () => "19'",
    getEffectiveANCDirection: () => "N",
    getResolvedRiserDirection: () => "S",
    getImportedRiserDirection: () => ""
  },
  Calculations: {
    getPoleInsetDetails: () => ({
      spanId: "S1",
      selectedHeight: "",
      maxHeightAtMidspan: "17'1\"",
      endDropToPole: "—",
      endDropToNextPole: "—",
      spanOptions: [{ spanId: "S1", label: "P1 → P2" }]
    })
  }
};
const sandbox = { window, console, Set, Map, String, Number, Boolean, Array, Object, Math, RegExp };
const rendererPath = path.join(__dirname, "..", "js", "ui", "pole-renderer.js");
vm.runInNewContext(fs.readFileSync(rendererPath, "utf8"), sandbox, { filename: rendererPath });
const escapeHtml = value => String(value ?? "").replace(/[&<>\"]/g, char => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", "\"": "&quot;" })[char]);
const renderer = window.PoleWorkspaceRenderer.create({
  S: {
    getPole: () => pole,
    getState: () => state,
    getSpan: spanId => ({ "SPAN-W": { ugActive: true }, "SPAN-E": { ugActive: false } })[spanId] || null
  },
  H: { parseHeight: () => null, formatHeight: value => String(value) },
  escapeHtml,
  ownerCanonicalOf: () => "",
  ownerImportedOf: () => "",
  poleResolutionOf: () => "AERIAL",
  SPAN_COLOR_CLASS_COUNT: 8,
  spanRelationOf: () => "",
  els: {}
});
const markup = renderer.renderPoleActions("P1");

assert.match(markup, /data-field="dgHoa"[^]*?<\/label>\s*<div class="proposed-anc-row-controls"><button[^>]*data-add-proposed-anc/,
  "the primary ANC plus should occupy the same action-grid row as DG HOA");
assert.match(markup, /data-field="dgHoa"[^]*?<\/label>\s*<div class="proposed-anc-row-controls"><button[^>]*data-add-proposed-anc[^]*?<button[^>]*data-remove-proposed-anc/,
  "an additional ANC should show plus and minus beside its DG HOA row");
assert.match(markup, /class="pole-action-section pole-inset-section"/, "Pole Inset should have a separated action block");
assert.match(markup, /class="pole-action-section"><label class="pole-action-field riser-direction-field"/, "Riser should have a separated action block");
assert.match(markup, /class="pole-action-section proposed-anc-section"/, "ANC should have a separated action block");
assert.match(markup, /class="pole-action-section span-ug-transfer-section"[^]*?data-span="SPAN-W"[^]*?checked[^]*?data-span="SPAN-E"/,
  "an aerial pole should expose independent UG controls for each connected physical span");
assert.match(markup, /The pole stays aerial; ANC and Riser remain available\./,
  "span-level UG should clearly state that normal aerial-pole actions remain available");
assert.ok((markup.match(/class="pole-action-section/g) || []).length >= 3,
  "simultaneously active Pole Inset, Riser, and ANC actions should each be separated");

const styles = fs.readFileSync(path.join(__dirname, "..", "css", "styles.css"), "utf8");
assert.match(styles, /\.pole-action-section\s*\{[^}]*border-top:\s*1px solid/s, "each action divider should use the shared subtle line style");
assert.match(styles, /\.proposed-anc-row-controls\s*\{[^}]*align-self:\s*end/s, "ANC controls should align to the bottom of the DG HOA row");

console.log("Pole action divider and ANC button layout tests passed.");
