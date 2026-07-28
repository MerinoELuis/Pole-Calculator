"use strict";

const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");

const window = {
  HeightUtils: {
    formatHeight(value) {
      const feet = Math.floor(Number(value) / 12);
      const inches = Number(value) % 12;
      return `${feet}'${inches ? `${inches}\"` : ""}`;
    }
  }
};

const sourcePath = path.join(__dirname, "..", "js", "ui", "auto-calculate-status.js");
vm.runInNewContext(fs.readFileSync(sourcePath, "utf8"), { window, JSON, String, Number, Object }, { filename: sourcePath });
const ui = window.AutoCalculateStatusUI;

assert.equal(ui.statusLabel("BEST_AVAILABLE"), "BEST AVAILABLE");
assert.equal(ui.statusClass("CRITICAL"), "critical");
assert.equal(ui.modeLabel("LOW_COMM"), "LOW COMM");
assert.equal(ui.escapeHtml("<test>"), "&lt;test&gt;");

const result = {
  status: "BEST_AVAILABLE",
  mode: "TOP_COMM",
  message: "Pole clearances satisfied.",
  poleViolationCount: 0,
  midspanViolationCount: 1,
  movedCommCount: 2,
  totalMovementInches: 14,
  idealProposed: "21'8\"",
  recommendation: "Review UG or PCO."
};
const markup = ui.resultMarkup(result);
assert.match(markup, /BEST AVAILABLE/);
assert.match(markup, /TOP COMM/);
assert.match(markup, /Midspan issues/);
assert.match(markup, /Review UG or PCO/);
assert.equal(ui.resultSignature(result), ui.resultSignature({ ...result }));
assert.notEqual(ui.resultSignature(result), ui.resultSignature({ ...result, movedCommCount: 3 }));

console.log("Auto Calculate status UI tests passed");
