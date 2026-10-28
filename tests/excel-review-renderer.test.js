"use strict";

const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");

function classList() {
  return { toggle() {} };
}

const reviewResults = {
  innerHTML: "",
  querySelectorAll: () => []
};
const window = {
  ExcelReview: {
    getReviewState: () => ({
      reviewedAt: "",
      summary: { errors: 0, warnings: 0, passed: 0 },
      globalChecks: [],
      results: []
    }),
    warningControlState: () => ({ count: 0, allIgnored: false })
  }
};
const modulePath = path.join(__dirname, "..", "js", "ui", "excel-review-renderer.js");
vm.runInNewContext(fs.readFileSync(modulePath, "utf8"), { window }, { filename: modulePath });

const renderer = window.ExcelReviewRenderer.create({
  els: {
    excelReviewTabWarning: { classList: classList() },
    excelReviewTabWarningCount: { textContent: "" },
    excelReviewResults: reviewResults,
    excelReviewTimestamp: { textContent: "" },
    ignoreHoaWarningsBtn: { classList: classList(), setAttribute() {} },
    ignoreFinalWarningsBtn: { classList: classList(), setAttribute() {} },
    rerunExcelReviewBtn: { classList: classList() }
  },
  escapeHtml: value => String(value).replaceAll("&", "&amp;").replaceAll("<", "&lt;").replaceAll(">", "&gt;"),
  markDirty() {},
  recordUndoSnapshot() {},
  S: {
    getState: () => ({ excelReviewSource: { collection: { rows: [] } }, poles: {} }),
    canonicalPoleIdentity: value => String(value).toUpperCase()
  },
  selectPole() {},
  toast() {}
});

assert.match(renderer.reviewStatusBadge("NOT_READY"), /badge warning.*NOT READY/);
assert.match(renderer.reviewStatusBadge("NOT_APPLICABLE"), /badge ignored.*N\/A/);
assert.match(renderer.renderReviewCheck({ status: "ERROR", title: "<script>", message: "<unsafe>" }), /&lt;script&gt;/);
assert.doesNotMatch(renderer.renderReviewCheck({ status: "ERROR", title: "<script>", message: "<unsafe>" }), /<script>/);

renderer.renderExcelReviewResults();
assert.match(reviewResults.innerHTML, /Import a raw Excel file/);

console.log("Excel Review renderer tests passed");
