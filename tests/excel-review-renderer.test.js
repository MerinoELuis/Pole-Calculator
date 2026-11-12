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
    expandExcelReviewIssuesBtn: { classList: classList(), dataset: {}, setAttribute() {} },
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

const activeProblemDetail = {
  dataset: { reviewResultPole: "P1" },
  open: false,
  addEventListener(name, handler) { if (name === "toggle") this.onToggle = handler; }
};
const passingDetail = {
  dataset: { reviewResultPole: "P2" },
  open: false,
  addEventListener() {}
};
const reviewState = {
  reviewedAt: "2026-11-11T12:00:00.000Z",
  summary: { errors: 1, warnings: 0, passed: 1 },
  globalChecks: [],
  results: [
    { poleId: "P1", sequence: "1", overallStatus: "ERROR", hoaStatus: "ERROR", finalStatus: "PASS", checks: [{ status: "ERROR", phase: "HOA", title: "Test" }] },
    { poleId: "P2", sequence: "2", overallStatus: "PASS", hoaStatus: "PASS", finalStatus: "PASS", checks: [{ status: "PASS", phase: "HOA", title: "Test" }] }
  ]
};
window.ExcelReview.getReviewState = () => reviewState;
const reviewUiState = { ui: { excelReviewExpandedPoleIds: [] }, excelReviewSource: { collection: { rows: [{}] } }, poles: {} };
const reviewDetails = [activeProblemDetail, passingDetail];
const populatedReviewResults = {
  innerHTML: "",
  querySelectorAll(selector) {
    if (selector.startsWith("details[open]")) return reviewDetails.filter(detail => detail.open);
    if (selector === "details[data-review-result-pole]") return reviewDetails;
    if (selector.includes("data-review-has-problem=\"true\"")) return [activeProblemDetail];
    return [];
  }
};
let dirtyCount = 0;
const interactiveRenderer = window.ExcelReviewRenderer.create({
  els: {
    excelReviewTabWarning: { classList: classList() },
    excelReviewTabWarningCount: { textContent: "" },
    excelReviewResults: populatedReviewResults,
    excelReviewTimestamp: { textContent: "" },
    expandExcelReviewIssuesBtn: { classList: classList(), dataset: {}, setAttribute() {} },
    ignoreHoaWarningsBtn: { classList: classList(), setAttribute() {} },
    ignoreFinalWarningsBtn: { classList: classList(), setAttribute() {} },
    rerunExcelReviewBtn: { classList: classList(), setAttribute() {} }
  },
  escapeHtml: value => String(value),
  markDirty() { dirtyCount += 1; },
  recordUndoSnapshot() {},
  S: { getState: () => reviewUiState, canonicalPoleIdentity: value => String(value).toUpperCase() },
  selectPole() {},
  toast() {}
});
interactiveRenderer.renderExcelReviewResults();
assert.match(populatedReviewResults.innerHTML, /data-review-has-problem="true"/);
assert.match(populatedReviewResults.innerHTML, /data-review-has-problem="false"/);
const expandResult = interactiveRenderer.toggleExcelReviewProblemPoles();
assert.equal(expandResult.count, 1, "the action affects only poles with active findings");
assert.equal(expandResult.expanded, true, "the first click expands the problem poles");
assert.equal(activeProblemDetail.open, true);
assert.equal(passingDetail.open, false, "passing poles remain collapsed");
assert.equal(Array.from(reviewUiState.ui.excelReviewExpandedPoleIds).join(","), "P1", "expanded pole IDs are saved with UI state");
assert.equal(dirtyCount, 1, "saving the expanded view marks the job dirty");
const collapseResult = interactiveRenderer.toggleExcelReviewProblemPoles();
assert.equal(collapseResult.expanded, false, "the same action collapses poles once all problem poles are open");
assert.equal(activeProblemDetail.open, false);
assert.equal(Array.from(reviewUiState.ui.excelReviewExpandedPoleIds).join(","), "", "collapsed problem poles are removed from saved open state");
assert.equal(dirtyCount, 2);
activeProblemDetail.open = false;
activeProblemDetail.onToggle();
assert.equal(Array.from(reviewUiState.ui.excelReviewExpandedPoleIds).join(","), "", "manually collapsing a pole updates the saved view state");

console.log("Excel Review renderer tests passed");
