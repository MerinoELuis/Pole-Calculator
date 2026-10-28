"use strict";

const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");

function renderTable(spanRecords) {
  const sides = Object.fromEntries(spanRecords.map(span => [span.spanId, {
    spanId: span.spanId,
    poleId: "P34",
    proposedHOA: `20'`,
    proposedHOAChange: `19'4\"`,
    endDrop: `-8\"`,
    ocalcMS: `20'`,
    msProposed: `19'`,
    finalMidspan: `19'`,
    clearanceMSStatus: "OK",
    proposedFlaggingStatus: "OK"
  }]));
  const store = {
    ENVIRONMENT_OPTIONS: [],
    getState: () => ({ settings: { proposeForeSpanWithoutMidspan: true } }),
    getPole: () => ({ poleId: "P34", maxCommHeight: `25'` }),
    getSpanSide: spanId => sides[spanId] || null,
    upsertSpanSide: side => (sides[side.spanId] ||= side),
    getSpan: spanId => spanRecords.find(span => span.spanId === spanId) || null,
    getConnectedSpans: () => spanRecords
  };
  const window = {
    Calculations: {
      isSpanEligibleForProposed: () => true,
      spanHasRealMidspan: () => true,
      evaluateProposedPoleClearance: () => ({ ok: true }),
      evaluateSpanSideFlagging: () => ({ status: "OK" }),
      isReciprocalBackSpan: () => false,
      isTerminalBackSpan: () => false
    }
  };
  const reviewRendererPath = path.join(__dirname, "..", "js", "ui", "excel-review-renderer.js");
  vm.runInNewContext(fs.readFileSync(reviewRendererPath, "utf8"), { window }, { filename: reviewRendererPath });
  const sourcePath = path.join(__dirname, "..", "js", "ui", "pole-renderer.js");
  vm.runInNewContext(fs.readFileSync(sourcePath, "utf8"), { window }, { filename: sourcePath });
  const escapeHtml = value => String(value ?? "").replace(/[&<>\"]/g, char => ({
    "&": "&amp;", "<": "&lt;", ">": "&gt;", "\"": "&quot;"
  })[char]);
  const renderer = window.PoleWorkspaceRenderer.create({
    S: store,
    H: { compareHeights: () => 0, parseHeight: () => null },
    escapeHtml,
    spanRelationOf: span => span?.type || "",
    SPAN_COLOR_CLASS_COUNT: 10
  });
  return renderer.renderSpanProposedTable("P34");
}

const backspanOnly = renderTable([{
  spanId: "BACK",
  fromPole: "P34",
  toPole: "P33",
  type: "Backspan",
  midspanMaxCommHeight: `25'9\"`
}]);
assert.match(backspanOnly, /backspan-proposed-only-table/);
assert.match(backspanOnly, /<th>Span<\/th><th>Proposed<\/th><th>Actions<\/th>/);
const backspanOnlyBody = backspanOnly.match(/<tbody>([\s\S]*?)<\/tbody>/)?.[1] || "";
const backspanOnlyRows = Array.from(backspanOnlyBody.matchAll(/<tr\b[^>]*>([\s\S]*?)<\/tr>/g));
assert.equal(backspanOnlyRows.length, 1, "a backspan-only pole without a standalone Proposed HOA should render one data row");
assert.equal(
  (backspanOnlyRows[0][1].match(/<td\b/g) || []).length,
  3,
  "backspan-only rows must contain only the Span, Proposed, and Actions cells"
);
assert.doesNotMatch(backspanOnly, /End Drop|Next Pole Proposed|O-CALC MS|MS Proposed|Max Height at MS|Adjusted Final MS|Environment Clearance/);
assert.match(backspanOnly, /data-field="proposedHOA"/);
assert.match(backspanOnly, /data-delete-proposed-span/);

const mixedSpans = renderTable([
  { spanId: "BACK", fromPole: "P34", toPole: "P33", type: "Backspan", midspanMaxCommHeight: `25'9\"` },
  { spanId: "FORE", fromPole: "P34", toPole: "P35", type: "Forespan", midspanMaxCommHeight: `25'9\"` }
]);
assert.doesNotMatch(mixedSpans, /backspan-proposed-only-table/);
assert.match(mixedSpans, /<th>End Drop<\/th>/);
const backspanRow = mixedSpans.match(/<tr class="[^"]*">(?=[\s\S]*?data-span="BACK")[\s\S]*?<\/tr>/)?.[0] || "";
assert.match(backspanRow, /data-field="proposedHOA"/);
assert.match(backspanRow, /data-field="proposedHOAChange"/);
assert.match(backspanRow, /data-field="ocalcMS"/);
assert.match(mixedSpans, /data-span="FORE" data-field="ocalcMS"/);

console.log("Backspan Proposed table display tests passed");
