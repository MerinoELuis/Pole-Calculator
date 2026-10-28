"use strict";

const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");

const rendererPath = path.join(__dirname, "..", "js", "ui", "pole-renderer.js");
const source = fs.readFileSync(rendererPath, "utf8");
const functionStart = source.indexOf("function renderCommMovementTable(poleId)");
const nextFunction = source.indexOf("\n      function ", functionStart + 1);
assert.notEqual(functionStart, -1, "communication movement table renderer should exist");
assert.notEqual(nextFunction, -1, "communication movement table renderer should end before the next function");

const renderer = source.slice(functionStart, nextFunction);
const header = renderer.slice(renderer.indexOf("<thead><tr>"), renderer.indexOf("</tr></thead>"));
const body = renderer.slice(renderer.indexOf("<tbody>"), renderer.indexOf("</tbody>"));

function assertInOrder(text, markers, context) {
  const positions = markers.map(marker => text.indexOf(marker));
  assert.ok(positions.every(position => position >= 0), `${context}: all expected columns should be rendered`);
  assert.deepEqual(
    positions,
    [...positions].sort((a, b) => a - b),
    `${context}: columns should appear in the requested order`
  );
}

assertInOrder(header, [
  "Service Drop",
  "Re-sag Service Drop",
  "Transfer to New Pole",
  "<th>DG</th>",
  "<th>Other HOA</th>",
  "<th>POF</th>",
  "<th>Actions</th>"
], "communication movement table header");

assertInOrder(body, [
  "renderCommServiceDropValues(group, poleId)",
  "renderCommResagValues(group, poleId)",
  "renderCommTransferValues(group, poleId)",
  "renderCommDownGuyValues(group, poleId)",
  "renderCommOtherHoaValues(group, poleId)",
  "group.pofEligible",
  "<div class=\"row-actions\">"
], "communication movement table cells");

console.log("Communication movement column order tests passed");
