"use strict";

const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");

const appPath = path.join(__dirname, "..", "js", "ui", "app.js");
const rendererPath = path.join(__dirname, "..", "js", "ui", "pole-renderer.js");
const excelReviewRendererPath = path.join(__dirname, "..", "js", "ui", "excel-review-renderer.js");
const appSource = fs.readFileSync(appPath, "utf8");
const rendererSource = fs.readFileSync(rendererPath, "utf8");
const excelReviewRendererSource = fs.readFileSync(excelReviewRendererPath, "utf8");

assert.doesNotMatch(appSource, /setAttribute\(["']pattern["']/i, "height fields should not set an HTML pattern rejected by modern browsers");

const reviewMarkup = excelReviewRendererSource.match(/return `<div class="excel-review-pole-entry">([\s\S]*?)`;/)?.[1];
assert.ok(reviewMarkup, "Excel review pole rows should keep the pole-opening action outside the disclosure summary");
const summaryMarkup = reviewMarkup.match(/<summary>([\s\S]*?)<\/summary>/)?.[1];
assert.ok(summaryMarkup, "Excel review pole rows should retain their disclosure summary");
assert.doesNotMatch(summaryMarkup, /<(?:button|input|select|textarea|a)\b/i, "summary content must not contain interactive descendants");
assert.match(reviewMarkup, /<button\b[^>]*data-review-pole=/i, "the pole-opening action should remain available as a sibling control");

const poleHeader = rendererSource.match(/function renderPoleEditableHeader\(poleId\) \{([\s\S]*?)\n      function renderSpanProposedTable/ )?.[1];
assert.ok(poleHeader, "pole header controls should be rendered in their dedicated function");
assert.match(poleHeader, /data-hide-pole[^>]*aria-label="Hide \$\{escapeHtml\(poleId\)\}"[^>]*><svg/i, "Hide should use an accessible icon button");
assert.match(poleHeader, /data-delete-pole[^>]*aria-label="Delete \$\{escapeHtml\(poleId\)\}"[^>]*><svg/i, "Delete should use an accessible icon button");
assert.doesNotMatch(poleHeader, />Hide<|>Delete</, "pole header actions should no longer show text labels");
assert.match(rendererSource, /Original Midspan before movements: \$\{originalMidspan\}/, "a Midspan hover title should expose the pre-movement value");

console.log("UI accessibility source tests passed");
