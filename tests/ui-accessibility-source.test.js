"use strict";

const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");

const appPath = path.join(__dirname, "..", "js", "ui", "app.js");
const controlsPath = path.join(__dirname, "..", "js", "ui", "controls.js");
const rendererPath = path.join(__dirname, "..", "js", "ui", "pole-renderer.js");
const excelReviewRendererPath = path.join(__dirname, "..", "js", "ui", "excel-review-renderer.js");
const stylesPath = path.join(__dirname, "..", "css", "styles.css");
const appSource = fs.readFileSync(appPath, "utf8");
const controlsSource = fs.readFileSync(controlsPath, "utf8");
const rendererSource = fs.readFileSync(rendererPath, "utf8");
const excelReviewRendererSource = fs.readFileSync(excelReviewRendererPath, "utf8");
const stylesSource = fs.readFileSync(stylesPath, "utf8");

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
assert.match(poleHeader, /<h3[^>]*data-inline-pole-edit[^>]*data-pole="\$\{escapeHtml\(poleId\)\}"[^>]*title="Double-click to edit pole ID"/, "the pole ID heading should advertise inline editing");
assert.match(controlsSource, /\[data-inline-pole-edit\].*?beginInlinePoleEdit/s, "double-clicking a pole ID should start its inline editor");
assert.match(controlsSource, /S\.renamePole\(poleId, nextPoleId\)/, "the inline pole editor should rename the real pole ID in the store");
assert.doesNotMatch(poleHeader, />Hide<|>Delete</, "pole header actions should no longer show text labels");
assert.match(rendererSource, /class="collapse-hint" aria-hidden="true">\s*<svg class="collapse-show-icon"/i, "Imported Power disclosure should use an eye icon instead of Show text");
assert.match(rendererSource, /<svg class="collapse-hide-icon"/i, "the expanded Imported Power disclosure should show the hide-eye icon");
assert.match(rendererSource, /<details class="subsection wide equipment-section" open>\s*<summary class="subsection-collapse-summary">\s*<h4>Equipment<\/h4>\s*\$\{renderCollapseEyeIcon\(\)\}\s*<\/summary>[\s\S]*?\$\{renderEquipmentTable\(poleId\)\}/i, "Equipment should use the same collapsible eye control as Imported Power and start expanded");
assert.match(rendererSource, /const equipmentSection = oldCard\.querySelector\("\.equipment-section"\);[\s\S]*?const equipmentSectionOpen = equipmentSection \? equipmentSection\.open : false;/, "Equipment disclosure state should be captured before the pole card is refreshed");
assert.match(rendererSource, /const nextEquipmentSection = nextCard\.querySelector\("\.equipment-section"\);\s*if \(nextEquipmentSection\) nextEquipmentSection\.open = equipmentSectionOpen;/, "Equipment disclosure state should survive pole card refreshes");
assert.match(stylesSource, /\.collapse-hint\s*\{[^}]*margin-right:\s*8px/s, "Imported Power eye icon should sit slightly left of the right edge");
assert.match(stylesSource, /\.equipment-section\[open\] \.collapse-show-icon/, "expanded Equipment should switch to the hide-eye icon");
assert.match(rendererSource, /Original Midspan before movements: \$\{originalMidspan\}/, "a Midspan hover title should expose the pre-movement value");

console.log("UI accessibility source tests passed");
