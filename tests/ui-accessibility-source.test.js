"use strict";

const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");

const appPath = path.join(__dirname, "..", "js", "ui", "app.js");
const indexPath = path.join(__dirname, "..", "index.html");
const controlsPath = path.join(__dirname, "..", "js", "ui", "controls.js");
const rendererPath = path.join(__dirname, "..", "js", "ui", "pole-renderer.js");
const excelReviewRendererPath = path.join(__dirname, "..", "js", "ui", "excel-review-renderer.js");
const stylesPath = path.join(__dirname, "..", "css", "styles.css");
const appSource = fs.readFileSync(appPath, "utf8");
const indexSource = fs.readFileSync(indexPath, "utf8");
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
assert.match(poleHeader, /data-match-pole[^>]*aria-label="Match cables for \$\{escapeHtml\(poleId\)\}"[^>]*><svg/i, "Match should use an accessible pole-header icon");
assert.match(poleHeader, /pole-header-match-btn/, "Match should have its own styling hook");
assert.match(poleHeader, /data-delete-pole[^>]*aria-label="Delete \$\{escapeHtml\(poleId\)\}"[^>]*><svg/i, "Delete should use an accessible icon button");
assert.match(controlsSource, /\[data-match-pole\].*?editCommPairing/s, "the pole-header Match icon should open the cable matching dialog");
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
assert.match(appSource, /comm-match-dialog-wide/, "Match Cables should use the wide two-pole dialog layout");
assert.match(appSource, /getCommMatchMidspanDisplay\(row\)/, "the Match table should show Midspan only when that endpoint provides it");
assert.match(appSource, /Type not provided/, "cables without imported size/type should be identified as missing type, not labeled Cable");
assert.match(appSource, /cableParts\.length > 2 \? cableParts\.slice\(0, -1\)\.join\(" > "\)/,
  "cable labels should omit the trailing construction/status detail such as Static or Slack");
assert.match(appSource, /columnWidths\.map\(width => `<col style="width:\$\{width\}">`\)/,
  "the Match table should allocate the available width to its data columns");
assert.doesNotMatch(appSource, /data-match-remote-hoa|<th>On pole<\/th>\$\{showRemoteMidspan/, "the matched endpoint's On-pole column should not duplicate the HOA already shown in each selector option");
assert.match(appSource, /commPairingRowLabel\(remote\) \+ " · On pole " \+ \(remote\.existingHOA \|\| "—"\)/,
  "remote On-pole height should remain visible in the cable selector options");
assert.match(appSource, /data-match-span/, "the Match dialog should allow switching physical spans");
assert.match(appSource, /spanRelationOf\(span\)/, "Match span selection should use the store's span-relation helper");
assert.match(appSource, /\.filter\(span => span\.fromPole === poleId && !\/back\\s\*span\|backspan\/i\.test\(spanRelationOf\(span\)\)\)/,
  "Match Cables should list only forward spans belonging to the selected pole");
assert.match(appSource, /function commPairingPendingSelections\([\s\S]*?selections\.set\(commPairingRowKey\(local\), remoteIndex < 0 \? "" : String\(remoteIndex\)\)/,
  "saved cable matches should be restored as row-keyed selections after switching spans");
assert.match(appSource, /renderCommPairingTable\(poleId, activeGroup, commPairingPendingSelections\(poleId, activeGroup, pendingByOtherPole\)\)/,
  "switching back to a span should rebuild its selections instead of passing raw stored pairs");
assert.doesNotMatch(appSource, /Calculations\.spanRelationOf/, "Match span selection must not call an unexposed calculation helper");
assert.match(stylesSource, /\.mini-btn\.pole-header-match-btn\s*\{[^}]*background:\s*#15803d/s, "Match icon should use the requested green color");
assert.match(stylesSource, /\.comm-match-dialog-wide\s*\{[^}]*place-self:\s*center/s, "Match Cables dialog should be centered in the viewport");
assert.match(stylesSource, /\.comm-match-table\s*\{[^}]*table-layout:\s*fixed/s, "Match Cables columns should use a stable full-width layout");
assert.match(stylesSource, /\.comm-match-select\s*\{[^}]*width:\s*100%/s, "Match cable selector should fill its column");
assert.match(stylesSource, /\.comm-match-table\.has-local-midspan[\s\S]*?border-left:\s*1px solid var\(--line\)/s,
  "a subtle table-colored divider should separate the two pole groups");
const reviewActionMarkup = indexSource.match(/<div class="excel-review-actions">([\s\S]*?)<\/div>/)?.[1];
assert.ok(reviewActionMarkup, "Excel Review should have a dedicated action toolbar");
assert.ok(reviewActionMarkup.indexOf("ignoreFinalWarningsBtn") < reviewActionMarkup.indexOf("expandExcelReviewIssuesBtn"), "expand-issues icon should follow the ignore controls");
assert.ok(reviewActionMarkup.indexOf("expandExcelReviewIssuesBtn") < reviewActionMarkup.indexOf("rerunExcelReviewBtn"), "expand-issues icon should sit before the rerun icon");
assert.match(reviewActionMarkup, /id="expandExcelReviewIssuesBtn"[^>]*aria-expanded="false"[^>]*aria-label="Expand all poles with active Excel Review findings"[\s\S]*?excel-review-expand-icon[\s\S]*?excel-review-collapse-icon/i,
  "expand/collapse issues action should be an accessible toggle icon");
assert.match(reviewActionMarkup, /id="rerunExcelReviewBtn"[^>]*aria-label="Re-run Excel Review"[\s\S]*?<svg/i,
  "rerun review should be an accessible icon button");
assert.match(reviewActionMarkup, /class="excel-review-rerun-icon"[\s\S]*M6\.7 6\.7a7\.5 7\.5 0 1 0 10\.6 0[\s\S]*excel-review-rerun-arrowhead/,
  "rerun review should use a prominent circular arrow with a filled arrowhead");
assert.match(stylesSource, /\.excel-review-icon-btn\[data-expanded="true"\] \.excel-review-collapse-icon/,
  "the review issue toggle should switch to its collapse icon when all problem poles are open");
assert.match(appSource, /expandExcelReviewIssuesBtn: qs\("expandExcelReviewIssuesBtn"\)/,
  "the review expansion control should be connected to the page elements");
assert.match(excelReviewRendererSource, /data-review-has-problem="\$\{hasActiveProblem \? "true" : "false"\}"/,
  "only poles with active findings should be marked for bulk expansion");
assert.match(excelReviewRendererSource, /excelReviewExpandedPoleIds/,
  "expanded Excel Review poles should be preserved in saved UI state");

console.log("UI accessibility source tests passed");
