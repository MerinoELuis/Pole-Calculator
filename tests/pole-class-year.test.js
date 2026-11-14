"use strict";

const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");

function renderPoleTypeCheck(projectProfile, rows, additionalState = {}) {
  const state = { settings: { projectProfile }, poleClassChecks: rows, ...additionalState };
  const poleClassResults = {
    innerHTML: "",
    insertAdjacentHTML(_position, html) { this.referenceHtml = html; },
    querySelectorAll: () => []
  };
  const window = {
    ExcelImport: {
      ANSI_POLE_CLASSES: [], ANSI_CLASS_TABLE: {}, ANSI_APPROX_GROUNDLINE_DISTANCE: [],
      pick(row, candidates) {
        for (const candidate of candidates) {
          if (Object.prototype.hasOwnProperty.call(row, candidate)) return row[candidate];
        }
        return "";
      }
    }
  };
  const sandbox = { window, console, Date, Set, Map, Math, JSON, Object, Array, String, Number, Boolean, RegExp };
  const files = [
    "js/ui/excel-review-renderer.js",
    "js/ui/poles/power-equipment-renderer.js",
    "js/ui/settings-renderer.js",
    "js/ui/pole-renderer.js"
  ];
  files.forEach(file => {
    const modulePath = path.join(__dirname, "..", file);
    vm.runInNewContext(fs.readFileSync(modulePath, "utf8"), sandbox, { filename: modulePath });
  });
  const renderer = window.PoleWorkspaceRenderer.create({
    S: { getState: () => state },
    els: { poleClassResults },
    escapeHtml: value => String(value ?? "").replace(/[&<>\"]/g, char => ({
      "&": "&amp;", "<": "&lt;", ">": "&gt;", "\"": "&quot;"
    })[char]),
    wireEditableEvents() {},
    bindScrollLinks() {}
  });
  renderer.renderPoleClassResults();
  return poleClassResults.innerHTML;
}

const intecHtml = renderPoleTypeCheck("INTEC", [
  { poleId: "P18", yearInstalled: "2018", expectedType: "Pine > 2 > 35", calculatedHeight: 35, calculatedClass: "2", status: "OK" },
  { poleId: "P17", yearInstalled: "2017", expectedType: "Pine > 2 > 35", calculatedHeight: 35, calculatedClass: "2", status: "OK" }
]);
assert.match(intecHtml, /Expected Type<\/th><th>Year Installed<\/th><th>Status<\/th>/,
  "INTEC should place Year Installed between Expected Type and Status");
assert.equal((intecHtml.match(/pole-class-year-warning-icon/g) || []).length, 1,
  "only INTEC poles installed in 2018 or later should show the warning icon");
assert.ok(intecHtml.indexOf("pole-class-year-warning-icon") < intecHtml.indexOf("pole-class-year-text"),
  "the warning icon should appear to the left of the year");
assert.match(intecHtml, /pole-class-summary[\s\S]*?OK 2/,
  "the year warning must not change Pole Type Check status counts");
assert.equal((intecHtml.match(/badge changed\">OK<\/span>/g) || []).length, 2,
  "year warnings must not change either pole's Height/Class status");

const restoredIntecHtml = renderPoleTypeCheck("INTEC", [
  { poleId: "P18-UPDATED", expectedType: "Pine > 2 > 35", calculatedHeight: 35, calculatedClass: "2", status: "OK" }
], {
  poleIdAliases: { "P18-ORIGINAL": "P18-UPDATED" },
  excelReviewSource: {
    collection: { rows: [{ Id: "P18-ORIGINAL", "Year Installed": 2018 }] }
  }
});
assert.match(restoredIntecHtml, /pole-class-year-warning-icon/,
  "older saved projects should recover the INTEC year from their preserved Collection source");
assert.match(restoredIntecHtml, /pole-class-year-text\">2018<\/span>/,
  "the recovered year should be displayed after a pole rename");

const metronetHtml = renderPoleTypeCheck("METRONET", [
  { poleId: "P18", yearInstalled: "2020", expectedType: "Pine > 2 > 35", calculatedHeight: 35, calculatedClass: "2", status: "OK" }
]);
assert.doesNotMatch(metronetHtml, /Year Installed|pole-class-year-warning-icon/,
  "the Year Installed column and warning icon must be INTEC-only");
assert.match(metronetHtml, /<th>Expected Type<\/th><th>Status<\/th>/,
  "other projects should retain the original Pole Type Check columns");

const styles = fs.readFileSync(path.join(__dirname, "..", "css", "styles.css"), "utf8");
assert.match(styles, /\.pole-class-year-value\s*\{[^}]*text-align:\s*center/s,
  "the year should stay centered in its column regardless of the warning icon");
assert.match(styles, /\.pole-class-year-warning\s*\{[^}]*position:\s*absolute/s,
  "the warning icon should not participate in the year text layout");

console.log("INTEC Pole Type Check year warning tests passed.");
