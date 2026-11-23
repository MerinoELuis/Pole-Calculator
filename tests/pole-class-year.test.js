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
    "js/ui/poles/anchor-renderer.js",
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

function renderProjectViews(projectProfile, activeView) {
  const state = { settings: { projectProfile }, ui: { activeView } };
  const makeElement = dataset => {
    const classes = new Set();
    return {
      dataset,
      disabled: false,
      title: "",
      textContent: dataset.viewTab || "",
      classes,
      classList: {
        toggle(name, enabled) { if (enabled) classes.add(name); else classes.delete(name); },
        contains(name) { return classes.has(name); }
      }
    };
  };
  const tabs = ["calculator", "poleClass", "plaExisting", "fiber", "excelReview"].map(viewTab => makeElement({ viewTab }));
  const panels = ["calculator", "poleClass", "plaExisting", "fiber", "excelReview"].map(viewPanel => makeElement({ viewPanel }));
  const window = {
    ExcelReviewRenderer: { create: () => ({}) },
    ExcelImport: {
      ANSI_POLE_CLASSES: [], ANSI_CLASS_TABLE: {}, ANSI_APPROX_GROUNDLINE_DISTANCE: [],
      pick() { return ""; }
    }
  };
  const document = {
    body: { classList: { toggle() {} } },
    querySelectorAll(selector) { return selector === "[data-view-tab]" ? tabs : panels; }
  };
  const sandbox = { window, document, console, Date, Set, Map, Math, JSON, Object, Array, String, Number, Boolean, RegExp };
  [
    "js/ui/excel-review-renderer.js",
    "js/ui/poles/power-equipment-renderer.js",
    "js/ui/poles/anchor-renderer.js",
    "js/ui/settings-renderer.js",
    "js/ui/pole-renderer.js"
  ].forEach(file => {
    const modulePath = path.join(__dirname, "..", file);
    vm.runInNewContext(fs.readFileSync(modulePath, "utf8"), sandbox, { filename: modulePath });
  });
  const renderer = window.PoleWorkspaceRenderer.create({
    S: { getState: () => state },
    els: {},
    escapeHtml: value => String(value ?? ""),
    setPoleIndexOpen() {},
    updatePoleIndexToggleVisibility() {}
  });
  renderer.renderActiveView();
  return { state, tabs, panels };
}

const intecHtml = renderPoleTypeCheck("INTEC", [
  { poleId: "P18", yearInstalled: "2018", expectedType: "Pine > 2 > 35", calculatedHeight: 35, calculatedClass: "2", status: "OK" },
  { poleId: "P17", yearInstalled: "2017", expectedType: "Pine > 2 > 35", calculatedHeight: 35, calculatedClass: "2", status: "OK" }
]);
assert.doesNotMatch(intecHtml, /Year Installed|pole-class-year-warning-icon/,
  "Year Installed and its load-case warning belong in PLA Existing, not Pole Type Check");
assert.match(intecHtml, /<th>Expected Type<\/th><th>Status<\/th>/,
  "Pole Type Check should keep its original columns without Year Installed");
assert.match(intecHtml, /pole-class-summary[\s\S]*?OK 2/,
  "moving the year display must not change Pole Type Check status counts");
assert.equal((intecHtml.match(/badge changed\">OK<\/span>/g) || []).length, 2,
  "the year display must not change either pole's Height/Class status");

const restoredIntecHtml = renderPoleTypeCheck("INTEC", [
  { poleId: "P18-UPDATED", expectedType: "Pine > 2 > 35", calculatedHeight: 35, calculatedClass: "2", status: "OK" }
], {
  poleIdAliases: { "P18-ORIGINAL": "P18-UPDATED" },
  excelReviewSource: {
    collection: { rows: [{ Id: "P18-ORIGINAL", "Year Installed": 2018 }] }
  }
});
assert.doesNotMatch(restoredIntecHtml, /Year Installed|2018|pole-class-year-warning-icon/,
  "Pole Type Check should not display years even when they can be recovered from Collection");

const metronetHtml = renderPoleTypeCheck("METRONET", [
  { poleId: "P18", yearInstalled: "2020", expectedType: "Pine > 2 > 35", calculatedHeight: 35, calculatedClass: "2", status: "OK" }
]);
assert.doesNotMatch(metronetHtml, /Year Installed|pole-class-year-warning-icon/,
  "the Year Installed column and warning icon must be INTEC-only");
assert.match(metronetHtml, /<th>Expected Type<\/th><th>Status<\/th>/,
  "other projects should retain the original Pole Type Check columns");

const hiddenIntecViews = renderProjectViews("OLSSON_OPPD", "fiber");
assert.equal(hiddenIntecViews.state.ui.activeView, "calculator", "switching away from INTEC while on Fiber returns to Calculator");
for (const view of ["plaExisting", "fiber"]) {
  const tab = hiddenIntecViews.tabs.find(item => item.dataset.viewTab === view);
  const panel = hiddenIntecViews.panels.find(item => item.dataset.viewPanel === view);
  assert.equal(tab.classList.contains("hidden"), true, `${view} tab is hidden outside INTEC`);
  assert.equal(tab.disabled, true, `${view} tab is disabled outside INTEC`);
  assert.equal(panel.classList.contains("hidden"), true, `${view} panel is hidden outside INTEC`);
}
const visibleIntecViews = renderProjectViews("INTEC", "calculator");
for (const view of ["plaExisting", "fiber"]) {
  assert.equal(visibleIntecViews.tabs.find(item => item.dataset.viewTab === view).classList.contains("hidden"), false, `${view} tab is available in INTEC`);
}

const styles = fs.readFileSync(path.join(__dirname, "..", "css", "styles.css"), "utf8");
assert.match(styles, /\.pole-class-year-value\s*\{[^}]*text-align:\s*center/s,
  "the year should stay centered in its column regardless of the warning icon");
assert.match(styles, /\.pole-class-year-warning\s*\{[^}]*position:\s*absolute/s,
  "the warning icon should not participate in the year text layout");

console.log("INTEC Pole Type Check year warning tests passed.");
