"use strict";

const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");

let downloaded = null;
const window = {
  AppStore: {
    canonicalPoleIdentity(value) {
      return String(value || "").trim().toUpperCase().replace(/\s+(STEEL|RECOLLECT|T|UG|PCO)$/i, "");
    }
  },
  AppPersistence: { safeJobFilePart: value => String(value || "pole_job").replace(/\s+/g, "_") },
  ExcelImport: {
    pick(row, names) {
      for (const name of names) if (Object.prototype.hasOwnProperty.call(row, name)) return row[name];
      return "";
    }
  },
  ProjectExport: { downloadJson: (filename, payload) => { downloaded = { filename, payload }; } }
};
const exportPath = path.join(__dirname, "..", "js", "export", "autoexisting-export.js");
vm.runInNewContext(fs.readFileSync(exportPath, "utf8"), { window, Number, String, Object, Array, RegExp }, { filename: exportPath });
const api = window.AutoExistingExport;

const state = {
  jobName: "CTL PRJ27564",
  settings: { projectProfile: "INTEC", plaExistingLoadingProfile: "MEDIUM" },
  poles: {
    P1: { poleId: "P1", sequence: "P1", isGenerated: false, metadata: {} },
    P2: { poleId: "P2", sequence: "P2", isGenerated: false, metadata: { plaExisting: { windSpeedMph: "115" } } },
    P3: { poleId: "P3", sequence: "P3", isGenerated: false, metadata: { plaExisting: {} } },
    "P4-not-in-collection": { poleId: "P4-not-in-collection", sequence: "P4", isGenerated: false, metadata: {} },
    "Unknown-temp": { poleId: "Unknown-temp", sequence: "P4", isGenerated: true, metadata: {} }
  },
  poleClassChecks: [
    { poleId: "P1", yearInstalled: "2017" },
    { poleId: "P2", yearInstalled: "2018" }
  ],
  poleIdAliases: {},
  excelReviewSource: { collection: { rows: [
    { Id: "P1", "Year Installed": "2017" },
    { Id: "P2", "Year Installed": "2018" },
    { Id: "P3", "Year Installed": "" }
  ] } }
};

const result = api.inspect(state);
assert.equal(result.canExport, false, "export must wait until every pole has a year and modern wind data");
assert.deepEqual(Array.from(result.plans.map(plan => plan.id)), ["P1", "P2", "P3"], "PLA Existing must include only poles found in Collection");
assert.deepEqual(Array.from(result.missingYearPoleIds), ["P3"], "a pole with no year stays unassigned");
assert.equal(result.payload.poles.find(pole => pole.id === "P3").yearInstalled, null);
assert.deepEqual(Array.from(result.payload.poles.find(pole => pole.id === "P3").loadings), [], "missing year must not select a loading");

const oldPole = result.payload.poles.find(pole => pole.id === "P1");
assert.equal(oldPole.loadingProfile, "PRE_2018");
assert.equal(oldPole.loadings.length, 1);
assert.equal(oldPole.loadings[0].caseId, "NESC_23_EXT_WIND_250C_GRADE_C_LE_100_MPH");
assert.equal(oldPole.loadings[0].windSpeedMph, 90);

const mediumPole = result.payload.poles.find(pole => pole.id === "P2");
assert.equal(mediumPole.loadings.length, 3, "2018+ Medium uses the base and both additional loadings");
assert.equal(mediumPole.loadings[1].windSpeedMph, 115);
assert.equal(mediumPole.loadings[2].grade, "C", "250D must be emitted as Grade C, not the supplied Grade B placeholder");
assert.equal(mediumPole.loadings[2].windSpeedMph, 30);
assert.equal(mediumPole.loadings[2].radialIceInches, 0.25);
assert.equal(result.payload.loadingPrincipal, "MEDIUM", "the principal loading is job-wide and explicit in AutoExisting");

state.settings.plaExistingLoadingProfile = "LIGHT";
const lightResult = api.inspect(state);
const lightPole = lightResult.payload.poles.find(pole => pole.id === "P2");
assert.equal(lightPole.loadings.length, 2, "2018+ Light gets only the base and 250C extra-wind cases");
assert.equal(lightPole.loadings[0].windPressurePsf, 9);
assert.equal(lightPole.loadings[1].caseId, "NESC_23_EXT_WIND_250C_GRADE_C_GT_100_MPH");

state.poles.P3.metadata.plaExisting.yearInstalled = "2016";
assert.equal(api.getYearInfo(state, "P3").year, 2016, "a year entered in PLA Existing is used when Excel is blank");
assert.equal(api.inspect(state).canExport, true, "all known pole years and modern wind data allow export");
state.poles.P2.metadata.plaExisting.windSpeedMph = "100";
assert.equal(api.inspect(state).canExport, false, "the modern wind case requires a speed strictly greater than 100 mph");
state.poles.P2.metadata.plaExisting.windSpeedMph = "120";
state.settings.projectProfile = "OLSSON_OPPD";
assert.equal(api.inspect(state).canExport, false, "AutoExisting is INTEC-only");
state.settings.projectProfile = "INTEC";
const exportResult = api.exportAutoExistingJson(state);
assert.equal(exportResult.ok, true);
assert.equal(downloaded.filename, "CTL_PRJ27564_AutoExisting.json");
assert.equal(downloaded.payload.format, "AutoExisting");

delete state.poles.P3.metadata.plaExisting.yearInstalled;
const rendererPath = path.join(__dirname, "..", "js", "ui", "pla-existing-renderer.js");
vm.runInNewContext(fs.readFileSync(rendererPath, "utf8"), { window, String }, { filename: rendererPath });
const results = { innerHTML: "" };
const summary = { innerHTML: "" };
const exportButton = { disabled: false, title: "", classList: { toggle() {} } };
const renderer = window.PlaExistingRenderer.create({
  S: { getState: () => state },
  els: { plaExistingResults: results, plaExistingSummary: summary, exportAutoExistingBtn: exportButton },
  escapeHtml: value => String(value ?? "").replace(/[&<>\"]/g, char => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", "\"": "&quot;" })[char]),
  wireEditableEvents() {}
});
renderer.render();
assert.match(results.innerHTML, /YEAR INSTALLED/);
assert.match(results.innerHTML, /data-field="yearInstalled"/, "the operator can fill missing year in PLA Existing");
assert.match(results.innerHTML, /pole-class-year-warning-icon/, "2018+ warning icon is shown in the new loading table");
assert.match(results.innerHTML, /data-scope="plaExistingGlobal" data-field="loadingPrincipal"/, "principal loading is selected once for the whole job");
assert.match(results.innerHTML, /<th>LOADINGS<\/th>/);
assert.doesNotMatch(results.innerHTML, /EXISTING LOADING|ADDITIONAL LOADINGS|data-field="loadingProfile"/);
assert.doesNotMatch(results.innerHTML, /9 psf/, "loading names render without extra pressure suffixes");
assert.match(results.innerHTML, /pla-existing-year-value/);
assert.match(results.innerHTML, /pla-existing-wind-field/);
const lightPoleRow = results.innerHTML.match(/<tr data-pla-existing-pole="P2">([\s\S]*?)<\/tr>/)?.[1] || "";
assert.match(lightPoleRow, /NESC 23 Ext Wind \(250C\)/);
assert.doesNotMatch(lightPoleRow, /NESC 23 Light \(250B\)/, "the global principal Light case is not repeated in each pole row");
assert.equal((lightPoleRow.match(/<li>/g) || []).length, 1, "Light displays only its one additional loading on 2018+ poles");
state.settings.plaExistingLoadingProfile = "MEDIUM";
renderer.render();
const mediumPoleRow = results.innerHTML.match(/<tr data-pla-existing-pole="P2">([\s\S]*?)<\/tr>/)?.[1] || "";
assert.doesNotMatch(mediumPoleRow, /NESC Medium \(250B\)/, "the global principal Medium case is not repeated in each pole row");
assert.match(mediumPoleRow, /NESC 23 Ext Wind &amp; Ice \(250D\)/);
assert.equal((mediumPoleRow.match(/<li>/g) || []).length, 2, "Medium displays both additional loadings on 2018+ poles");
const styles = fs.readFileSync(path.join(__dirname, "..", "css", "styles.css"), "utf8");
assert.match(styles, /\.pla-existing-load-list\s*\{[^}]*list-style:\s*none/s, "loading list markers are hidden");
assert.match(styles, /\.pla-existing-wind-field\s*\{[^}]*display:\s*grid[^}]*justify-items:\s*center[^}]*width:\s*100%/s, "wind input box is centered against the full cell width");
assert.match(styles, /\.pla-existing-wind-field \.input\s*\{[^}]*width:\s*96px[^}]*text-align:\s*center/s, "wind input box is smaller and its number centered");
assert.match(styles, /\.pla-existing-wind-field > span\s*\{[^}]*position:\s*absolute/s, "mph suffix does not shift the centered wind input box");
assert.match(styles, /\.pla-existing-year-value\s*\{[^}]*position:\s*relative[^}]*text-align:\s*center/s, "the year remains centered independently of its warning icon");
assert.equal(exportButton.disabled, true, "AutoExisting export stays disabled while any pole year is missing");
assert.doesNotMatch(results.innerHTML, /Year Installed<\/th><th>Status/, "year is no longer part of the Pole Type Check table");

state.settings.projectProfile = "OLSSON_OPPD";
renderer.render();
assert.equal(exportButton.disabled, true, "AutoExisting export is disabled outside INTEC");
assert.match(results.innerHTML, /available only for INTEC/);

const html = fs.readFileSync(path.join(__dirname, "..", "index.html"), "utf8");
assert.ok(html.indexOf('data-view-tab="poleClass"') < html.indexOf('data-view-tab="plaExisting"'));
assert.ok(html.indexOf('data-view-tab="plaExisting"') < html.indexOf('data-view-tab="fiber"'));
assert.match(html, /data-view-tab="plaExisting">PLA Existing<\/button>/);

const schema = JSON.parse(fs.readFileSync(path.join(__dirname, "..", "schemas", "autoexisting.schema.json"), "utf8"));
assert.equal(schema.properties.format.const, "AutoExisting");
assert.equal(schema.properties.project.const, "INTEC");
assert.deepEqual(schema.properties.loadingPrincipal.enum, ["MEDIUM", "LIGHT"]);
assert.equal(schema.$defs.loading.properties.grade.enum.includes("C"), true);

console.log("AutoExisting export and PLA Existing tests passed.");
