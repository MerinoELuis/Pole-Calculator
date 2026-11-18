"use strict";

const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");

const repoRoot = path.join(__dirname, "..");
const exportSource = fs.readFileSync(path.join(repoRoot, "js", "export", "json-export.js"), "utf8");
const appSource = fs.readFileSync(path.join(repoRoot, "js", "ui", "app.js"), "utf8");

const state = {
  jobName: "Debug Save Test",
  importedFileName: "test.xlsx",
  poles: { P1: {} },
  spans: {},
  spanComms: {},
  spanPower: {}
};
const diagnostics = {
  status: "failed",
  attemptedAt: "2026-11-17T12:00:00.000Z",
  completedAt: "2026-11-17T12:00:01.000Z",
  fileName: "Debug_Save_Test_Pole_Calculator.json",
  method: "browser-download",
  location: "Browser download location",
  locationKnown: false,
  stage: "fallback_download",
  reason: "Browser blocked the download",
  fallbackFrom: { stage: "write_selected_file", name: "NotAllowedError", message: "Permission denied" },
  warning: null
};
const loadDiagnostics = {
  status: "failed",
  attemptedAt: "2026-11-17T12:01:00.000Z",
  completedAt: "2026-11-17T12:01:01.000Z",
  fileName: "CTL-PRJ27564_Pole_Calculator.json",
  method: "html-file-input",
  location: "Browser compatibility chooser",
  locationKnown: false,
  stage: "import_json",
  reason: "Invalid JSON",
  fallbackFrom: { stage: "choose_json_file", name: "AbortError", message: "Picker aborted" },
  warning: null,
  error: { stage: "import_json", name: "SyntaxError", message: "Invalid JSON" }
};

let downloadedText = "";
const anchor = { click() {}, remove() {} };
class TestBlob {
  constructor(parts) { this.text = parts.join(""); }
}
const sandbox = {
  window: {
    AppDeploymentVersion: { source: "development" },
    AppStore: { getState: () => state, getOwnerCanonical: row => row.owner || "" },
    Calculations: { recalculateAll() {} }
  },
  Blob: TestBlob,
  URL: {
    createObjectURL(blob) { downloadedText = blob.text; return "blob:test"; },
    revokeObjectURL() {}
  },
  document: { body: { appendChild() {} }, createElement: () => anchor },
  Date,
  JSON,
  Map,
  Object
};

vm.runInNewContext(exportSource, sandbox, { filename: "json-export.js" });
sandbox.window.ProjectExport.exportDebugJson(diagnostics, loadDiagnostics);

const debug = JSON.parse(downloadedText);
assert.equal(debug.exportType, "calculation-debug");
assert.deepEqual(JSON.parse(JSON.stringify(debug.saveDiagnostics)), diagnostics);
assert.deepEqual(JSON.parse(JSON.stringify(debug.loadDiagnostics)), loadDiagnostics);
assert.match(appSource, /getSaveDiagnostics\?\.\(\) \|\| null,[\s\S]*getLoadDiagnostics\?\.\(\) \|\| null/);

console.log("Save diagnostics Debug export tests passed.");
