"use strict";

const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");

const persistenceSource = fs.readFileSync(
  path.join(__dirname, "..", "js", "core", "persistence.js"),
  "utf8"
);

function createPersistence({ showSaveFilePicker, showOpenFilePicker, indexedDB, writableError, downloadError, confirmFallback } = {}) {
  const state = { jobName: "New Job", importedFileName: "", poles: { P1: {} } };
  const downloads = [];
  const toasts = [];
  const importedFiles = [];
  const saveButton = { textContent: "Save", title: "" };
  const handle = {
    name: "New_Job_Pole_Calculator.json",
    async createWritable() {
      if (writableError) throw writableError;
      return {
        written: "",
        async write(text) { this.written = text; },
        async close() {}
      };
    }
  };
  const window = {
    AppStore: { CURRENT_VERSION: 1, getState: () => state },
    showSaveFilePicker: showSaveFilePicker || (async () => handle),
    showOpenFilePicker,
    indexedDB,
    document: {
      body: { appendChild() {} },
      createElement() {
        const listeners = {};
        const input = {
          files: [{ name: "backup.json" }],
          addEventListener(event, callback) { listeners[event] = callback; },
          click() { listeners.change?.(); },
          remove() {}
        };
        return input;
      }
    },
    AppCalculations: { recalculateAll() {} }
  };
  const testConsole = { ...console, warn() {} };
  vm.runInNewContext(persistenceSource, { window, console: testConsole, Date, JSON, Promise });
  const persistence = window.AppPersistence.create({
    S: window.AppStore,
    Calculations: window.AppCalculations,
    ProjectExport: { downloadJson: (...args) => {
      if (downloadError) throw downloadError;
      downloads.push(args);
    } },
    ExcelImport: { importJsonFile: async file => importedFiles.push(file) },
    ExcelReview: { clearResults() {}, runReview() {} },
    els: { saveLocalBtn: saveButton },
    render() {},
    recordUndoSnapshot() {},
    toast: (...args) => toasts.push(args),
    confirmFallback
  });
  return { persistence, downloads, toasts, handle, state, importedFiles };
}

async function run() {
  {
    const { persistence, downloads, toasts, handle } = createPersistence({
      writableError: new Error("The selected destination cannot be written")
    });
    assert.equal(await persistence.saveLocalFile(), true);
    assert.equal(downloads.length, 1, "failed direct write should download a JSON backup");
    assert.equal(downloads[0][0], "New_Job_Pole_Calculator.json");
    assert.equal(downloads[0][1].app, "pole-calculator");
    assert.equal(persistence.hasUnsavedChanges(), false, "fallback download should mark the job saved");
    assert.equal(toasts.at(-1)[1], "warning");
    assert.match(toasts.at(-1)[0], /JSON download was started/);
    assert.equal(handle.name, "New_Job_Pole_Calculator.json");
    const diagnostics = persistence.getSaveDiagnostics();
    assert.equal(diagnostics.status, "download_started");
    assert.equal(diagnostics.method, "browser-download");
    assert.match(diagnostics.location, /usually Downloads/);
    assert.equal(diagnostics.fallbackFrom.message, "The selected destination cannot be written");
  }

  {
    const { persistence, downloads, toasts } = createPersistence({
      indexedDB: {
        open() { throw new Error("IndexedDB is unavailable"); }
      }
    });
    assert.equal(await persistence.saveLocalFile(), true);
    assert.equal(downloads.length, 0, "a completed native file write should not be duplicated");
    assert.equal(persistence.hasUnsavedChanges(), false);
    assert.equal(toasts.at(-1)[1], "success");
    assert.match(toasts.at(-1)[0], /^Saved /);
    const diagnostics = persistence.getSaveDiagnostics();
    assert.equal(diagnostics.status, "saved");
    assert.equal(diagnostics.method, "file-system-access-api");
    assert.equal(diagnostics.fileName, "New_Job_Pole_Calculator.json");
    assert.match(diagnostics.location, /does not expose its full path/);
    assert.equal(diagnostics.warning.stage, "remember_save_location");
  }

  {
    const { persistence, downloads, toasts } = createPersistence({
      showSaveFilePicker: async () => { throw new DOMException("User cancelled", "AbortError"); }
    });
    assert.equal(await persistence.saveLocalFile(), false);
    assert.equal(downloads.length, 0, "cancelling the picker must not start a surprise download");
    assert.equal(toasts.at(-1)[1], "info");
    assert.equal(persistence.getSaveDiagnostics().status, "cancelled");
    assert.match(persistence.getSaveDiagnostics().reason, /browser-download fallback was declined/);
    assert.equal(persistence.getSaveDiagnostics().error.name, "AbortError");
    assert.equal(persistence.getSaveDiagnostics().error.stage, "choose_save_location");
  }

  {
    const { persistence, downloads } = createPersistence({
      showSaveFilePicker: async () => { throw new DOMException("Picker aborted despite selection", "AbortError"); },
      confirmFallback: async () => true
    });
    assert.equal(await persistence.saveLocalFile(), true);
    assert.equal(downloads.length, 1, "confirmed compatibility save should start a browser download");
    assert.equal(persistence.getSaveDiagnostics().status, "download_started");
    assert.equal(persistence.getSaveDiagnostics().fallbackFrom.name, "AbortError");
  }

  {
    const { persistence, importedFiles, toasts } = createPersistence({
      showOpenFilePicker: async () => { throw new DOMException("Picker aborted despite selection", "AbortError"); },
      confirmFallback: async () => true
    });
    assert.equal(await persistence.loadLocalFile(), true);
    assert.equal(importedFiles.length, 1, "confirmed compatibility chooser should import its selected JSON");
    assert.equal(importedFiles[0].name, "backup.json");
    assert.equal(persistence.getLoadDiagnostics().status, "loaded");
    assert.equal(persistence.getLoadDiagnostics().method, "html-file-input");
    assert.equal(persistence.getLoadDiagnostics().fallbackFrom.name, "AbortError");
    assert.equal(toasts.at(-1)[0], "JSON loaded.");
  }

  {
    const { persistence, downloads } = createPersistence({
      writableError: new Error("Selected file is read-only"),
      downloadError: new Error("Browser blocked the download")
    });
    await assert.rejects(persistence.saveLocalFile(), /Browser blocked the download/);
    assert.equal(downloads.length, 0);
    const diagnostics = persistence.getSaveDiagnostics();
    assert.equal(diagnostics.status, "failed");
    assert.equal(diagnostics.stage, "fallback_download");
    assert.equal(diagnostics.reason, "Browser blocked the download");
    assert.equal(diagnostics.fallbackFrom.message, "Selected file is read-only");
  }

  console.log("Persistence save fallback tests passed.");
}

run().catch(error => {
  console.error(error);
  process.exitCode = 1;
});
