"use strict";

const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");

function parseHeight(value) {
  if (typeof value === "number" && Number.isFinite(value)) return value;
  const raw = String(value ?? "").trim();
  if (!raw) return null;
  if (/^-?\d+(?:\.\d+)?$/.test(raw)) return Number(raw);
  const feet = Number(raw.match(/(\d+)\s*'/)?.[1] || 0);
  const inches = Number(raw.match(/(\d+)\s*"/)?.[1] || 0);
  return feet || inches ? feet * 12 + inches : null;
}

let downloaded = null;
const state = {
  settings: {
    projectProfile: "INTEC",
    proposedOwner: "Wecom",
    attachmentMessengerSize: "0.242",
    fiberSizes: { "144CT Fiber": "0.51" }
  },
  poles: {
    P1: { poleId: "P1" },
    P16: { poleId: "P16", ugActive: true, standaloneProposedHOA: "21'4\"" }
  },
  spans: {
    S1: { spanId: "S1", fromPole: "P1", toPole: "P16", type: "Fore Span", direction: "E", bearingDegrees: 90, lengthDisplay: "100'" },
    S16: { spanId: "S16", fromPole: "P16", toPole: "P1", type: "Back Span", direction: "W", bearingDegrees: 270, lengthDisplay: "100'" }
  },
  spanSides: {
    S1_P1: { spanId: "S1", poleId: "P1", proposedHOA: "22'" },
    S16_P16: { spanId: "S16", poleId: "P16", proposedHOA: "21'4\"" }
  },
  spanComms: {
    P16_CATV: { spanId: "S16", poleId: "P16", owner: "CATV", existingHOA: "18'6\"", existingHOAChange: "20'4\"", downGuy: true }
  },
  makeReadyReferences: [
    { poleId: "P1", attachmentFiber: "144CT Fiber", attachmentDirectionTokens: ["E"] },
    { poleId: "P16", attachmentFiber: "144CT Fiber", attachmentDirectionTokens: ["W"] }
  ],
  mr: []
};

const window = {
  AppStore: {
    getState: () => state,
    getPole: poleId => state.poles[poleId] || null,
    setState() {},
    updateSetting() {}
  },
  HeightUtils: {
    parseHeight,
    formatHeight: value => String(value)
  },
  Calculations: { recalculateAll() {} },
  ProjectExport: {
    downloadJson(filename, payload) {
      downloaded = { filename, payload };
    }
  },
  MRLogic: {
    generateMRForPole() { return []; },
    generateAllMR() { return []; }
  },
  console,
  setTimeout,
  queueMicrotask,
  alert() {}
};

const root = path.join(__dirname, "..");
vm.runInNewContext(fs.readFileSync(path.join(root, "js", "compact-autoproposed.js"), "utf8"), { window });
vm.runInNewContext(fs.readFileSync(path.join(root, "js", "compact-autoproposed-ug-guard.js"), "utf8"), { window });

const payload = window.CompactAutoProposed.buildCompactPayload(state);
const p16 = payload.poles.find(pole => pole.id === "P16");
assert.ok(p16);
assert.equal("moves" in p16, false, "fully UG poles must not export make-space movements");
assert.equal("terminalHoa" in p16, false, "fully UG poles must not export terminal attachments");
assert.ok(p16.spans.every(span => span.ug === true));
assert.ok(p16.spans.every(span => !("hoa" in span) && !("fiber" in span)));

assert.equal(window.ProjectExport.exportProposedJson(), true);
const downloadedP16 = downloaded.payload.poles.find(pole => pole.id === "P16");
assert.equal("moves" in downloadedP16, false);
assert.equal("terminalHoa" in downloadedP16, false);
assert.ok(downloadedP16.spans.every(span => span.ug === true));

console.log("compact-autoproposed UG guard tests passed");
