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
  const sign = raw.startsWith("-") ? -1 : 1;
  const clean = raw.replace(/^-/, "");
  const feet = Number(clean.match(/(\d+)\s*'/)?.[1] || 0);
  const inches = Number(clean.match(/(\d+)\s*"/)?.[1] || 0);
  return feet || inches ? sign * (feet * 12 + inches) : null;
}

function canonicalPoleIdentity(value) {
  const parts = String(value || "").trim().split(/\s+/).filter(Boolean);
  while (parts.length > 1 && /^(STEEL|UG|PCO)$/i.test(parts[parts.length - 1])) parts.pop();
  return parts.join(" ").toUpperCase();
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
    P2: { poleId: "P2", pcoActive: true, standaloneProposedHOA: "21'4\"" },
    P3: { poleId: "P3", ugActive: true, standaloneProposedHOA: "20'10\"" },
    P4: { poleId: "P4", pcoActive: true, ugActive: true, standaloneProposedHOA: "20'6\"" },
    P5: { poleId: "P5", pcoActive: true, standaloneProposedHOA: "19'6\"" }
  },
  spans: {
    P1_TO_P2: { spanId: "P1_TO_P2", fromPole: "P1", toPole: "P2", type: "Fore Span", direction: "E", bearingDegrees: 90, lengthDisplay: "100'" },
    P2_TO_P1: { spanId: "P2_TO_P1", fromPole: "P2", toPole: "P1", type: "Back Span", direction: "W", bearingDegrees: 270, lengthDisplay: "100'" },
    P1_TO_P3: { spanId: "P1_TO_P3", fromPole: "P1", toPole: "P3", type: "Other", direction: "N", bearingDegrees: 0, lengthDisplay: "80'" },
    P3_TO_P1: { spanId: "P3_TO_P1", fromPole: "P3", toPole: "P1", type: "Other", direction: "S", bearingDegrees: 180, lengthDisplay: "80'" },
    P4_TO_P1: { spanId: "P4_TO_P1", fromPole: "P4", toPole: "P1", type: "Other", direction: "S", bearingDegrees: 180, lengthDisplay: "70'" }
  },
  spanSides: {
    P1_TO_P2_P1: { spanId: "P1_TO_P2", poleId: "P1", proposedHOA: "22'", endDrop: "-4\"", proposedHOAChange: "21'8\"" },
    P2_TO_P1_P2: { spanId: "P2_TO_P1", poleId: "P2", proposedHOA: "21'4\"", endDrop: "2\"", proposedHOAChange: "21'6\"" },
    P1_TO_P3_P1: { spanId: "P1_TO_P3", poleId: "P1", proposedHOA: "22'" },
    P3_TO_P1_P3: { spanId: "P3_TO_P1", poleId: "P3", proposedHOA: "20'10\"" },
    P4_TO_P1_P4: { spanId: "P4_TO_P1", poleId: "P4", proposedHOA: "20'6\"" }
  },
  spanComms: {
    P1_CATV: { spanId: "P1_TO_P2", poleId: "P1", owner: "CATV", existingHOA: "20'", existingHOAChange: "21'" },
    P2_CATV: { spanId: "P2_TO_P1", poleId: "P2", owner: "CATV", existingHOA: "18'6\"", existingHOAChange: "20'4\"", downGuy: true },
    P3_CTL: { spanId: "P3_TO_P1", poleId: "P3", owner: "CenturyLink", existingHOA: "18'", existingHOAChange: "19'" },
    P4_3J: { spanId: "P4_TO_P1", poleId: "P4", owner: "3J Communications", existingHOA: "17'", existingHOAChange: "18'" },
    P5_CATV: { spanId: "", poleId: "P5", owner: "CATV", existingHOA: "18'", existingHOAChange: "19'" }
  },
  makeReadyReferences: [
    { poleId: "P1", attachmentFiber: "144CT Fiber", attachmentDirectionTokens: ["E", "N"] },
    { poleId: "P2", attachmentFiber: "144CT Fiber", attachmentDirectionTokens: ["W"] },
    { poleId: "P3", attachmentFiber: "144CT Fiber", attachmentDirectionTokens: ["S"] },
    { poleId: "P4", attachmentFiber: "144CT Fiber", attachmentDirectionTokens: ["S"] },
    { poleId: "P5", attachmentFiber: "144CT Fiber", attachmentDirectionTokens: ["E"] }
  ],
  mr: []
};

const window = {
  AppStore: {
    getState: () => state,
    getPole: poleId => state.poles[poleId] || null,
    canonicalPoleIdentity,
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

const normal = payload.poles.find(pole => pole.id === "P1");
assert.ok(normal);
assert.equal(normal.moves.length, 1, "normal poles must keep local movements");
const normalToPco = normal.spans.find(span => span.to === "P2");
assert.equal(normalToPco.ug, undefined, "a neighboring span toward PCO must remain aerial");
assert.equal(normalToPco.hoa, 264);
assert.equal(normalToPco.fiber, 144);
assert.equal(normalToPco.endDrop, -4);
const normalToUg = normal.spans.find(span => span.to === "P3");
assert.equal(normalToUg.ug, true, "a neighboring span toward UG must remain UG");
assert.equal("hoa" in normalToUg, false);
assert.equal("fiber" in normalToUg, false);

const pco = payload.poles.find(pole => pole.id === "P2");
assert.ok(pco);
assert.equal("moves" in pco, false, "PCO poles must not export make-space movements");
assert.equal("terminalHoa" in pco, false, "PCO poles must not export terminal attachments");
assert.ok(pco.spans.every(span => span.ug !== true), "PCO local spans must not be converted to UG");
assert.ok(pco.spans.every(span => !("hoa" in span) && !("fiber" in span)));
assert.ok(pco.spans.every(span => !("endDrop" in span) && !("nextHoa" in span)));

const ug = payload.poles.find(pole => pole.id === "P3");
assert.ok(ug);
assert.equal("moves" in ug, false);
assert.equal("terminalHoa" in ug, false);
assert.ok(ug.spans.every(span => span.ug === true));
assert.ok(ug.spans.every(span => !("hoa" in span) && !("fiber" in span)));

const pcoAndUg = payload.poles.find(pole => pole.id === "P4");
assert.ok(pcoAndUg);
assert.equal("moves" in pcoAndUg, false);
assert.equal("terminalHoa" in pcoAndUg, false);
assert.ok(pcoAndUg.spans.every(span => span.ug === true), "UG must take priority over PCO");

assert.equal(payload.poles.some(pole => pole.id === "P5"), false, "a PCO pole with only local moves/terminal and no spans must be omitted");
assert.equal(window.CompactAutoProposed.isPolePco(state, "P2"), true);
assert.equal(window.CompactAutoProposed.isPolePco(state, "P1"), false);

assert.equal(window.ProjectExport.exportProposedJson(), true);
const downloadedNormal = downloaded.payload.poles.find(pole => pole.id === "P1");
const downloadedPco = downloaded.payload.poles.find(pole => pole.id === "P2");
const downloadedUg = downloaded.payload.poles.find(pole => pole.id === "P3");
assert.equal(downloadedNormal.spans.find(span => span.to === "P2").fiber, 144);
assert.equal("moves" in downloadedPco, false);
assert.ok(downloadedPco.spans.every(span => span.ug !== true));
assert.equal("moves" in downloadedUg, false);
assert.ok(downloadedUg.spans.every(span => span.ug === true));

console.log("compact-autoproposed UG/PCO guard tests passed");
