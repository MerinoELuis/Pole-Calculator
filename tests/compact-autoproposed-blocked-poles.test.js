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

function formatHeight(total) {
  const sign = total < 0 ? "-" : "";
  const abs = Math.abs(Math.round(total));
  const feet = Math.floor(abs / 12);
  const inches = abs % 12;
  return `${sign}${feet}'${inches ? `${inches}\"` : ""}`;
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
    fiberSizes: { "144CT Fiber": "0.51" },
    mrCase: "LOWER"
  },
  poles: {
    P1: { poleId: "P1" },
    P2: { poleId: "P2", pcoActive: true, standaloneProposedHOA: "21'4\"", pcoMRText: "Replace pole. Transfer CATV to new pole." },
    P3: { poleId: "P3", ugActive: true, standaloneProposedHOA: "20'10\"", ugMRText: "Unable to attach due to clearance violation." },
    P4: { poleId: "P4", pcoActive: true, ugActive: true, standaloneProposedHOA: "20'6\"", ugMRText: "Unable to attach due to pole condition." },
    P5: { poleId: "P5", pcoActive: true, standaloneProposedHOA: "19'6\"", pcoMRText: "Replace pole." },
    "P6 UG": { poleId: "P6 UG" }
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
    P2_CTL_SERVICE: { spanId: "P2_TO_P1", poleId: "P2", owner: "CenturyLink", existingHOA: "16'10\"", existingHOAChange: "17'2\"", serviceDrop: true },
    P3_CTL: { spanId: "P3_TO_P1", poleId: "P3", owner: "CenturyLink", existingHOA: "18'", existingHOAChange: "19'" },
    P4_3J: { spanId: "P4_TO_P1", poleId: "P4", owner: "3J Communications", existingHOA: "17'", existingHOAChange: "18'" },
    P5_CATV: { spanId: "", poleId: "P5", owner: "CATV", existingHOA: "18'", existingHOAChange: "19'" },
    P6_CTL: { spanId: "", poleId: "P6 UG", owner: "CenturyLink", existingHOA: "18'", existingHOAChange: "19'" }
  },
  makeReadyReferences: [
    { poleId: "P1", attachmentFiber: "144CT Fiber", attachmentDirectionTokens: ["E", "N"] },
    { poleId: "P2", attachmentFiber: "144CT Fiber", attachmentDirectionTokens: ["W"] },
    { poleId: "P3", attachmentFiber: "144CT Fiber", attachmentDirectionTokens: ["S"] },
    { poleId: "P4", attachmentFiber: "144CT Fiber", attachmentDirectionTokens: ["S"] }
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
  HeightUtils: { parseHeight, formatHeight },
  Calculations: { recalculateAll() {} },
  ProjectExport: {
    downloadJson(filename, payload) { downloaded = { filename, payload }; }
  },
  MRLogic: {
    generateMRForPole(poleId) {
      state.mr = state.mr.filter(item => item.poleId !== poleId);
      const pole = state.poles[poleId];
      const base = pole?.ugActive ? pole.ugMRText : pole?.pcoActive ? pole.pcoMRText : "Base Make Ready instruction.";
      state.mr.push({ poleId, spanId: "", owner: "MR", text: base, imported: false });
      return state.mr.filter(item => item.poleId === poleId);
    },
    generateAllMR() { return []; }
  },
  console,
  setTimeout,
  queueMicrotask,
  alert() {}
};

const sourcePath = path.join(__dirname, "..", "js", "compact-autoproposed.js");
vm.runInNewContext(fs.readFileSync(sourcePath, "utf8"), { window }, { filename: sourcePath });
const api = window.CompactAutoProposed;
const payload = api.buildCompactPayload(state);

const normal = payload.poles.find(pole => pole.id === "P1");
assert.ok(normal);
assert.equal(normal.moves.length, 1);
assert.equal(normal.spans.find(span => span.to === "P2").ug, undefined, "normal-to-PCO stays aerial");
assert.equal(normal.spans.find(span => span.to === "P2").fiber, 144);
assert.equal(normal.spans.find(span => span.to === "P3").ug, true, "normal-to-UG stays UG");

const pco = payload.poles.find(pole => pole.id === "P2");
assert.ok(pco);
assert.equal("moves" in pco, false);
assert.equal("terminalHoa" in pco, false);
assert.ok(pco.spans.every(span => span.ug !== true));
assert.ok(pco.spans.every(span => !("hoa" in span) && !("fiber" in span) && !("endDrop" in span) && !("nextHoa" in span)));

const ug = payload.poles.find(pole => pole.id === "P3");
assert.ok(ug);
assert.equal("moves" in ug, false);
assert.equal("terminalHoa" in ug, false);
assert.ok(ug.spans.every(span => span.ug === true));
assert.ok(ug.spans.every(span => !("hoa" in span) && !("fiber" in span)));

const both = payload.poles.find(pole => pole.id === "P4");
assert.ok(both);
assert.ok(both.spans.every(span => span.ug === true), "UG takes priority over PCO");
assert.equal(payload.poles.some(pole => pole.id === "P5"), false, "blocked pole with only local data is omitted");

const namedUg = payload.poles.find(pole => pole.id === "P6 UG");
assert.ok(namedUg, "a pole named with UG remains exportable while its UG action is inactive");
assert.equal(api.isPoleFullyUg(state, "P6 UG"), false);
assert.ok(namedUg.moves.some(move => move.owner === "CTL" && move.from === 216 && move.to === 228));

window.MRLogic.generateMRForPole("P1");
assert.match(state.mr.find(item => item.poleId === "P1").text, /At HOA 20' raise CATV to HOA 21'\./);
window.MRLogic.generateMRForPole("P2");
assert.equal(state.mr.find(item => item.poleId === "P2").text, state.poles.P2.pcoMRText);
window.MRLogic.generateMRForPole("P3");
assert.equal(state.mr.find(item => item.poleId === "P3").text, state.poles.P3.ugMRText);

assert.equal(window.ProjectExport.exportProposedJson(), true);
assert.equal("moves" in downloaded.payload.poles.find(pole => pole.id === "P2"), false);
assert.ok(downloaded.payload.poles.find(pole => pole.id === "P3").spans.every(span => span.ug === true));

console.log("compact-autoproposed blocked-pole tests passed");
