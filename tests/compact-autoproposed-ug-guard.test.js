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
    P2: {
      poleId: "P2",
      pcoActive: true,
      standaloneProposedHOA: "21'4\"",
      pcoMRText: "Existing clearance violations on midspan 12'4\" N, replace pole. Transfer CATV & CTL to new pole."
    },
    P3: {
      poleId: "P3",
      ugActive: true,
      standaloneProposedHOA: "20'10\"",
      ugMRText: "Unable to attach due to clearance violation."
    },
    P4: {
      poleId: "P4",
      pcoActive: true,
      ugActive: true,
      standaloneProposedHOA: "20'6\"",
      ugMRText: "Unable to attach due to pole condition."
    },
    P5: {
      poleId: "P5",
      pcoActive: true,
      standaloneProposedHOA: "19'6\"",
      pcoMRText: "Replace pole. Transfer CATV to new pole."
    }
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
    formatHeight
  },
  Calculations: { recalculateAll() {} },
  ProjectExport: {
    downloadJson(filename, payload) {
      downloaded = { filename, payload };
    }
  },
  MRLogic: {
    generateMRForPole(poleId) {
      state.mr = state.mr.filter(item => item.poleId !== poleId);
      const pole = state.poles[poleId];
      const text = pole?.ugActive
        ? pole.ugMRText
        : pole?.pcoActive
          ? pole.pcoMRText
          : "Base Make Ready instruction.";
      state.mr.push({ poleId, spanId: "", owner: "MR", text, imported: false });
      return state.mr.filter(item => item.poleId === poleId);
    },
    generateAllMR() {
      state.mr = [];
      Object.keys(state.poles).forEach(poleId => this.generateMRForPole(poleId));
      return state.mr;
    }
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

window.MRLogic.generateMRForPole("P1");
const normalMr = state.mr.find(item => item.poleId === "P1").text;
assert.match(normalMr, /Base Make Ready instruction\./);
assert.match(normalMr, /At HOA 20' raise CATV to HOA 21'\./, "normal poles must continue showing calculated movement MR");

window.MRLogic.generateMRForPole("P2");
const pcoMr = state.mr.find(item => item.poleId === "P2").text;
assert.equal(pcoMr, state.poles.P2.pcoMRText, "PCO Make Ready must keep only the replacement/transfer template");
assert.doesNotMatch(pcoMr, /^At HOA\b/im, "PCO Make Ready must hide regular attachment movements");
assert.doesNotMatch(pcoMr, /^Relocate\b/im, "PCO Make Ready must hide service-drop movements");

window.MRLogic.generateMRForPole("P3");
const ugMr = state.mr.find(item => item.poleId === "P3").text;
assert.equal(ugMr, state.poles.P3.ugMRText, "UG Make Ready must keep only the UG template");
assert.doesNotMatch(ugMr, /^At HOA\b/im, "UG Make Ready must hide attachment movements");

window.MRLogic.generateMRForPole("P4");
const ugPriorityMr = state.mr.find(item => item.poleId === "P4").text;
assert.equal(ugPriorityMr, state.poles.P4.ugMRText, "UG must control the Make Ready text when UG and PCO are both active");
assert.doesNotMatch(ugPriorityMr, /^At HOA\b/im);

window.MRLogic.generateAllMR();
assert.doesNotMatch(state.mr.find(item => item.poleId === "P2").text, /^At HOA\b/im);
assert.doesNotMatch(state.mr.find(item => item.poleId === "P2").text, /^Relocate\b/im);
assert.doesNotMatch(state.mr.find(item => item.poleId === "P3").text, /^At HOA\b/im);
assert.match(state.mr.find(item => item.poleId === "P1").text, /^At HOA\b/im);

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
