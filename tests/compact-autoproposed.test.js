"use strict";

const assert = require("node:assert/strict");
const fs = require("node:fs");
const vm = require("node:vm");
const path = require("node:path");

function parseHeight(value) {
  if (typeof value === "number" && Number.isFinite(value)) return value;
  const raw = String(value ?? "").trim();
  if (!raw) return null;
  if (/^-?\d+(?:\.\d+)?$/.test(raw)) return Number(raw);
  const sign = raw.startsWith("-") ? -1 : 1;
  const clean = raw.replace(/^-/, "");
  const feetMatch = clean.match(/(\d+)\s*'/);
  const inchMatch = clean.match(/(\d+)\s*"/);
  if (!feetMatch && !inchMatch) return null;
  return sign * ((Number(feetMatch?.[1] || 0) * 12) + Number(inchMatch?.[1] || 0));
}

function formatHeight(total) {
  const sign = total < 0 ? "-" : "";
  const abs = Math.abs(Math.round(total));
  const feet = Math.floor(abs / 12);
  const inches = abs % 12;
  return `${sign}${feet}'${inches ? `${inches}\"` : ""}`;
}

function baseState(profile = "INTEC") {
  return {
    settings: {
      projectProfile: profile,
      proposedOwner: "Wecom",
      attachmentMessengerSize: "",
      fiberSizes: {},
      mrCase: "LOWER"
    },
    poles: {},
    spans: {},
    spanSides: {},
    spanComms: {},
    makeReadyReferences: [],
    mr: []
  };
}

let state = baseState();
let download = null;
const window = {
  AppStore: {
    getState: () => state,
    setState: next => { state = next; },
    updateSetting: (field, value) => { state.settings[field] = value; },
    getPole: poleId => state.poles[poleId] || null
  },
  HeightUtils: { parseHeight, formatHeight },
  Calculations: { recalculateAll() {} },
  ProjectExport: { downloadJson: (filename, payload) => { download = { filename, payload }; } },
  MRLogic: {
    generateMRForPole(poleId) {
      state.mr = state.mr.filter(item => item.poleId !== poleId);
      state.mr.push({ poleId, spanId: "", owner: "MR", text: "Relocate CTL drop at HOA 18'10\" to HOA 20'.", imported: false });
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

state = baseState("INTEC");
state.makeReadyReferences = [{ attachmentFiber: "144CT Fiber" }];
assert.equal(api.applyAttachmentDefaults(state), true);
assert.equal(state.settings.attachmentMessengerSize, "0.242");
assert.equal(state.settings.fiberSizes["144CT Fiber"], "0.51");

state = baseState("METRONET");
state.makeReadyReferences = [{ attachmentSizeRaw: "6.6M 144CT Fiber (E/W)" }];
api.applyAttachmentDefaults(state);
assert.equal(state.settings.attachmentMessengerSize, "0.242");
assert.equal(state.settings.fiberSizes["144CT Fiber"], "0.51");

state.settings.attachmentMessengerSize = "0.188";
state.settings.fiberSizes["144CT Fiber"] = "0.44";
assert.equal(api.applyAttachmentDefaults(state), false);
assert.equal(state.settings.attachmentMessengerSize, "0.188");
assert.equal(state.settings.fiberSizes["144CT Fiber"], "0.44");

state = baseState("INTEC");
state.settings.attachmentMessengerSize = "0.242";
state.settings.fiberSizes = { "144CT Fiber": "0.51" };
state.poles = {
  P1: { poleId: "P1" },
  P2: { poleId: "P2" },
  "P3 UG": { poleId: "P3 UG" },
  P4: { poleId: "P4", standaloneProposedHOA: "18'10\"" }
};
state.spans = {
  S1: { spanId: "S1", fromPole: "P1", toPole: "P2", type: "Fore Span", direction: "E", bearingDegrees: 81.94, lengthDisplay: "51'8\"" },
  S2: { spanId: "S2", fromPole: "P1", toPole: "P3 UG", type: "Back Span", direction: "W", bearingDegrees: 265.06, lengthDisplay: "167'9\"" },
  S3: { spanId: "S3", fromPole: "P1", toPole: "P4", type: "Other", direction: "S", bearingDegrees: 171.83, lengthDisplay: "109'7\"" }
};
state.spanSides = {
  S1_P1: { spanId: "S1", poleId: "P1", proposedHOA: "23'", endDrop: "-4\"", proposedHOAChange: "22'8\"" }
};
state.makeReadyReferences = [
  { poleId: "P1", attachmentFiber: "144CT Fiber", attachmentDirectionTokens: ["E", "W", "S"] }
];
state.spanComms = {
  regularFore: { spanId: "S1", poleId: "P1", owner: "CenturyLink", ownerBase: "CenturyLink", wireId: "W1", existingHOA: "18'10\"", existingHOAChange: "20'" },
  regularOtherRelation: { spanId: "S3", poleId: "P1", owner: "CenturyLink", ownerBase: "CenturyLink", wireId: "W1", existingHOA: "18'10\"", existingHOAChange: "20'" },
  service: { spanId: "S1", poleId: "P1", owner: "CenturyLink", ownerBase: "CenturyLink", wireId: "W2", existingHOA: "18'10\"", existingHOAChange: "20'", serviceDrop: true },
  dg: { spanId: "S1", poleId: "P1", owner: "CATV", wireId: "W3", existingHOA: "23'6\"", existingHOAChange: "21'8\"", downGuy: true },
  transfer: { spanId: "S1", poleId: "P1", owner: "Cox", wireId: "W4", existingHOA: "18'", existingHOAChange: "20'", transferToNewPole: true },
  resagOnly: { spanId: "S1", poleId: "P1", owner: "Drop", wireId: "W5", existingHOA: "17'", serviceDrop: true, resagServiceDrop: true }
};

const payload = api.buildCompactPayload(state);
assert.equal(JSON.stringify(payload.sizes), JSON.stringify({ messenger: 0.242, fiber: { 144: 0.51 } }));
assert.equal(payload.owner, "Wecom");
const p1 = payload.poles.find(pole => pole.id === "P1");
assert.ok(p1);
assert.equal(p1.spans.length, 3);
const fore = p1.spans.find(span => span.kind === "F");
assert.equal(JSON.stringify(fore), JSON.stringify({ to: "P2", kind: "F", bearing: 81.94, length: 620, hoa: 276, fiber: 144, endDrop: -4, nextHoa: 272 }));
const backUg = p1.spans.find(span => span.kind === "B");
assert.equal(backUg.ug, true);
assert.equal("fiber" in backUg, false);
assert.equal("hoa" in backUg, false);
const other = p1.spans.find(span => span.kind === "O");
assert.equal(other.hoa, 276);
assert.equal(other.fiber, 144);
assert.equal(p1.moves.length, 4, "same physical wire across two spans must collapse, but service/fiber must remain separate");
assert.ok(p1.moves.some(move => move.owner === "CTL" && !move.service && move.from === 226 && move.to === 240));
assert.ok(p1.moves.some(move => move.owner === "CTL" && move.service === true && move.from === 226 && move.to === 240));
assert.ok(p1.moves.some(move => move.owner === "CATV" && move.dg === true && move.from === 282 && move.to === 260));
assert.ok(p1.moves.some(move => move.owner === "Cox" && move.from === 216 && move.to === 240));
const p4 = payload.poles.find(pole => pole.id === "P4");
assert.equal(p4.terminalHoa, 226);
assert.equal("moves" in p4, false);

assert.equal(JSON.stringify(api.validationErrors(payload)), JSON.stringify([]));
const missingMessenger = structuredClone(payload);
missingMessenger.sizes.messenger = null;
assert.equal(JSON.stringify(api.validationErrors(missingMessenger)), JSON.stringify(["Missing Messenger Size"]));
const missingFiber = structuredClone(payload);
missingFiber.sizes.fiber[144] = null;
assert.equal(JSON.stringify(api.validationErrors(missingFiber)), JSON.stringify(["Missing 144CT Fiber Size"]));
const onlyUg = { sizes: { messenger: null, fiber: {} }, owner: "Wecom", poles: [{ id: "P1", spans: [{ to: "P2 UG", kind: "F", ug: true }] }] };
assert.equal(JSON.stringify(api.validationErrors(onlyUg)), JSON.stringify([]));

state.mr = [];
window.MRLogic.generateMRForPole("P1");
const mr = state.mr.find(item => item.poleId === "P1").text;
assert.match(mr, /Relocate CTL drop/);
assert.match(mr, /At HOA 18'10" raise CTL to HOA 20'\./, "regular fiber movement must remain beside service drop movement");

state.settings.attachmentMessengerSize = "";
state.settings.fiberSizes = { "72CT Fiber": "0.40" };
state.makeReadyReferences = [{ poleId: "P1", attachmentFiber: "72CT Fiber", attachmentDirectionTokens: ["E", "W", "S"] }];
download = null;
assert.equal(api.exportCompactProposedJson(), false);
assert.equal(download, null, "invalid export must not download a file");
state.settings.attachmentMessengerSize = "0.242";
assert.equal(api.exportCompactProposedJson(), true);
assert.equal(download.filename, "pole_job_AutoProposed.json");
assert.equal("app" in download.payload, false);
assert.equal("commMakeReady" in download.payload, false);

console.log("compact-autoproposed tests passed");
