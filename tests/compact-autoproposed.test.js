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

function canonicalPoleIdentity(value) {
  const parts = String(value || "").trim().split(/\s+/).filter(Boolean);
  while (parts.length > 1 && /^(STEEL|UG|PCO)$/i.test(parts[parts.length - 1])) parts.pop();
  return parts.join(" ").toUpperCase();
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
    getPole: poleId => state.poles[poleId] || null,
    canonicalPoleIdentity
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
state.settings.proposedOwner = "Cox";
state.makeReadyReferences = [{ attachmentFiber: "96CT Fiber" }];
assert.equal(api.applyAttachmentDefaults(state), true);
assert.equal(state.settings.attachmentMessengerSize, "0.25");
assert.equal(state.settings.fiberSizes["96CT Fiber"], "0.53");

state.settings.attachmentMessengerSize = "0.22";
state.settings.fiberSizes["96CT Fiber"] = "0.49";
assert.equal(api.applyAttachmentDefaults(state), false);
assert.equal(state.settings.attachmentMessengerSize, "0.22");
assert.equal(state.settings.fiberSizes["96CT Fiber"], "0.49");

state = baseState("INTEC");
state.makeReadyReferences = [{ attachmentFiber: "96CT Fiber" }];
window.AppStore.updateSetting("proposedOwner", "Cox");
assert.equal(state.settings.attachmentMessengerSize, "0.25");
assert.equal(state.settings.fiberSizes["96CT Fiber"], "0.53");

state = baseState("INTEC");
state.settings.attachmentMessengerSize = "0.242";
state.settings.fiberSizes = { "144CT Fiber": "0.51" };
state.poles = {
  P1: { poleId: "P1" },
  P2: { poleId: "P2" },
  "P3 UG": { poleId: "P3 UG", ugActive: true },
  P4: { poleId: "P4", standaloneProposedHOA: "18'10\"" },
  P5: { poleId: "P5" },
  P6: { poleId: "P6" },
  P10: { poleId: "P10" },
  P11: { poleId: "P11", ugActive: true },
  P12: { poleId: "P12" }
};
state.spans = {
  S1: { spanId: "S1", fromPole: "P1", toPole: "P2", type: "Fore Span", direction: "E", bearingDegrees: 81.94, lengthDisplay: "51'8\"" },
  S1_DUP: { spanId: "S1_DUP", fromPole: "P1", toPole: "P2", type: "Other", direction: "E", bearingDegrees: 81.94, lengthDisplay: "51'8\"", sourceSpanId: "S1" },
  S2: { spanId: "S2", fromPole: "P1", toPole: "P3 UG", type: "Back Span", direction: "W", bearingDegrees: 265.06, lengthDisplay: "167'9\"" },
  S3: { spanId: "S3", fromPole: "P1", toPole: "P4", type: "Other", direction: "S", bearingDegrees: 171.83, lengthDisplay: "109'7\"" },
  S4: { spanId: "S4", fromPole: "P1", toPole: "P5", type: "Other", direction: "N", bearingDegrees: 0, lengthDisplay: "70'" },
  S6: { spanId: "S6", fromPole: "P6", toPole: "P7", type: "Fore Span", direction: "E", bearingDegrees: 102.64, lengthDisplay: "173'6\"" },
  P10_TO_P11: { spanId: "P10_TO_P11", fromPole: "P10", toPole: "P11", type: "Fore Span", direction: "N", bearingDegrees: 0, lengthDisplay: "100'" },
  P10_TO_P12: { spanId: "P10_TO_P12", fromPole: "P10", toPole: "P12", type: "Other", direction: "E", bearingDegrees: 90, lengthDisplay: "80'" },
  P11_TO_P10: { spanId: "P11_TO_P10", fromPole: "P11", toPole: "P10", type: "Back Span", direction: "S", bearingDegrees: 180, lengthDisplay: "100'" },
  P11_TO_P12: { spanId: "P11_TO_P12", fromPole: "P11", toPole: "P12", type: "Fore Span", direction: "E", bearingDegrees: 90, lengthDisplay: "60'" }
};
state.spanSides = {
  S1_P1: { spanId: "S1", poleId: "P1", proposedHOA: "23'" },
  S1_DUP_P1: { spanId: "S1_DUP", poleId: "P1", proposedHOA: "23'", endDrop: "-4\"", proposedHOAChange: "22'8\"", isAdditionalProposed: true },
  S6_P6: { spanId: "S6", poleId: "P6", proposedHOA: "19'10\"" },
  P10_E: { spanId: "P10_TO_P12", poleId: "P10", proposedHOA: "22'" },
  P11_S: { spanId: "P11_TO_P10", poleId: "P11", proposedHOA: "20'10\"" }
};
state.makeReadyReferences = [
  { poleId: "P1", attachmentFiber: "144CT Fiber", attachmentDirectionTokens: ["E", "W", "S"] },
  { poleId: "P10", attachmentSizeRaw: "144CT Fiber N/E", attachmentDirectionTokens: ["N", "E"] },
  { poleId: "P11", attachmentSizeRaw: "144CT Fiber S/E", attachmentDirectionTokens: ["S", "E"] },
  { poleId: "P6", attachmentFiber: "144CT Fiber", attachmentDirectionTokens: ["SE"] }
];
state.spanComms = {
  regularFore: { spanId: "S1", poleId: "P1", owner: "CenturyLink", ownerBase: "CenturyLink", wireId: "W1", existingHOA: "18'10\"", existingHOAChange: "20'" },
  regularOtherRelation: { spanId: "S3", poleId: "P1", owner: "CenturyLink", ownerBase: "CenturyLink", wireId: "W1", existingHOA: "18'10\"", existingHOAChange: "20'" },
  duplicatePhysicalInstruction: { spanId: "S4", poleId: "P1", owner: "CenturyLink", ownerBase: "CenturyLink", wireId: "W9", existingHOA: "18'10\"", existingHOAChange: "20'" },
  service: { spanId: "S1", poleId: "P1", owner: "CenturyLink", ownerBase: "CenturyLink", wireId: "W2", existingHOA: "18'10\"", existingHOAChange: "20'", serviceDrop: true, downGuy: true },
  dg: { spanId: "S1", poleId: "P1", owner: "CATV", wireId: "W3", existingHOA: "23'6\"", existingHOAChange: "21'8\"", downGuy: true },
  duplicateDgInstruction: { spanId: "S3", poleId: "P1", owner: "CATV", wireId: "W8", existingHOA: "23'6\"", existingHOAChange: "21'8\"", downGuy: false },
  transfer: { spanId: "S1", poleId: "P1", owner: "Cox", wireId: "W4", existingHOA: "18'", existingHOAChange: "20'", transferToNewPole: true },
  resagOnly: { spanId: "S1", poleId: "P1", owner: "Drop", wireId: "W5", existingHOA: "17'", serviceDrop: true, resagServiceDrop: true }
};

const payload = api.buildCompactPayload(state);
assert.equal(JSON.stringify(payload.sizes), JSON.stringify({ messenger: 0.242, fiber: { 144: 0.51 } }));
assert.equal(payload.owner, "Wecom");

const p1 = payload.poles.find(pole => pole.id === "P1");
assert.ok(p1);
assert.equal(p1.spans.length, 4, "duplicate physical spans must export once");
const fore = p1.spans.find(span => span.to === "P2");
assert.equal(JSON.stringify(fore), JSON.stringify({ to: "P2", kind: "F", bearing: 81.94, length: 620, hoa: 276, fiber: 144, endDrop: -4, nextHoa: 272 }));
const backUg = p1.spans.find(span => span.to === "P3 UG");
assert.equal(backUg.ug, true);
assert.equal("fiber" in backUg, false);
assert.equal("hoa" in backUg, false);
const other = p1.spans.find(span => span.to === "P4");
assert.equal(other.hoa, 276);
assert.equal(other.fiber, 144);

state.makeReadyReferences.push({
  poleId: "P4",
  attachmentSizeRaw: "72CT Fiber (N)",
  attachmentType: "Overlash",
  attachmentDirectionTokens: ["N"]
});
state.spanComms.overlashMessenger = {
  spanId: "S3",
  poleId: "P1",
  owner: "COMMUNICATION > Wecom Inc",
  rawOwner: "COMMUNICATION > Wecom Inc",
  size: "Telco Bundles > 0.5\" Communication Bundle Msgr:0.242\""
};
const overlash = api.buildCompactPayload(state).poles.find(pole => pole.id === "P1").spans.find(span => span.to === "P4");
assert.equal(overlash.overlash, true, "an Overlash Make Ready row on the opposite endpoint must reuse the job-owner messenger");
const northNotSelected = p1.spans.find(span => span.to === "P5");
assert.equal("fiber" in northNotSelected, false, "a span outside Fiber E/W/S must stay geometry-only");
assert.equal("hoa" in northNotSelected, false);

const tolerant = payload.poles.find(pole => pole.id === "P6").spans.find(span => span.to === "P7");
assert.equal(tolerant.hoa, 238, "a nearby directional fiber reference must preserve the proposed HOA");
assert.equal(tolerant.fiber, 144, "a nearby directional fiber reference must be accepted");

assert.equal(p1.moves.length, 4, "identical physical instructions must collapse while service and normal remain separate");
assert.ok(p1.moves.some(move => move.owner === "CTL" && !move.service && move.from === 226 && move.to === 240));
const ctlDropMove = p1.moves.find(move => move.owner === "CTL" && move.service === true && move.from === 226 && move.to === 240);
assert.ok(ctlDropMove);
assert.equal("dg" in ctlDropMove, false, "service drop moves must never carry DG");
assert.ok(p1.moves.some(move => move.owner === "CATV" && move.dg === true && move.from === 282 && move.to === 260));
assert.ok(p1.moves.some(move => move.owner === "Cox" && move.from === 216 && move.to === 240));

const p4 = payload.poles.find(pole => pole.id === "P4");
assert.equal(p4.terminalHoa, 226);
assert.equal("moves" in p4, false);

const p10 = payload.poles.find(pole => pole.id === "P10");
const p10ToP11 = p10.spans.find(span => span.to === "P11");
const p10ToP12 = p10.spans.find(span => span.to === "P12");
assert.equal(p10ToP11.ug, true, "a normal pole must mark only the span toward a fully UG destination as UG");
assert.equal("fiber" in p10ToP11, false);
assert.equal(p10ToP12.ug, undefined);
assert.equal(p10ToP12.fiber, 144, "the other selected aerial direction must still receive fiber");

const p11 = payload.poles.find(pole => pole.id === "P11");
assert.ok(p11.spans.every(span => span.ug === true), "a pole with the UG action active must export every outgoing span as UG");
assert.ok(p11.spans.every(span => !("fiber" in span) && !("hoa" in span)));

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
assert.doesNotMatch(mr, /Relocate CTL drop[^\n]*with DG/i, "service drop MR must never include with DG");
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
