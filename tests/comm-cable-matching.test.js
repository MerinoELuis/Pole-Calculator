"use strict";

const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");

const window = {
  localStorage: { getItem: () => null, setItem: () => {} },
  MRLogic: { generateAllMR: () => {}, generateMRForPole: () => {} },
  Validations: { validateAll: () => {}, validatePole: () => {}, validateSpan: () => {} }
};
const sandbox = { window, console, Date, Set, Map, Math, JSON, Object, Array, String, Number, Boolean, RegExp };

for (const file of [
  "domain/height-math.js",
  "core/project-config.js",
  "domain/pole-resolution.js",
  "core/state-compat.js",
  "core/store-identity.js",
  "core/store.js",
  "domain/comm-movements.js",
  "domain/midspan-engine.js",
  "domain/pole-inset-calculation.js",
  "domain/clearance-rules.js"
]) {
  const modulePath = path.join(__dirname, "..", "js", file);
  vm.runInNewContext(fs.readFileSync(modulePath, "utf8"), sandbox, { filename: modulePath });
}
const reconcilerPath = path.join(__dirname, "..", "js", "import", "update-data-reconciler.js");
vm.runInNewContext(fs.readFileSync(reconcilerPath, "utf8"), sandbox, { filename: reconcilerPath });

const S = window.AppStore;
const C = window.Calculations;
const updateReconciler = window.UpdateDataReconciler.create({
  S,
  ownerCanonicalOf: row => S.getOwnerCanonical(row),
  safeJobFilePart: value => String(value || "")
});
const owner = "COMMUNICATION > Telco";

function seedTwoCableSpan() {
  S.resetState();
  S.upsertPole(S.createPole({ poleId: "P1", lowPower: "30'" }));
  S.upsertPole(S.createPole({ poleId: "P2", lowPower: "30'" }));
  S.upsertSpan(S.createSpan("S1", "P1", "P2", "E", "", { type: "Fore Span" }));
  S.upsertSpanComm({
    spanId: "S1", poleId: "P1", owner, ownerBase: owner,
    wireId: "P1-A", size: "1in Telco Bundle", existingHOA: "20'", midspan: "18'"
  });
  S.upsertSpanComm({
    spanId: "S1", poleId: "P1", owner, ownerBase: owner,
    wireId: "P1-B", size: "1in Telco Bundle", existingHOA: "20'", midspan: "16'"
  });
  S.upsertSpanComm({
    spanId: "S1", poleId: "P2", owner, ownerBase: owner,
    wireId: "P2-A", size: "1in Telco Bundle", existingHOA: "20'", existingHOAChange: "19'", midspan: "18'"
  });
  S.upsertSpanComm({
    spanId: "S1", poleId: "P2", owner, ownerBase: owner,
    wireId: "P2-B", size: "1in Telco Bundle", existingHOA: "20'", midspan: "16'"
  });
}

seedTwoCableSpan();
let warnings = C.getCommPairingWarnings("P1");
assert.equal(warnings.length, 1, "different remote cable movements must produce a pairing warning");
assert.equal(C.getRemoteCommMatchResult("S1", "P1", owner, "P1-A").status, "AMBIGUOUS");

C.recalculateAll();
let localA = S.getSpanComm("S1", "P1", owner, "P1-A");
assert.equal(localA.midspanMatchStatus, "AMBIGUOUS");
assert.equal(localA.calculatedMidspan, "", "an unresolved match must not publish a guessed Midspan");
assert.equal(localA.clearanceMSStatus, "MATCH_REQUIRED", "clearance evaluation must be blocked until matching");
assert.equal(C.displayMidspanForComm(localA), "");
assert.equal(C.getReferenceMidspansForSpanSide("S1", "P1").length, 0, "an unresolved cable must not enter clearance references");

warnings = C.getCommPairingWarnings("P1");
const warning = warnings[0];
const remoteForLocal = local => warning.remoteRows.find(row => row.wireId === local.wireId.replace("P1-", "P2-"));
const pairs = warning.localRows.map(local => ({
  local: {
    spanId: local.spanId,
    poleId: local.poleId,
    ownerCanonical: S.getOwnerCanonical(local),
    wireId: local.wireId
  },
  remote: {
    spanId: remoteForLocal(local).spanId,
    poleId: remoteForLocal(local).poleId,
    ownerCanonical: S.getOwnerCanonical(remoteForLocal(local)),
    wireId: remoteForLocal(local).wireId
  }
}));
assert.ok(warning.localRows.some(row => row.wireId === "P1-A"), "test fixture should contain the first local cable");

const invalidPairs = pairs.map(pair => ({ ...pair, remote: pairs[0].remote }));
const rejected = C.setCommRemoteMatchPairs("P1", "P2", warning.ownerToken, invalidPairs);
assert.equal(rejected.ok, false, "a remote cable cannot be assigned to more than one local cable");

const saved = C.setCommRemoteMatchPairs("P1", "P2", warning.ownerToken, pairs);
assert.equal(saved.ok, true);
C.recalculateSpansForPole("P1");
C.recalculateSpansForPole("P2");
assert.equal(C.getCommPairingWarnings("P1").length, 0, "confirmed one-to-one matches clear the warning");
localA = S.getSpanComm("S1", "P1", owner, "P1-A");
const localB = S.getSpanComm("S1", "P1", owner, "P1-B");
assert.equal(localA.calculatedMidspan, "17'6\"", "cable A receives its own remote 12-inch movement, halved");
assert.equal(localB.calculatedMidspan, "16'", "cable B does not borrow cable A's remote movement");
assert.equal(localA.manualRemoteMatch.wireId, "P2-A");
assert.equal(S.getSpanComm("S1", "P2", owner, "P2-A").manualRemoteMatch.wireId, "P1-A");

S.setState(JSON.parse(JSON.stringify(S.getState())));
C.recalculateAll();
localA = S.getSpanComm("S1", "P1", owner, "P1-A");
assert.equal(localA.manualRemoteMatch.wireId, "P2-A", "manual cable links survive state Save/Load normalization");
assert.equal(localA.calculatedMidspan, "17'6\"");
assert.equal(C.getCommPairingWarnings("P1").length, 0);

const previousUpdate = JSON.parse(JSON.stringify(S.getState()));
const importedUpdate = JSON.parse(JSON.stringify(previousUpdate));
importedUpdate.spanComms = {};
Object.values(previousUpdate.spanComms).forEach(row => {
  const refreshed = { ...row, wireId: row.wireId + "-new", manualRemoteMatch: null };
  const key = S.keyForSpanComm(
    refreshed.spanId, refreshed.poleId, S.getOwnerCanonical(refreshed), refreshed.wireId
  );
  importedUpdate.spanComms[key] = refreshed;
});
const reconciled = updateReconciler.mergeImportedUpdate(previousUpdate, importedUpdate);
const remappedLocalA = Object.values(reconciled.spanComms).find(row => row.wireId === "P1-A-new");
assert.equal(remappedLocalA.manualRemoteMatch.wireId, "P2-A-new", "Update Data remaps manual cable links when Wire IDs are refreshed");
S.setState(reconciled);
C.recalculateAll();
assert.equal(S.getSpanComm("S1", "P1", owner, "P1-A-new").calculatedMidspan, "17'6\"");
assert.equal(C.getCommPairingWarnings("P1").length, 0);

S.resetState();
S.upsertPole(S.createPole({ poleId: "P3", lowPower: "30'" }));
S.upsertPole(S.createPole({ poleId: "P4", lowPower: "30'" }));
S.upsertSpan(S.createSpan("S2", "P3", "P4", "E", "", { type: "Fore Span" }));
S.upsertSpanComm({ spanId: "S2", poleId: "P3", owner, ownerBase: owner, wireId: "P3-A", existingHOA: "20'", midspan: "18'" });
S.upsertSpanComm({ spanId: "S2", poleId: "P3", owner, ownerBase: owner, wireId: "P3-B", existingHOA: "19'", midspan: "17'" });
S.upsertSpanComm({ spanId: "S2", poleId: "P4", owner, ownerBase: owner, wireId: "P4-A", existingHOA: "20'", existingHOAChange: "19'", midspan: "18'" });
C.recalculateAll();
assert.equal(C.getCommPairingWarnings("P3").length, 1, "one moved remote row must not be reused for multiple same-owner cables");
assert.equal(C.getRemoteCommMatchResult("S2", "P3", owner, "P3-B").status, "AMBIGUOUS");
assert.equal(S.getSpanComm("S2", "P3", owner, "P3-A").midspanMatchStatus, "AMBIGUOUS");
assert.equal(S.getSpanComm("S2", "P3", owner, "P3-B").midspanMatchStatus, "AMBIGUOUS");

// The real-world pattern: P10 has one measured Midspan source; P11's Back
// Span rows have only On-pole heights. Match the top drop to the top drop,
// while treating the two equal-height lower drops as interchangeable.
S.resetState();
S.upsertPole(S.createPole({ poleId: "P10", lowPower: "30'" }));
S.upsertPole(S.createPole({ poleId: "P11", lowPower: "30'" }));
S.upsertSpan(S.createSpan("P10-P11-F", "P10", "P11", "E", "", { type: "Fore Span" }));
S.upsertSpan(S.createSpan("P10-P11-B", "P11", "P10", "W", "", { type: "Back Span" }));
const dropSize = "Service Drops > TELCO Drop > Slack";
const bundleSize = "TELCO > 1\" Telco 6M Messenger";
[
  { pole: "P10", wire: "P10-D1", size: dropSize, hoa: "18'2\"", change: "17'8\"", midspan: "17'" },
  { pole: "P10", wire: "P10-D2", size: dropSize, hoa: "18'2\"", change: "17'8\"", midspan: "15'2\"" },
  { pole: "P10", wire: "P10-D3", size: dropSize, hoa: "18'2\"", change: "17'8\"", midspan: "14'7\"" },
  { pole: "P10", wire: "P10-M", size: bundleSize, hoa: "18'2\"", change: "17'8\"", midspan: "16'7\"" },
  { pole: "P11", wire: "P11-D1", size: dropSize, hoa: "19'2\"", change: "", midspan: "" },
  { pole: "P11", wire: "P11-D2", size: dropSize, hoa: "17'4\"", change: "16'8\"", midspan: "" },
  { pole: "P11", wire: "P11-D3", size: dropSize, hoa: "17'4\"", change: "16'8\"", midspan: "" },
  { pole: "P11", wire: "P11-M", size: bundleSize, hoa: "17'4\"", change: "", midspan: "" }
].forEach(row => S.upsertSpanComm({
  spanId: row.pole === "P10" ? "P10-P11-F" : "P10-P11-B",
  poleId: row.pole,
  owner,
  ownerBase: owner,
  wireId: row.wire,
  size: row.size,
  serviceDrop: row.size === dropSize,
  existingHOA: row.hoa,
  existingHOAChange: row.change,
  midspan: row.midspan
}));
assert.equal(C.getCommPairingWarnings("P10").length, 0, "exact cable type plus one-sided Midspan rank should resolve the span");
assert.equal(C.getRemoteCommMatchResult("P10-P11-F", "P10", owner, "P10-D1").row.wireId, "P11-D1", "highest drop Midspan maps to the highest remote On-pole attachment");
assert.equal(C.getRemoteCommMatchResult("P10-P11-F", "P10", owner, "P10-D2").row.wireId, "P11-D2", "second drop maps into the tied lower-height group");
assert.equal(C.getRemoteCommMatchResult("P10-P11-F", "P10", owner, "P10-D3").row.wireId, "P11-D3", "third drop maps into the same interchangeable group");
assert.equal(C.getCommMatchMidspanDisplay(S.getSpanComm("P10-P11-B", "P11", owner, "P11-D1")), "", "P11 has no Midspan source to display");
assert.notEqual(C.getRemoteCommMatchResult("P10-P11-F", "P10", owner, "P10-M").row.wireId, "P11-D1", "messenger cannot be paired with a service drop");
assert.equal(C.getRemoteCommMatchResult("P10-P11-F", "P10", owner, "P10-M").row.wireId, "P11-M", "messenger is matched only to the same cable type");
assert.equal(C.getCommPairingWarnings("P11").length, 0, "P11's equal-height lower drops are interchangeable without a false warning");

// When both ends contain Midspans, those duplicate measurements must not
// override the On-pole ordering used to match physical cable positions.
S.resetState();
S.upsertPole(S.createPole({ poleId: "P9", lowPower: "30'" }));
S.upsertPole(S.createPole({ poleId: "P10", lowPower: "30'" }));
S.upsertSpan(S.createSpan("S5", "P9", "P10", "E", "", { type: "Fore Span" }));
[
  { poleId: "P9", wireId: "P9-A", existingHOA: "20'", midspan: "14'" },
  { poleId: "P9", wireId: "P9-B", existingHOA: "19'", midspan: "18'" },
  { poleId: "P10", wireId: "P10-A", existingHOA: "20'", midspan: "17'" },
  { poleId: "P10", wireId: "P10-B", existingHOA: "19'", midspan: "15'" }
].forEach(row => S.upsertSpanComm({
  spanId: "S5", owner, ownerBase: owner, size: "1in Telco Bundle", ...row
}));
assert.equal(C.getRemoteCommMatchResult("S5", "P9", owner, "P9-A").row.wireId, "P10-A", "when both ends have Midspan values, a reversed Midspan rank does not override the On-pole match");

// Without any imported/manual Midspan on either end, cable pairing cannot
// alter a Midspan calculation and should not ask the user to confirm matches.
S.resetState();
S.upsertPole(S.createPole({ poleId: "P5", lowPower: "30'" }));
S.upsertPole(S.createPole({ poleId: "P6", lowPower: "30'" }));
S.upsertSpan(S.createSpan("S3", "P5", "P6", "E", "", { type: "Fore Span" }));
[
  { poleId: "P5", wireId: "P5-1", existingHOA: "20'" },
  { poleId: "P5", wireId: "P5-2", existingHOA: "19'" },
  { poleId: "P6", wireId: "P6-1", existingHOA: "20'" },
  { poleId: "P6", wireId: "P6-2", existingHOA: "19'" }
].forEach(row => S.upsertSpanComm({
  spanId: "S3", owner, ownerBase: owner, size: "0.5in Telco", ...row
}));
assert.equal(C.getCommPairingWarnings("P5").length, 0, "no Midspan source means no Match Required warning");

// Exact cable category is mandatory; a Telco drop cannot match a CATV drop.
S.resetState();
S.upsertPole(S.createPole({ poleId: "P7", lowPower: "30'" }));
S.upsertPole(S.createPole({ poleId: "P8", lowPower: "30'" }));
S.upsertSpan(S.createSpan("S4", "P7", "P8", "E", "", { type: "Fore Span" }));
S.upsertSpanComm({ spanId: "S4", poleId: "P7", owner, ownerBase: owner, wireId: "P7-T", size: "0.5in Telco", existingHOA: "20'", midspan: "18'", serviceDrop: true });
S.upsertSpanComm({ spanId: "S4", poleId: "P8", owner, ownerBase: owner, wireId: "P8-C", size: "0.5in CATV", existingHOA: "20'", existingHOAChange: "19'", serviceDrop: true });
assert.equal(C.getRemoteCommMatchResult("S4", "P7", owner, "P7-T").status, "MISSING", "a similar-size CATV drop is not a compatible Telco cable");
assert.equal(C.setCommRemoteMatchPairs("P7", "P8", "", [{
  local: { spanId: "S4", poleId: "P7", ownerCanonical: S.getOwnerCanonical(S.getSpanComm("S4", "P7", owner, "P7-T")), wireId: "P7-T" },
  remote: { spanId: "S4", poleId: "P8", ownerCanonical: S.getOwnerCanonical(S.getSpanComm("S4", "P8", owner, "P8-C")), wireId: "P8-C" }
}]).ok, false, "manual matching must also reject incompatible cable types");
assert.equal(C.canMatchCommRows(
  { owner, size: "0.5in Telco", serviceDrop: true },
  { owner, size: "0.5\" Telco", serviceDrop: true }
), true, "equivalent inch notation should normalize without changing the cable type");

// A blank cable type can be matched safely when there is exactly one cable
// from the same owner and service-drop class at each end of the physical span.
S.resetState();
S.upsertPole(S.createPole({ poleId: "P1", lowPower: "30'" }));
S.upsertPole(S.createPole({ poleId: "P2", lowPower: "30'" }));
S.upsertSpan(S.createSpan("S6-F", "P1", "P2", "E", "", { type: "Fore Span" }));
S.upsertSpan(S.createSpan("S6-B", "P2", "P1", "W", "", { type: "Back Span" }));
S.upsertSpanComm({
  spanId: "S6-F", poleId: "P1", owner, ownerBase: owner,
  wireId: "P1-MANUAL", size: "", existingHOA: "20'6\"", existingHOAChange: "19'8\"", midspan: "19'2\""
});
S.upsertSpanComm({
  spanId: "S6-B", poleId: "P2", owner, ownerBase: owner,
  wireId: "P2-IMPORT", size: "Comm Messengers Type M > 6M D:0.27\" > Static",
  existingHOA: "19'4\"", existingHOAChange: "17'4\"", serviceDrop: false
});
const p1Manual = S.getSpanComm("S6-F", "P1", owner, "P1-MANUAL");
const uniqueUnknownMatch = C.getRemoteCommMatchResult("S6-F", "P1", owner, p1Manual.wireId, p1Manual);
assert.equal(uniqueUnknownMatch.status, "UNIQUE", "one unknown-type cable may use the sole compatible endpoint cable");
assert.equal(uniqueUnknownMatch.row.wireId, "P2-IMPORT");
assert.equal(C.canMatchCommRows(p1Manual, uniqueUnknownMatch.row), true, "the unique unknown-type cable should be available in Match Cables");
assert.equal(C.getCommPairingWarnings("P2").length, 0, "an incoming Back Span is reviewed from its Fore Span pole only");

// Unknown size must not turn into a guess when several cables exist remotely.
S.upsertSpanComm({
  spanId: "S6-B", poleId: "P2", owner, ownerBase: owner,
  wireId: "P2-CATV", size: "0.5in CATV", existingHOA: "19'", serviceDrop: false
});
assert.equal(C.getRemoteCommMatchResult("S6-F", "P1", owner, p1Manual.wireId, p1Manual).status, "AMBIGUOUS",
  "a cable with unknown type must require manual review when multiple opposite cables are available");
assert.equal(C.canMatchCommRows(
  { owner, size: "", serviceDrop: false },
  { owner, size: "0.5in Telco", serviceDrop: true }
), false, "unknown cable type must not override the service-drop class");

console.log("Comm cable matching tests passed.");
