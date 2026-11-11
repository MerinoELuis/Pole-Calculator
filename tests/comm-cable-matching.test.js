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
    wireId: "P1-A", size: "0.5in Telco", existingHOA: "20'", midspan: "18'"
  });
  S.upsertSpanComm({
    spanId: "S1", poleId: "P1", owner, ownerBase: owner,
    wireId: "P1-B", size: "1in Telco", existingHOA: "19'", midspan: "16'"
  });
  S.upsertSpanComm({
    spanId: "S1", poleId: "P2", owner, ownerBase: owner,
    wireId: "P2-A", size: "0.5in Telco", existingHOA: "20'", existingHOAChange: "19'", midspan: "18'"
  });
  S.upsertSpanComm({
    spanId: "S1", poleId: "P2", owner, ownerBase: owner,
    wireId: "P2-B", size: "1in Telco", existingHOA: "19'", midspan: "16'"
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
const localBySize = size => warning.localRows.find(row => row.size === size);
const remoteBySize = size => warning.remoteRows.find(row => row.size === size);
const pairs = warning.localRows.map(local => ({
  local: {
    spanId: local.spanId,
    poleId: local.poleId,
    ownerCanonical: S.getOwnerCanonical(local),
    wireId: local.wireId
  },
  remote: {
    spanId: remoteBySize(local.size).spanId,
    poleId: remoteBySize(local.size).poleId,
    ownerCanonical: S.getOwnerCanonical(remoteBySize(local.size)),
    wireId: remoteBySize(local.size).wireId
  }
}));
assert.ok(localBySize("0.5in Telco"), "test fixture should contain the expected local cable");

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

console.log("Comm cable matching tests passed.");
