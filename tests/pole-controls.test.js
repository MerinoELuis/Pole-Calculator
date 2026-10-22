"use strict";

const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");

const window = { localStorage: { getItem: () => null, setItem: () => {} } };
const sandbox = { window, console, Date, Set, Map, Math, JSON, Object, Array, String, Number, Boolean, RegExp };

for (const file of ["domain/height-math.js", "core/project-config.js", "domain/pole-resolution.js", "core/state-compat.js", "core/store.js"]) {
  const modulePath = path.join(__dirname, "..", "js", file);
  vm.runInNewContext(fs.readFileSync(modulePath, "utf8"), sandbox, { filename: modulePath });
}

const S = window.AppStore;
S.resetState();
const canonicalOwnerComm = S.createComm("CTL", "20'", "", {
  ownerBase: "CTL",
  rawOwner: "COMMUNICATION > Century Link"
});
assert.equal(canonicalOwnerComm.ownerCanonical, "CTL", "communications should expose a canonical owner alias");
assert.equal(canonicalOwnerComm.ownerImported, "COMMUNICATION > Century Link", "communications should retain the imported owner alias");
assert.equal(canonicalOwnerComm.ownerDisplay, "CTL", "communications should expose a display owner alias");
assert.equal(S.getOwnerCanonical(canonicalOwnerComm), "CTL", "AppStore should centralize canonical owner access");
assert.equal(S.getOwnerImported(canonicalOwnerComm), "COMMUNICATION > Century Link", "AppStore should centralize imported owner access");
assert.equal(S.getOwnerDisplay(canonicalOwnerComm), "CTL", "AppStore should centralize display owner access");
assert.equal(Object.prototype.hasOwnProperty.call(JSON.parse(JSON.stringify(canonicalOwnerComm)), "ownerCanonical"), false,
  "canonical owner aliases must not change the legacy JSON contract");
assert.equal(
  S.createPole({ poleId: "P0" }).poleInsetReason,
  "FAILING_CLEARANCES",
  "Pole Inset must default to failing clearances"
);
const resolutionPole = S.createPole({ poleId: "P-RESOLUTION", poleResolution: "UG" });
assert.equal(resolutionPole.poleResolution, "UG", "poles should expose a canonical resolution name");
resolutionPole.poleResolution = "PCO";
assert.equal(resolutionPole.pcoActive, true, "canonical pole resolution should update the legacy PCO flag");
assert.equal(resolutionPole.ugActive, false, "canonical pole resolution should keep UG and PCO mutually exclusive");
const relationSpan = S.createSpan("RELATION", "P1", "P2", "E", "", {
  spanRelation: "Fore Span",
  importedRelation: "Fore Span"
});
assert.equal(relationSpan.spanRelation, "Fore Span", "spans should expose a canonical relation name");
assert.equal(relationSpan.importedRelation, "Fore Span", "spans should retain their imported relation name");
assert.equal(S.getSpanRelation(relationSpan), "Fore Span", "AppStore should centralize span relation access");
assert.equal(S.getImportedSpanRelation(relationSpan), "Fore Span", "AppStore should centralize imported span relation access");
const canonicalSpanComm = S.createSpanComm({
  spanId: "RELATION",
  poleId: "P1",
  ownerCanonical: "CTL",
  ownerDisplay: "Century Link",
  ownerImported: "COMMUNICATION > Century Link"
});
assert.equal(canonicalSpanComm.owner, "Century Link", "SpanComm should accept the canonical display owner field");
assert.equal(canonicalSpanComm.ownerBase, "CTL", "SpanComm should accept the canonical matching owner field");
assert.equal(canonicalSpanComm.rawOwner, "COMMUNICATION > Century Link", "SpanComm should accept the canonical imported owner field");
S.upsertPole(S.createPole({ poleId: "P1 STEEL" }));
S.upsertPole(S.createPole({ poleId: "P2" }));
S.upsertPole(S.createPole({ poleId: "P3" }));
S.upsertSpan(S.createSpan("S1", "P1 STEEL", "P2", "E", "", { type: "Fore Span" }));
S.upsertSpan(S.createSpan("S2", "P2", "P3", "E", "", { type: "Fore Span" }));
S.upsertSpanComm({ spanId: "S1", poleId: "P1 STEEL", owner: "CATV", wireId: "W1", existingHOA: "20'" });
S.upsertSpanComm({ spanId: "S1", poleId: "P2", owner: "CATV", wireId: "W1", existingHOA: "21'" });
S.upsertSpanComm({ spanId: "S2", poleId: "P2", owner: "CTL", wireId: "W2", existingHOA: "19'" });
S.upsertSpanSide({ spanId: "S1", poleId: "P1 STEEL", proposedHOA: "22'" });
S.addSpanPower({ spanId: "S1", poleId: "P1 STEEL", wireId: "PW1", midspan: "25'" });

const result = S.removePole("P1 STEEL");
assert.deepEqual(Array.from(result.spanIds), ["S1"], "only spans connected to the deleted pole should be removed");
assert.equal(S.getPole("P1 STEEL"), null, "the selected pole should be deleted");
assert.equal(S.getSpan("S1"), null, "a connected span should be deleted");
assert.ok(S.getSpan("S2"), "an unrelated span should remain");
assert.equal(S.getSpanCommsForSpan("S1").length, 0, "both endpoints of a deleted span should be removed");
assert.equal(S.getSpanCommsForSpan("S2").length, 1, "unrelated comm rows should remain");
assert.deepEqual(Array.from(S.getState().deletedPoleIds), ["P1"], "the canonical pole identity should be retained for Update Data");

const saved = JSON.parse(JSON.stringify(S.getState()));
saved.ui.hiddenPoleIds = ["P2", "P2", "missing"];
saved.poles["Unknown-S1"] = S.createPole({ poleId: "Unknown-S1", isGenerated: true });
S.setState(saved);
assert.deepEqual(Array.from(S.getState().ui.hiddenPoleIds), ["P2", "Unknown-S1"], "saved poles and new Unknown poles should be hidden");
assert.deepEqual(Array.from(S.getState().deletedPoleIds), ["P1"], "deleted pole tombstones should survive Save and Load");

const revealed = JSON.parse(JSON.stringify(S.getState()));
revealed.ui.hiddenPoleIds = revealed.ui.hiddenPoleIds.filter(id => id !== "Unknown-S1");
revealed.ui.revealedUnknownPoleIds = ["Unknown-S1"];
S.setState(revealed);
assert.equal(S.getState().ui.hiddenPoleIds.includes("Unknown-S1"), false, "an Unknown explicitly opened by the user should remain visible after Load");

// Canonical fields must survive the full normalized-state JSON round trip even
// though they remain non-enumerable compatibility aliases in saved files.
S.resetState();
S.upsertPole(S.createPole({ poleId: "ROUNDTRIP-1" }));
S.upsertPole(S.createPole({ poleId: "ROUNDTRIP-2" }));
S.upsertSpan(S.createSpan("ROUNDTRIP-SPAN", "ROUNDTRIP-1", "ROUNDTRIP-2", "E", "", {
  spanRelation: "Back Span",
  importedRelation: "Back Span"
}));
S.upsertSpanComm(S.createSpanComm({
  spanId: "ROUNDTRIP-SPAN",
  poleId: "ROUNDTRIP-1",
  ownerCanonical: "CTL",
  ownerDisplay: "Century Link",
  ownerImported: "COMMUNICATION > Century Link"
}));
const roundTrip = JSON.parse(JSON.stringify(S.getState()));
S.setState(roundTrip);
const restoredSpan = S.getSpan("ROUNDTRIP-SPAN");
const restoredComm = S.getSpanComm("ROUNDTRIP-SPAN", "ROUNDTRIP-1", "CTL", "");
assert.equal(S.getSpanRelation(restoredSpan), "Back Span", "span relation should survive Save/Load");
assert.equal(S.getOwnerCanonical(restoredComm), "CTL", "canonical owner should survive Save/Load");
assert.equal(S.getOwnerImported(restoredComm), "COMMUNICATION > Century Link", "imported owner should survive Save/Load");

console.log("pole-controls.test.js passed");
