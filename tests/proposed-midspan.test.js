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

for (const file of ["height-utils.js", "project-config.js", "state.js", "calculations.js"]) {
  const modulePath = path.join(__dirname, "..", "js", file);
  vm.runInNewContext(fs.readFileSync(modulePath, "utf8"), sandbox, { filename: modulePath });
}

const S = window.AppStore;
const C = window.Calculations;

function seedSpan(profile, spanId, lengthDisplay, proposedHOA) {
  S.resetState();
  S.applyProjectProfile(profile);
  S.upsertPole(S.createPole({ poleId: "P1", lowPower: "35'" }));
  S.upsertPole(S.createPole({ poleId: "P2", lowPower: "35'" }));
  S.upsertSpan(S.createSpan(spanId, "P1", "P2", "E", "", {
    type: "Fore Span",
    rawType: "Fore Span",
    lengthDisplay
  }));
  S.upsertSpanSide(S.createSpanSide({ spanId, poleId: "P1", proposedHOA }));
}

seedSpan("INTEC", "WITH-MS", "150'", "22'");
S.upsertSpanComm(S.createSpanComm({
  spanId: "WITH-MS",
  poleId: "P1",
  owner: "COMMUNICATION > CATV",
  existingHOA: "20'",
  midspan: "17'"
}));
S.upsertSpan(S.createSpan("UNRELATED", "P1", "P3", "N", "", { type: "Other", rawType: "Other", lengthDisplay: "100'" }));
S.upsertSpanComm(S.createSpanComm({
  spanId: "UNRELATED",
  poleId: "P1",
  owner: "COMMUNICATION > Fiber",
  existingHOA: "30'",
  midspan: "29'"
}));
let side = S.getSpanSide("WITH-MS", "P1");
let span = S.getSpan("WITH-MS");
assert.equal(C.calculateProposedMidspanBase(side, span), null, "INTEC must require a manually entered O-CALC MS");
C.calculateSpanSideMidspan("WITH-MS", "P1");
assert.equal(S.getSpanSide("WITH-MS", "P1").msProposed, "", "INTEC must not derive MS Proposed from comm midspans or span length");
S.upsertSpanSide({ ...S.getSpanSide("WITH-MS", "P1"), ocalcMS: "17.5" });
C.calculateSpanSideMidspan("WITH-MS", "P1");
assert.equal(S.getSpanSide("WITH-MS", "P1").msProposed, "17'6\"", "INTEC must convert a manually entered decimal O-CALC MS for display");

S.upsertPole(S.createPole({ poleId: "P3", lowPower: "35'" }));
S.upsertPole(S.createPole({ poleId: "P4", lowPower: "35'" }));
S.upsertSpan(S.createSpan("ENV-ONLY", "P3", "P4", "E", "", {
  type: "Fore Span",
  rawType: "Fore Span",
  lengthDisplay: "150'"
}));
S.upsertSpanSide(S.createSpanSide({ spanId: "ENV-ONLY", poleId: "P3", ocalcMS: "11.93" }));
C.calculateSpanSideMidspan("ENV-ONLY", "P3");
assert.equal(
  S.getSpanSide("ENV-ONLY", "P3").finalMidspan,
  "11'11\"",
  "Environment clearance must not raise Proposed MS automatically"
);
assert.equal(
  S.getSpanSide("ENV-ONLY", "P3").clearanceMSStatus,
  "PROBLEM",
  "An O-CALC MS below the environment minimum must remain a visible issue"
);
assert.match(
  S.getSpanSide("ENV-ONLY", "P3").clearanceMSMessage,
  /Environment: 11'11\" < 15'6\"/,
  "The environment shortfall must explain why Proposed MS is flagged"
);

S.upsertSpan(S.createSpan("POWER-ONLY", "P1", "P2", "E", "", {
  type: "Fore Span",
  rawType: "Fore Span",
  lengthDisplay: "147'4\""
}));
S.addSpanPower(S.createSpanPower({
  spanId: "POWER-ONLY",
  poleId: "P1",
  wireId: "NEUTRAL-1",
  size: "Neutral",
  midspan: "25'10\""
}));
assert.equal(C.spanHasRealMidspan("POWER-ONLY"), true, "an INTEC Fore Span with only a Power midspan must be eligible for Proposed by Span");

S.upsertSpan(S.createSpan("EMPTY-FORE", "P1", "P2", "E", "", {
  type: "Fore Span",
  rawType: "Fore Span",
  lengthDisplay: "140'"
}));
assert.equal(C.spanHasRealMidspan("EMPTY-FORE"), false, "a truly empty Fore Span must remain hidden under the INTEC profile");

S.upsertSpan(S.createSpan("OTHER-WITH-MS", "P1", "P4", "S", "", {
  type: "Other",
  rawType: "Other",
  lengthDisplay: "125'"
}));
S.upsertSpanComm(S.createSpanComm({
  spanId: "OTHER-WITH-MS",
  poleId: "P1",
  owner: "COMMUNICATION > CATV",
  existingHOA: "20'",
  midspan: "16'8\""
}));
assert.equal(
  C.isSpanEligibleForProposed(S.getSpan("OTHER-WITH-MS"), "P1"),
  false,
  "an Other span must remain reference/manual even when it has real midspan data"
);
assert.equal(
  C.isSpanEligibleForProposed(S.getSpan("OTHER-WITH-MS"), "P4"),
  false,
  "the receiving endpoint must not duplicate Proposed for the same Other span"
);
S.upsertSpan(S.createSpan("OTHER-UNKNOWN-MS", "P1", "Unknown-OTHER-MS", "W", "", {
  type: "Other",
  rawType: "Other",
  lengthDisplay: "30'"
}));
S.upsertSpanComm(S.createSpanComm({
  spanId: "OTHER-UNKNOWN-MS",
  poleId: "P1",
  owner: "COMMUNICATION > CATV",
  existingHOA: "20'",
  midspan: "16'"
}));
assert.equal(
  C.isSpanEligibleForProposed(S.getSpan("OTHER-UNKNOWN-MS"), "P1"),
  false,
  "an Other span ending at a generated Unknown pole must remain reference/manual"
);

seedSpan("INTEC", "DELETABLE-FORE", "150'", "22'");
S.upsertSpanComm(S.createSpanComm({
  spanId: "DELETABLE-FORE",
  poleId: "P1",
  owner: "COMMUNICATION > CATV",
  existingHOA: "20'",
  midspan: "17'"
}));
assert.deepEqual(
  Array.from(C.autoCalcProposedSpansForPole("P1"), span => span.spanId),
  ["DELETABLE-FORE"],
  "an eligible Fore Span must initially appear in Proposed by Span"
);
S.upsertSpanSide({
  ...S.getSpanSide("DELETABLE-FORE", "P1"),
  proposedHOA: "",
  isProposedExcluded: true
});
assert.equal(
  S.getSpanSide("DELETABLE-FORE", "P1").isProposedExcluded,
  true,
  "the delete marker must survive SpanSide normalization"
);
assert.deepEqual(
  Array.from(C.autoCalcProposedSpansForPole("P1"), span => span.spanId),
  [],
  "deleting Proposed must hide its row without deleting the physical Fore Span"
);
assert.ok(S.getSpan("DELETABLE-FORE"), "deleting Proposed must preserve the physical span");

seedSpan("METRONET", "NO-MS", "148'5\"", "22'");
side = S.getSpanSide("NO-MS", "P1");
span = S.getSpan("NO-MS");
assert.equal(C.getEstimatedSagInches(span), 18, "148 feet 5 inches must round to the 150-foot sag bucket");
assert.equal(C.calculateProposedMidspanBase(side, span), 20 * 12 + 6, "MidAm Proposed MS without a comm midspan must subtract 18 inches of estimated sag");
C.calculateSpanSideMidspan("NO-MS", "P1");
assert.equal(S.getSpanSide("NO-MS", "P1").msProposed, "20'6\"", "the estimated MidAm base must be persisted for display and validation");

[
  ["100'", 12],
  ["150'", 18],
  ["200'", 24],
  ["250'", 30]
].forEach(([lengthDisplay, expectedSag]) => {
  assert.equal(C.getEstimatedSagInches({ lengthDisplay }), expectedSag, `${lengthDisplay} must produce ${expectedSag} inches of sag`);
});

S.upsertSpanSide({ ...side, ocalcMS: "19.25" });
side = S.getSpanSide("NO-MS", "P1");
assert.equal(C.calculateProposedMidspanBase(side, span), 19 * 12 + 3, "manual O-CALC MS must override the automatic Proposed calculation");

seedSpan("METRONET", "FORE", "101'6\"", "20'");
S.upsertSpan(S.createSpan("BACK", "P2", "P1", "W", "", {
  type: "Back Span",
  rawType: "Back Span",
  lengthDisplay: "101'6\""
}));
S.upsertSpanComm(S.createSpanComm({
  spanId: "FORE",
  poleId: "P1",
  owner: "COMMUNICATION > Fiber",
  existingHOA: "19'"
}));
S.upsertSpanComm(S.createSpanComm({
  spanId: "BACK",
  poleId: "P2",
  owner: "COMMUNICATION > Fiber",
  existingHOA: "20'2\"",
  midspan: "18'10\""
}));
side = S.getSpanSide("FORE", "P1");
span = S.getSpan("FORE");
assert.equal(C.calculateProposedMidspanBase(side, span), 19 * 12 + 10, "MidAm must use the reciprocal same-connection midspan before falling back to span sag");

seedSpan("INTEC", "TRANSFER-BOLT", "100'", "18'6\"");
S.upsertSpanComm(S.createSpanComm({
  spanId: "TRANSFER-BOLT",
  poleId: "P1",
  owner: "COMMUNICATION > CATV",
  existingHOA: "18'8\"",
  existingHOAChange: "17'6\"",
  transferToNewPole: true
}));
const transferSide = S.getSpanSide("TRANSFER-BOLT", "P1");
assert.equal(
  C.evaluateProposedPoleClearance(transferSide).ok,
  true,
  "a transfer must not reserve the old Existing HOA bolt point on the destination pole"
);
const transferComm = S.getSpanCommsForPole("P1")[0];
S.upsertSpanComm({ ...transferComm, transferToNewPole: false });
assert.equal(
  C.evaluateProposedPoleClearance(transferSide).ok,
  false,
  "the old Existing HOA must retain Bolt-bolt validation when the comm is not transferred"
);

seedSpan("INTEC", "SERVICE-DROP-BOLT", "100'", "");
const regularSameOwner = S.createSpanComm({
  spanId: "SERVICE-DROP-BOLT",
  poleId: "P1",
  owner: "COMMUNICATION > Century Link Communications",
  wireId: "REGULAR-WIRE",
  existingHOA: "18'10\""
});
const serviceDropSameOwner = S.createSpanComm({
  spanId: "SERVICE-DROP-BOLT",
  poleId: "P1",
  owner: "COMMUNICATION > Century Link Communications",
  wireId: "SERVICE-WIRE",
  existingHOA: "18'8\"",
  serviceDrop: true
});
regularSameOwner.midspan = "15'6\"";
serviceDropSameOwner.midspan = "15'4\"";
S.upsertSpanComm(regularSameOwner);
S.upsertSpanComm(serviceDropSameOwner);
assert.doesNotMatch(
  C.evaluateCommFlagging(regularSameOwner, "").flaggingMessage,
  /bolt-bolt/i,
  "a regular comm must not receive Bolt-bolt flagging against a same-owner Service Drop 2 inches away"
);
assert.doesNotMatch(
  C.evaluateCommFlagging(serviceDropSameOwner, "").flaggingMessage,
  /bolt-bolt/i,
  "a Service Drop must not receive Bolt-bolt flagging 2 inches from a same-owner comm"
);
assert.doesNotMatch(
  C.evaluateCommFlagging(serviceDropSameOwner, "").flaggingMessage,
  /Comm-comm MS/i,
  "a Service Drop must not receive Midspan comm-comm flagging"
);

seedSpan("INTEC", "SAME-OWNER-MIDSPAN", "100'", "");
const regularMidspanA = S.createSpanComm({
  spanId: "SAME-OWNER-MIDSPAN",
  poleId: "P1",
  owner: "COMMUNICATION > Century Link Communications",
  wireId: "REGULAR-A",
  existingHOA: "18'10\"",
  midspan: "15'6\""
});
const regularMidspanB = S.createSpanComm({
  spanId: "SAME-OWNER-MIDSPAN",
  poleId: "P1",
  owner: "COMMUNICATION > Century Link Communications",
  wireId: "REGULAR-B",
  existingHOA: "18'4\"",
  midspan: "15'3\""
});
S.upsertSpanComm(regularMidspanA);
S.upsertSpanComm(regularMidspanB);
assert.match(
  C.evaluateCommFlagging(regularMidspanA, "").flaggingMessage,
  /Comm-comm MS.*separation 3".*minimum 4"/i,
  "regular same-owner comms must receive Midspan clearance flagging"
);

seedSpan("INTEC", "DG-STILL-BOLT", "100'", "");
const regularWithDg = S.createSpanComm({
  spanId: "DG-STILL-BOLT",
  poleId: "P1",
  owner: "COMMUNICATION > Century Link Communications",
  wireId: "DG-WIRE",
  existingHOA: "18'10\"",
  downGuy: true
});
const regularWithoutDrop = S.createSpanComm({
  spanId: "DG-STILL-BOLT",
  poleId: "P1",
  owner: "COMMUNICATION > Century Link Communications",
  wireId: "REGULAR-WIRE",
  existingHOA: "18'8\""
});
S.upsertSpanComm(regularWithDg);
S.upsertSpanComm(regularWithoutDrop);
assert.match(
  C.evaluateCommFlagging(regularWithDg, "").flaggingMessage,
  /Pole bolt-bolt: 2"/i,
  "DG alone must not receive the Service Drop Bolt-bolt exception"
);

seedSpan("INTEC", "PROPOSED-SERVICE-BOLT", "100'", "18'10\"");
S.upsertSpanComm(S.createSpanComm({
  spanId: "PROPOSED-SERVICE-BOLT",
  poleId: "P1",
  owner: "COMMUNICATION > Century Link Communications",
  wireId: "SERVICE-WIRE",
  existingHOA: "18'8\"",
  serviceDrop: true
}));
const proposedAgainstServiceDrop = C.evaluateProposedPoleClearance(S.getSpanSide("PROPOSED-SERVICE-BOLT", "P1"));
assert.doesNotMatch(
  proposedAgainstServiceDrop.message,
  /Bolt-bolt/i,
  "a Service Drop Existing HOA must not reserve a Bolt-bolt exclusion point for Proposed"
);
assert.match(
  proposedAgainstServiceDrop.message,
  /Comm-comm/i,
  "the Service Drop Bolt-bolt exception must not remove Proposed Comm-comm clearance"
);

console.log("Proposed midspan fallback tests passed.");
