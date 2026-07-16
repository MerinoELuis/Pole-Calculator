"use strict";

const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");

const state = {
  settings: { projectProfile: "METRONET", proposedOwner: "Metronet" },
  poles: { P1: { poleId: "P1" }, P2: { poleId: "P2" } },
  spans: {
    F1: { spanId: "F1", fromPole: "P1", toPole: "P2", direction: "E" },
    F2: { spanId: "F2", fromPole: "P2", toPole: "P1", direction: "W" },
    B2: { spanId: "B2", fromPole: "P2", toPole: "P1", direction: "E" }
  },
  spanSides: {
    F1__P1: { spanId: "F1", poleId: "P1", proposedHOA: "23'2\"", finalMidspan: "20'1\"" }
  },
  spanComms: {},
  mr: [{ poleId: "P2", text: "Attach Metronet at HOA 20'." }],
  excelReviewSource: {}
};

function normalizeHeaderName(value) {
  return String(value || "").toLowerCase().replace(/[^a-z0-9]/g, "");
}

function pick(row, names, options = {}) {
  const entries = Object.entries(row || {});
  for (const name of names) {
    const wanted = normalizeHeaderName(name);
    const found = entries.find(([key]) => {
      const actual = normalizeHeaderName(key);
      return options.contains ? actual.includes(wanted) : actual === wanted;
    });
    if (found) return found[1];
  }
  return "";
}

function parseHeight(value) {
  const source = String(value ?? "").trim();
  if (!source) return null;
  if (/^-?\d+(?:\.\d+)?$/.test(source)) return Math.round(Number(source) * 12);
  const match = source.match(/^(-?\d+)\s*'(?:\s*(\d+)\s*\")?$/);
  return match ? (Number(match[1]) * 12) + Number(match[2] || 0) : null;
}

const AppStore = {
  getState: () => state,
  canonicalPoleIdentity(value) {
    const parts = String(value || "").trim().split(/\s+/).filter(Boolean);
    while (parts.length > 1 && /^(STEEL|UG|PCO)$/i.test(parts[parts.length - 1])) parts.pop();
    return parts.join(" ").toUpperCase();
  },
  getSpanCommsForPole: poleId => Object.values(state.spanComms).filter(row => row.poleId === poleId),
  getSpanSidesForPole: poleId => Object.values(state.spanSides).filter(row => row.poleId === poleId),
  getSpan: spanId => state.spans[spanId] || null,
  getOtherPoleId: (span, poleId) => span.fromPole === poleId ? span.toPole : span.fromPole
};

const sandbox = {
  window: {
    AppStore,
    HeightUtils: { parseHeight },
    ExcelImport: {
      pick,
      normalizeHeaderName,
      isPowerWire: row => /^utility\s*>/i.test(String(row.Owner || "")),
      isCommunicationWire: row => !/^utility\s*>/i.test(String(row.Owner || ""))
    }
  },
  console,
  Date,
  Set,
  Map
};

const modulePath = path.join(__dirname, "..", "js", "excel-review.js");
vm.runInNewContext(fs.readFileSync(modulePath, "utf8"), sandbox, { filename: modulePath });
const review = sandbox.window.ExcelReview;

state.excelReviewSource = {
  collection: {
    headers: ["Id", "Sequence", "Year Installed", "Low Power Attachment.display", "MRE Construction Type", "PLA STATUS"],
    rows: [
      { Id: "P1", Sequence: "P1", "Year Installed": 2010, "Low Power Attachment.display": "27'", "MRE Construction Type": "Aerial", "PLA STATUS": "Complete" },
      { Id: "P2", Sequence: "P2", "Year Installed": 2011, "Low Power Attachment.display": "", "Low Power": "27'", "MRE Construction Type": "Aerial", "PLA STATUS": "Complete" }
    ]
  },
  spans: {
    headers: ["Id", "Span Id", "Span Index", "Type", "Linked Collection.Title", "Environment"],
    rows: [
      { Id: "P1", "Span Id": "F1", "Span Index": 1, Type: "Fore Span", "Linked Collection.Title": "P2", Environment: "STREET" },
      { Id: "P1", "Span Id": "O1", "Span Index": 3, Type: "Other", "Linked Collection.Title": "P2", Environment: "STREET" },
      { Id: "P2", "Span Id": "F2", "Span Index": 4, Type: "Fore Span", "Linked Collection.Title": "P1", Environment: "STREET" },
      { Id: "P2", "Span Id": "B2", "Span Index": 5, Type: "Back Span", "Linked Collection.Title": "P1", Environment: "STREET" }
    ]
  },
  spanWires: { headers: [], rows: [] },
  makeReady: {
    headers: ["Id", "Attachment Size", "Attachment Height.display", "Proposed Mid Span.display"],
    rows: [{ Id: "P1", "Attachment Size": "6.6M 24CT Fiber (E/W)", "Attachment Height.display": 23.1667, "Proposed Mid Span.display": 20.0833 }]
  },
  commTransfers: { headers: [], rows: [] }
};

let output = review.runReview();
const p1 = review.reviewPole("P1");
const p2 = review.reviewPole("P2");
assert.equal(p1.hoaStatus, "WARNING", "Other must not count as Fore/Back and zero Back must warn");
assert.ok(p1.checks.some(item => item.code === "MISSING_BACK_SPAN" && item.status === "WARNING"), "zero Back Span must create a warning, not an error");
assert.equal(p1.finalStatus, "PASS", "decimal feet must match feet/inches");
assert.ok(p2.checks.some(item => item.code === "MISSING_LOW_POWER"), "Low Power fallback columns must not satisfy the exact display check");
assert.ok(p2.checks.some(item => item.code === "CALCULATOR_WORK_EXCEL_EMPTY"), "generated MR must count as Calculator final work");
assert.equal(output.summary.total, 2);
assert.deepEqual(output.results.map(item => item.poleId), ["P1", "P2"], "review poles must stay in natural sequence order regardless of severity");

state.excelReviewSource.spans.rows.find(row => row["Span Id"] === "O1")["Linked Collection.Title"] = "P14-NOT-IN-COLLECTION";
output = review.runReview();
assert.equal(
  review.reviewPole("P1").checks.some(item => item.code === "UNKNOWN_LINKED_COLLECTION" && item.details.some(detail => /Span Id: O1/.test(detail))),
  false,
  "Other linked to an unknown collection must not create a review warning"
);

state.excelReviewSource.spans.rows.find(row => row["Span Id"] === "O1")["Linked Collection.Title"] = "";
output = review.runReview();
assert.equal(
  review.reviewPole("P1").checks.some(item => item.code === "MISSING_LINKED_COLLECTION_TITLE" && item.details.some(detail => /Span Id: O1/.test(detail))),
  false,
  "Other without a linked pole must not create a review warning"
);

state.excelReviewSource.spans.rows.push({
  Id: "P1", "Span Id": "B-EXTRA-1", "Span Index": 6, Type: "Back Span", "Linked Collection.Title": "P2", Environment: "STREET"
}, {
  Id: "P1", "Span Id": "B-EXTRA-2", "Span Index": 7, Type: "Back Span", "Linked Collection.Title": "P2", Environment: "STREET"
});
output = review.runReview();
assert.ok(review.reviewPole("P1").checks.some(item => item.code === "MULTIPLE_BACK_SPANS"), "more than one Back Span must error");

state.settings.projectProfile = "INTEC";
state.poles = { P3: { poleId: "P3" } };
state.spans = {};
state.spanSides = {};
state.mr = [];
state.excelReviewSource = {
  collection: { headers: ["Id"], rows: [{ Id: "P3", Sequence: "P3", "Year Installed": 2010, "Low Power Attachment.display": "25'" }] },
  spans: { headers: [], rows: [] },
  spanWires: {
    headers: ["Id", "Owner", "Size", "Construction", "Insulator"],
    rows: [
      { Id: "P3", Owner: "COMMUNICATION > Unknown Network", Size: "Fiber", Construction: " DAVIT ", Insulator: "Hook" },
      { Id: "P3", Owner: "UTILITY > Other", Size: "Primary", Construction: "ON_POLE", Insulator: "Spool 3in" }
    ]
  },
  makeReady: { headers: [], rows: [] },
  commTransfers: { headers: [], rows: [] }
};

output = review.runReview();
const codes = new Set(output.results[0].checks.map(item => item.code));
["DAVIT_NOT_ALLOWED", "UNKNOWN_COMM_OWNER", "INVALID_COMM_INSULATOR", "INVALID_POWER_OWNER", "INVALID_PRIMARY_INSULATOR"].forEach(code => {
  assert.ok(codes.has(code), `expected INTEC check ${code}`);
});

state.poles = { PUG: { poleId: "PUG", ugActive: true } };
state.spans = {};
state.spanSides = {};
state.spanComms = {};
state.mr = [{
  poleId: "PUG",
  text: "Unable to attach due to (reasoning).\nRed tag\nInability to place ANC\nTDU replace required\nExisting neutral / multiplex above 26'9\"\nPCO neutral / multiplex exceeds 26'9\""
}];
state.excelReviewSource = {
  collection: {
    headers: ["Id", "Sequence", "Year Installed", "Low Power Attachment.display", "MRE Construction Type", "PLA STATUS"],
    rows: [{ Id: "PUG", Sequence: "PUG", "Year Installed": 2010, "Low Power Attachment.display": "27'", "MRE Construction Type": "Underground", "PLA STATUS": "Complete" }]
  },
  spans: { headers: [], rows: [] },
  spanWires: { headers: [], rows: [] },
  makeReady: {
    headers: ["Id", "Make Ready Notes"],
    rows: [{ Id: "PUG", "Make Ready Notes": "Unable to attach due to proposed pole overloaded." }]
  },
  commTransfers: { headers: [], rows: [] }
};

output = review.runReview();
let ugChecks = review.reviewPole("PUG").checks.filter(item => item.section === "Make Ready");
assert.ok(ugChecks.some(item => item.code === "VALID_UG_INSTRUCTION" && item.status === "PASS"), "a specific unable-to-attach reason must be valid UG MR");
assert.equal(ugChecks.filter(item => item.status === "ERROR").length, 0, "valid UG MR must not create template-line errors");

state.excelReviewSource.makeReady.rows[0]["Make Ready Notes"] = "Unable to attach due to (reasoning).";
output = review.runReview();
ugChecks = review.reviewPole("PUG").checks.filter(item => item.section === "Make Ready" && item.status === "ERROR");
assert.equal(ugChecks.length, 1, "invalid UG MR must create one consolidated error");
assert.equal(ugChecks[0].code, "MISSING_UG_INSTRUCTION");

state.poles = { PNOTE: { poleId: "PNOTE" } };
state.spans = {};
state.spanSides = {};
state.spanComms = {
  T1: { poleId: "PNOTE", spanId: "", owner: "Century Link Communications", ownerBase: "Century Link Communications", rawOwner: "COMMUNICATION > Century Link Communications", existingHOA: "18'4\"", transferToNewPole: true }
};
state.mr = [{ poleId: "PNOTE", text: "At HOA 18'10\" lower CATV to HOA 16'8\".\nTransfer CTL to new pole at HOA 18'4\"." }];
state.excelReviewSource = {
  collection: {
    headers: ["Id", "Sequence", "Year Installed", "Low Power Attachment.display", "MRE Construction Type", "PLA STATUS"],
    rows: [{ Id: "PNOTE", Sequence: "PNOTE", "Year Installed": 2010, "Low Power Attachment.display": "27'", "MRE Construction Type": "Aerial", "PLA STATUS": "Complete" }]
  },
  spans: { headers: [], rows: [] },
  spanWires: { headers: [], rows: [] },
  makeReady: {
    headers: ["Id", "Make Ready Notes"],
    rows: [
      { Id: "PNOTE", "Make Ready Notes": "At HOA 18'10\" lower CATV to HOA 16'8\".\nTransfer CTL to new pole at HOA 18'4\".\nProposed slack span S.\nPl new ANC 19' S and pl new DG at HOA 22'2\"." },
      { Id: "PNOTE", "Make Ready Notes": "At HOA 18'10\" lower CATV to HOA 16'8\"." }
    ]
  },
  commTransfers: {
    headers: ["Id", "Owner", "Height.display"],
    rows: [{ Id: "PNOTE", Owner: "TELCO", "Height.display": "18'4\"" }]
  }
};

output = review.runReview();
const noteChecks = review.reviewPole("PNOTE").checks;
assert.equal(noteChecks.some(item => item.code === "ADDITIONAL_MR_INSTRUCTION"), false, "duplicate and model-only MR instructions must not warn");
assert.equal(noteChecks.some(item => ["MISSING_COMM_TRANSFER", "TRANSFER_HEIGHT_MISMATCH", "UNEXPECTED_COMM_TRANSFER"].includes(item.code)), false, "TELCO must match CenturyLink/CTL transfers");

state.poles = { PMISMATCH: { poleId: "PMISMATCH" } };
state.spans = {};
state.spanSides = {};
state.spanComms = {};
state.excelReviewIgnoredChecks = {};
state.mr = [{ poleId: "PMISMATCH", text: "Backspan to go UG S due to existing pole overloaded.\nPl riser S at HOA 18'." }];
state.excelReviewSource = {
  collection: {
    headers: ["Id", "Sequence", "Year Installed", "Low Power Attachment.display", "MRE Construction Type", "PLA STATUS"],
    rows: [{ Id: "PMISMATCH", Sequence: "PMISMATCH", "Year Installed": 2010, "Low Power Attachment.display": "27'", "MRE Construction Type": "Aerial", "PLA STATUS": "Complete" }]
  },
  spans: {
    headers: ["Id", "Span Id", "Span Index", "Type", "Linked Collection.Title", "Environment"],
    rows: [
      { Id: "PMISMATCH", "Span Id": "PM-F", "Span Index": 1, Type: "Fore Span", "Linked Collection.Title": "PMISMATCH", Environment: "STREET" },
      { Id: "PMISMATCH", "Span Id": "PM-B", "Span Index": 2, Type: "Back Span", "Linked Collection.Title": "PMISMATCH", Environment: "STREET" }
    ]
  },
  spanWires: { headers: [], rows: [] },
  makeReady: {
    headers: ["Id", "Make Ready Notes"],
    rows: [{ Id: "PMISMATCH", "Make Ready Notes": "Backspan to go UG SE due to existing pole overloaded.\nPl riser S at HOA 18'8\"." }]
  },
  commTransfers: { headers: [], rows: [] }
};

output = review.runReview();
const mismatchChecks = review.reviewPole("PMISMATCH").checks.filter(item => item.section === "Make Ready");
assert.equal(mismatchChecks.filter(item => item.code === "MR_INSTRUCTION_MISMATCH").length, 2, "UG direction and riser height must remain separate paired mismatches");
assert.equal(mismatchChecks.some(item => ["MISSING_MR_INSTRUCTION", "ADDITIONAL_MR_INSTRUCTION"].includes(item.code)), false, "paired MR mismatches must not be duplicated as missing/additional results");

mismatchChecks.filter(item => item.code === "MR_INSTRUCTION_MISMATCH").forEach(item => review.setCheckIgnored(item.ignoreKey, true));
const ignoredResult = review.reviewPole("PMISMATCH");
assert.equal(ignoredResult.checks.filter(item => item.code === "MR_INSTRUCTION_MISMATCH" && item.ignored).length, 2, "ignored findings must remain visible in review details");
assert.equal(ignoredResult.finalStatus, "PASS", "ignored findings must not affect the phase status");
assert.equal(review.getSummary().errors, 0, "ignored findings must not be counted in the review summary");
const restoredKey = ignoredResult.checks.find(item => item.code === "MR_INSTRUCTION_MISMATCH").ignoreKey;
review.setCheckIgnored(restoredKey, false);
assert.equal(review.reviewPole("PMISMATCH").finalStatus, "ERROR", "restored findings must affect review status again");

console.log("Excel Review tests passed.");
