"use strict";

const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");

const appPath = path.join(__dirname, "..", "js", "app.js");
const appSource = fs.readFileSync(appPath, "utf8").replace(
  'document.addEventListener("DOMContentLoaded", init);',
  "global.__mergeImportedUpdate = mergeImportedUpdate;"
);
const sandbox = {
  window: {
    AppStore: {
      canonicalPoleIdentity(value) {
        const parts = String(value || "").trim().split(/\s+/).filter(Boolean);
        while (parts.length > 1 && /^(STEEL|UG|PCO)$/i.test(parts[parts.length - 1])) parts.pop();
        return parts.join(" ").toUpperCase();
      },
      keyForSpanSide: (spanId, poleId) => `${spanId}__${poleId}`,
      keyForSpanComm: (spanId, poleId, owner, wireId) => `${spanId}__${poleId}__${owner}__${wireId || ""}`
    },
    HeightUtils: {},
    ExcelImport: { recalculatePoleClassCheck: row => row }
  },
  console,
  Date,
  Map,
  Set,
  JSON
};
vm.runInNewContext(appSource, sandbox, { filename: appPath });
const merge = sandbox.window.__mergeImportedUpdate;

const previous = {
  poles: { P1: { poleId: "P1", lowPower: "30'", poleHeight: "40'", notes: "Saved note", ugActive: false, pcoActive: false } },
  spans: {
    S1: { spanId: "S1", fromPole: "P1", toPole: "P2", lengthDisplay: "100'", environment: "STREET" },
    S2: { spanId: "S2", fromPole: "P1", toPole: "P3", lengthDisplay: "80'", environment: "ALLEY" }
  },
  spanSides: { S1__P1: { spanId: "S1", poleId: "P1", proposedHOA: "22'", proposedHOAChange: "", ocalcMS: "" } },
  spanComms: {
    C1: { spanId: "S1", poleId: "P1", owner: "CATV", ownerBase: "CATV", wireId: "W1", existingHOA: "20'", midspan: "16'", size: "Old size" },
    C2: { spanId: "S2", poleId: "P1", owner: "CTL", ownerBase: "CTL", wireId: "W2", existingHOA: "19'", midspan: "15'", size: "Telco" }
  },
  spanPower: { PW1: { spanId: "S1", poleId: "P1", owner: "APS", wireId: "PW1", size: "Primary", attachmentHeight: "35'", midspan: "29'" } },
  makeReadyReferences: [],
  poleClassChecks: [],
  settings: { projectProfile: "INTEC", fiberSizes: {} },
  ui: {},
  excelReviewSource: { collection: { rows: [{ Id: "OLD" }] } }
};

const imported = {
  poles: { P1: { poleId: "P1", lowPower: "", poleHeight: "45'", notes: "", ugActive: false, pcoActive: false } },
  spans: { S1: { spanId: "S1", fromPole: "P1", toPole: "P2", lengthDisplay: "", environment: "ALLEY" } },
  spanSides: { S1__P1: { spanId: "S1", poleId: "P1", proposedHOA: "", proposedHOAChange: "", ocalcMS: "" } },
  spanComms: { C1: { spanId: "S1", poleId: "P1", owner: "CATV", ownerBase: "CATV", wireId: "W1", existingHOA: "", midspan: "", size: "New size" } },
  spanPower: { PW1: { spanId: "S1", poleId: "P1", owner: "APS", wireId: "PW1", size: "Primary", attachmentHeight: "", midspan: "" } },
  makeReadyReferences: [],
  poleClassChecks: [],
  settings: { projectProfile: "INTEC", fiberSizes: {} },
  ui: {},
  excelReviewSource: { collection: { rows: [{ Id: "NEW", "Low Power Attachment.display": "" }] } }
};

const result = merge(previous, imported);
assert.equal(result.poles.P1.lowPower, "30'", "blank updated pole value must retain the prior value");
assert.equal(result.poles.P1.poleHeight, "45'", "non-empty updated pole value must win");
assert.equal(result.spans.S1.lengthDisplay, "100'", "blank updated span value must retain the prior value");
assert.equal(result.spans.S1.environment, "ALLEY", "non-empty updated span value must win");
assert.equal(result.spans.S2.lengthDisplay, "80'", "a span omitted by Update Data must remain when it owns retained calculation data");
assert.equal(result.spanSides.S1__P1.proposedHOA, "22'", "saved Proposed must be retained");
assert.equal(result.spanComms.C1.existingHOA, "20'", "blank Existing HOA must retain the prior value");
assert.equal(result.spanComms.C1.midspan, "16'", "blank imported midspan must retain the prior value");
assert.equal(result.spanComms.C1.size, "New size", "new non-empty comm data must win");
assert.equal(
  Object.values(result.spanComms).find(row => row.spanId === "S2" && row.owner === "CTL").existingHOA,
  "19'",
  "an omitted comm with baseline HOA must remain available for calculations"
);
assert.equal(result.spanPower.PW1.midspan, "29'", "blank power midspan must retain the prior value");
assert.equal(result.excelReviewSource.collection.rows[0].Id, "NEW", "raw review source must remain the new workbook snapshot");
assert.ok(result.updateDiagnostics.blankValuesPreserved > 0);

const aliasPrevious = {
  poles: {
    "P01-LX339927 STEEL UG": { poleId: "P01-LX339927 STEEL UG", ugActive: true, pcoActive: false, notes: "UG work" }
  },
  spans: { SA: { spanId: "SA", fromPole: "P01-LX339927 STEEL UG", toPole: "P02-X339926 STEEL" } },
  spanSides: {
    "SA__P01-LX339927 STEEL UG": { spanId: "SA", poleId: "P01-LX339927 STEEL UG", proposedHOA: "22'", isManualProposed: true }
  },
  spanComms: {}, spanPower: {}, makeReadyReferences: [], poleClassChecks: [], settings: {}, ui: {}
};
const aliasImported = {
  poles: {
    "P01-LX339927": { poleId: "P01-LX339927", ugActive: false, pcoActive: false },
    "P02-X339926 STEEL": { poleId: "P02-X339926 STEEL", ugActive: false, pcoActive: false }
  },
  spans: { SA: { spanId: "SA", fromPole: "P01-LX339927", toPole: "P02-X339926 STEEL" } },
  spanSides: { "SA__P01-LX339927": { spanId: "SA", poleId: "P01-LX339927", proposedHOA: "" } },
  spanComms: {}, spanPower: {}, makeReadyReferences: [], poleClassChecks: [], settings: {}, ui: {}, excelReviewSource: {}
};
const aliasResult = merge(aliasPrevious, aliasImported);
assert.deepEqual(Object.keys(aliasResult.poles).sort(), ["P01-LX339927", "P02-X339926 STEEL"], "UG/STEEL aliases must not create a second pole");
assert.equal(aliasResult.poles["P01-LX339927"].ugActive, true, "UG state must move to the Collection pole identity");
assert.equal(aliasResult.spanSides["SA__P01-LX339927"].proposedHOA, "22'", "manual Proposed must follow the canonical pole identity");

const endpointPrevious = {
  poles: {
    P08: { poleId: "P08" },
    P09: { poleId: "P09" },
    P10: { poleId: "P10" }
  },
  spans: {
    BACK: { spanId: "BACK", fromPole: "P09", toPole: "P08", type: "Back Span" },
    FORE: { spanId: "FORE", fromPole: "P09", toPole: "P10", type: "Fore Span" }
  },
  spanSides: {},
  spanComms: {
    BACK_CATV: { spanId: "BACK", poleId: "P09", owner: "CATV", ownerBase: "CATV", wireId: "CATV", existingHOA: "20'6\"", midspan: "" },
    FORE_CATV: { spanId: "FORE", poleId: "P09", owner: "CATV", ownerBase: "CATV", wireId: "CATV", existingHOA: "20'6\"", midspan: "19'7\"" },
    BACK_CTL: { spanId: "BACK", poleId: "P09", owner: "CenturyLink", ownerBase: "CenturyLink", wireId: "CTL", existingHOA: "15'6\"", midspan: "" },
    FORE_CTL: { spanId: "FORE", poleId: "P09", owner: "CenturyLink", ownerBase: "CenturyLink", wireId: "CTL", existingHOA: "15'6\"", existingHOAChange: "18'6\"", midspan: "17'7\"" }
  },
  spanPower: {}, makeReadyReferences: [], poleClassChecks: [], settings: {}, ui: {}
};
const endpointImported = {
  poles: {
    P08: { poleId: "P08" },
    P09: { poleId: "P09" },
    P10: { poleId: "P10" }
  },
  spans: {
    INCOMING: { spanId: "INCOMING", fromPole: "P08", toPole: "P09", type: "Fore Span" },
    FORE: { spanId: "FORE", fromPole: "P09", toPole: "P10", type: "Fore Span" }
  },
  spanSides: {},
  spanComms: {
    PLACEHOLDER_CATV: { spanId: "INCOMING", poleId: "P09", owner: "CATV", ownerBase: "CATV", wireId: "CATV", existingHOA: "", midspan: "", isEndpointPlaceholder: true },
    PLACEHOLDER_CTL: { spanId: "INCOMING", poleId: "P09", owner: "CenturyLink", ownerBase: "CenturyLink", wireId: "CTL", existingHOA: "", midspan: "", isEndpointPlaceholder: true }
  },
  spanPower: {}, makeReadyReferences: [], poleClassChecks: [], settings: {}, ui: {}, excelReviewSource: {}
};
const endpointResult = merge(endpointPrevious, endpointImported);
const p09Rows = Object.values(endpointResult.spanComms).filter(row => row.poleId === "P09");
assert.equal(p09Rows.length, 4, "Update Data must not append blank inverse-span endpoint rows");
assert.equal(p09Rows.some(row => row.isEndpointPlaceholder), false, "superseded endpoint placeholders must be discarded");
assert.equal(p09Rows.find(row => row.spanId === "FORE" && row.owner === "CATV").midspan, "19'7\"", "prior CATV midspan must remain a calculation input");
assert.equal(endpointResult.updateDiagnostics.endpointPlaceholdersDiscarded, 2);
assert.equal(endpointResult.updateDiagnostics.baselineCommRowsPreserved, 3);

console.log("Update Data non-destructive merge tests passed.");
