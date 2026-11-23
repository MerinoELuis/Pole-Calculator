"use strict";

const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");

const appPath = path.join(__dirname, "..", "js", "ui", "app.js");
const reconcilerPath = path.join(__dirname, "..", "js", "import", "update-data-reconciler.js");
const dialogsPath = path.join(__dirname, "..", "js", "ui", "dialogs.js");
const commPairingDialogPath = path.join(__dirname, "..", "js", "ui", "comm-pairing-dialog.js");
const persistencePath = path.join(__dirname, "..", "js", "core", "persistence.js");
const reconcilerSource = fs.readFileSync(reconcilerPath, "utf8");
const dialogsSource = fs.readFileSync(dialogsPath, "utf8");
const commPairingDialogSource = fs.readFileSync(commPairingDialogPath, "utf8");
const persistenceSource = fs.readFileSync(persistencePath, "utf8");
const appSource = fs.readFileSync(appPath, "utf8").replace(
  'document.addEventListener("DOMContentLoaded", init);',
  "global.__mergeImportedUpdate = mergeImportedUpdate;"
);
const sandbox = {
  window: {
    AppStore: {
      canonicalPoleIdentity(value) {
        const parts = String(value || "").trim().split(/\s+/).filter(Boolean);
        while (parts.length > 1 && /^(STEEL|RECOLLECT|T|UG|PCO)$/i.test(parts[parts.length - 1])) parts.pop();
        return parts.join(" ").toUpperCase();
      },
      getOwnerCanonical: row => row?.ownerCanonical || row?.ownerBase || row?.owner || "",
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
vm.runInNewContext(`${reconcilerSource}\n${dialogsSource}\n${commPairingDialogSource}\n${persistenceSource}\n${appSource}`, sandbox, { filename: appPath });
const merge = sandbox.window.__mergeImportedUpdate;

const previous = {
  fiberAssignments: [{ id: "manual-fiber:save:1", manual: true, poleId: "P1", spanId: "S1", fiberType: "MST", mstPorts: "8", installType: "NEW" }],
  poles: { P1: { poleId: "P1", lowPower: "30'", poleHeight: "40'", notes: "Saved note", ugActive: true, ugReason: "proposed pole overloaded", ugMRText: "Unable to attach due to red tag.\nRed tag", ugRiserDirection: "E", riserActive: true, pcoActive: false, pcoMRText: "Replace pole due to overload.", metadata: { lowPowerBaseline: "29'", plaExisting: { yearInstalled: "2016", windSpeedMph: "115" }, powerEquipment: [{ equipmentId: "EQ-1", category: "TRANSFORMER", dripLoopHeight: "29'", actionActive: true, actionHeight: "31'" }, { equipmentId: "EQ-2", category: "STREETLIGHT", attachmentHeight: "24'", actionActive: false, actionHeight: "", raiseActive: true, raiseHeight: "25'" }] } } },
  spans: {
    S1: { spanId: "S1", fromPole: "P1", toPole: "P2", lengthDisplay: "100'", environment: "STREET" },
    S2: { spanId: "S2", fromPole: "P1", toPole: "P3", lengthDisplay: "80'", environment: "ALLEY" }
  },
  spanSides: { S1__P1: { spanId: "S1", poleId: "P1", proposedHOA: "22'", proposedHOAChange: "", ocalcMS: "" } },
  spanComms: {
    C1: { spanId: "S1", poleId: "P1", owner: "CATV", ownerBase: "CATV", wireId: "W1", existingHOA: "20'", midspan: "16'", size: "Old size" },
    C2: { spanId: "S2", poleId: "P1", owner: "CTL", ownerBase: "CTL", wireId: "W2", existingHOA: "19'", midspan: "15'", size: "Telco" }
  },
  spanPower: {
    PW1: { spanId: "S1", poleId: "P1", owner: "APS", wireId: "PW1", size: "Primary", attachmentHeight: "35'", attachmentHeightChange: "36'", midspan: "29'", calculatedMidspan: "29'6\"" },
    PW2: { spanId: "S2", poleId: "P1", owner: "APS", wireId: "PW2", size: "Secondary", attachmentHeight: "32'", midspan: "27'" }
  },
  makeReadyReferences: [],
  poleClassChecks: [{ poleId: "P2", measuredTip: "31'", expectedLength: 35 }],
  settings: { projectProfile: "INTEC", fiberSizes: {}, plaExistingLoadingProfile: "LIGHT" },
  ui: {},
  excelReviewIgnoredChecks: { "review-known-exception": true },
  excelReviewSource: {
    collection: { headers: ["Id", "Tip.display"], rows: [{ Id: "P2", "Tip.display": "31'" }] },
    spans: { headers: ["Id", "Span Id", "Type"], rows: [{ Id: "P2", "Span Id": "REVIEW-S2", Type: "Fore Span" }] }
  }
};

const imported = {
  poles: { P1: { poleId: "P1", lowPower: "28'", poleHeight: "45'", notes: "", ugActive: false, pcoActive: false, metadata: { powerEquipment: [{ equipmentId: "EQ-1", category: "TRANSFORMER", dripLoopHeight: "28'", actionActive: false, actionHeight: "" }, { equipmentId: "EQ-2", category: "STREETLIGHT", attachmentHeight: "24'", actionActive: false, actionHeight: "", raiseActive: false, raiseHeight: "" }] } } },
  spans: { S1: { spanId: "S1", fromPole: "P1", toPole: "P2", lengthDisplay: "", environment: "ALLEY" } },
  spanSides: { S1__P1: { spanId: "S1", poleId: "P1", proposedHOA: "", proposedHOAChange: "", ocalcMS: "" } },
  spanComms: { C1: { spanId: "S1", poleId: "P1", owner: "CATV", ownerBase: "CATV", wireId: "W1", existingHOA: "", midspan: "", size: "New size" } },
  spanPower: { PW1: { spanId: "S1", poleId: "P1", owner: "APS", wireId: "PW1", size: "Primary", attachmentHeight: "", midspan: "" } },
  makeReadyReferences: [],
  poleClassChecks: [{ poleId: "P1", measuredTip: "36'", expectedLength: 40 }],
  settings: { projectProfile: "INTEC", fiberSizes: {} },
  ui: {},
  excelReviewSource: {
    collection: { headers: ["Id", "Tip.display", "Low Power Attachment.display"], rows: [{ Id: "P1", "Tip.display": "36'", "Low Power Attachment.display": "" }] },
    spans: { headers: ["Id", "Span Id", "Type"], rows: [{ Id: "P1", "Span Id": "REVIEW-S1", Type: "Fore Span" }] }
  }
};

previous.poles.P1.metadata.anchors = [{
  anchorId: "A1", anchorIndex: "1", owner: "UTILITY > APS", leadLength: "18'", bearing: "90°",
  groundlineViolation: true, manualFields: ["groundlineViolation"],
  guys: [{ guyId: "G1", guyIndex: "1", size: "EHS 5/16 DG", owner: "UTILITY > APS", attachmentHeight: "19'", manualFields: ["attachmentHeight"] }]
}];
imported.poles.P1.metadata.anchors = [{
  anchorId: "A1", anchorIndex: "1", owner: "UTILITY > APS", leadLength: "19'", bearing: "180°",
  groundlineViolation: false, manualFields: [],
  guys: [{ guyId: "G1", guyIndex: "1", size: "EHS 3/8 DG", owner: "UTILITY > APS", attachmentHeight: "20'", manualFields: [] }]
}];

const result = merge(previous, imported);
assert.equal(result.poles.P1.lowPower, "28'", "non-empty updated Low Power must become the new imported value");
assert.equal(result.poles.P1.metadata.lowPowerBaseline, "28'", "new imported Low Power must become the equipment-action baseline");
assert.equal(result.poles.P1.metadata.powerEquipment[0].actionActive, true, "Update Data must preserve the equipment action");
assert.equal(result.poles.P1.metadata.powerEquipment[0].actionHeight, "31'", "Update Data must preserve the equipment target HOA");
assert.equal(result.poles.P1.metadata.powerEquipment[1].raiseActive, true, "Update Data must preserve the independent Streetlight Raise action");
assert.equal(result.poles.P1.metadata.powerEquipment[1].raiseHeight, "25'", "Update Data must preserve the Streetlight Raise target HOA");
assert.equal(result.poles.P1.metadata.plaExisting.yearInstalled, "2016", "Update Data must preserve a manually entered PLA Existing year");
assert.equal(result.poles.P1.metadata.plaExisting.windSpeedMph, "115", "Update Data must preserve the per-pole wind input");
assert.equal(result.settings.plaExistingLoadingProfile, "LIGHT", "Update Data must preserve the job-wide principal loading");
assert.equal(result.poles.P1.metadata.anchors[0].leadLength, "19'", "Update Data should refresh unedited imported anchor data");
assert.equal(result.poles.P1.metadata.anchors[0].bearing, "180°", "Update Data should refresh an unedited anchor bearing");
assert.equal(result.poles.P1.metadata.anchors[0].groundlineViolation, true, "Update Data must preserve the manually changed groundline violation toggle");
assert.equal(result.poles.P1.metadata.anchors[0].guys[0].size, "EHS 3/8 DG", "Update Data should refresh unedited Guy size");
assert.equal(result.poles.P1.metadata.anchors[0].guys[0].attachmentHeight, "19'", "Update Data must preserve a manually edited Guy attachment HOA");
assert.equal(result.poles.P1.ugActive, true, "Update Data must preserve the UG resolution");
assert.equal(result.poles.P1.ugReason, "proposed pole overloaded", "Update Data must preserve the editable UG reason");
assert.equal(result.poles.P1.ugMRText, "Unable to attach due to red tag.\nRed tag", "Update Data must preserve the editable UG template");
assert.equal(result.poles.P1.pcoMRText, "Replace pole due to overload.", "Update Data must preserve the editable PCO template");
assert.equal(result.poles.P1.ugRiserDirection, "E", "Update Data must preserve the editable riser direction");
assert.equal(result.poles.P1.riserActive, true, "Update Data must preserve the explicit Riser action");
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
assert.equal(result.spanPower.PW1.attachmentHeightChange, "36'", "Update Data must retain a saved Power HOA Change when the refreshed sheet is blank");
assert.equal(result.spanPower.PW1.calculatedMidspan, "29'6\"", "Update Data must preserve the derived Power Midspan until recalculation");
assert.equal(result.spanPower.PW2.midspan, "27'", "power rows omitted by a partial Update Data must remain");
assert.equal(result.fiberAssignments[0].fiberType, "MST", "Update Data must retain manually configured fiber rows");
assert.equal(result.fiberAssignments[0].mstPorts, "8", "Update Data must retain the manually selected MST port count");
const importedWithoutEquipment = JSON.parse(JSON.stringify(imported));
importedWithoutEquipment.poles.P1.metadata = { powerEquipment: [] };
const noEquipmentResult = merge(previous, importedWithoutEquipment);
assert.equal(noEquipmentResult.poles.P1.metadata.powerEquipment.length, 2, "an update without an Equipment sheet must retain existing power rows");
const importedWithoutAnchors = JSON.parse(JSON.stringify(imported));
delete importedWithoutAnchors.poles.P1.metadata.anchors;
const noAnchorSheetResult = merge(previous, importedWithoutAnchors);
assert.equal(noAnchorSheetResult.poles.P1.metadata.anchors[0].leadLength, "18'", "an update without an Anchor sheet should retain the existing ANC data");
assert.equal(result.excelReviewSource.collection.rows.find(row => row.Id === "P1")["Tip.display"], "36'", "updated Review rows must use the new workbook value");
assert.ok(result.excelReviewSource.collection.rows.some(row => row.Id === "P2"), "partial Update Data must retain other poles in Excel Review");
assert.ok(result.excelReviewSource.spans.rows.some(row => row["Span Id"] === "REVIEW-S2"), "partial Update Data must retain other poles' Review span rows");
assert.ok(result.poleClassChecks.some(row => row.poleId === "P1"), "Pole Type Check must include the updated pole");
assert.ok(result.poleClassChecks.some(row => row.poleId === "P2"), "Pole Type Check must retain poles omitted by a partial update");
assert.equal(result.excelReviewIgnoredChecks["review-known-exception"], true, "Update Data must preserve ignored review findings");
assert.ok(result.updateDiagnostics.blankValuesPreserved > 0);

const spanUGPrevious = JSON.parse(JSON.stringify(previous));
spanUGPrevious.spans.S1.ugActive = true;
spanUGPrevious.spans.S1.ugReason = "cutout clearance";
spanUGPrevious.spans.S1.ugSourcePoleId = "P1";
const spanUGResult = merge(spanUGPrevious, JSON.parse(JSON.stringify(imported)));
assert.equal(spanUGResult.spans.S1.ugActive, true, "Update Data must preserve selected span-level UG work");
assert.equal(spanUGResult.spans.S1.ugReason, "cutout clearance", "Update Data must preserve the reason from the shared UG text box");
assert.equal(spanUGResult.spans.S1.ugSourcePoleId, "P1", "Update Data must preserve which pole owns the shared UG text box");

const previousFiberUpdate = JSON.parse(JSON.stringify(previous));
previousFiberUpdate.makeReadyReferences = Array.from({ length: 10 }, (_, index) => ({
  poleId: `P${index + 1}`,
  makeReadyId: `MR-P${index + 1}`,
  makeReadyIndex: String(index + 1),
  attachmentFiber: "48CT Fiber",
  attachmentSizeRaw: "6.6M 48CT Fiber (N)",
  attachmentDirectionTokens: ["N"],
  attachmentType: "New"
}));
for (let index = 2; index <= 10; index += 1) {
  previousFiberUpdate.poles[`P${index}`] = {
    poleId: `P${index}`,
    metadata: { powerEquipment: [] }
  };
}
const importedFiberUpdate = JSON.parse(JSON.stringify(imported));
for (let index = 2; index <= 12; index += 1) {
  importedFiberUpdate.poles[`P${index}`] = importedFiberUpdate.poles[`P${index}`] || {
    poleId: `P${index}`,
    metadata: { powerEquipment: [] }
  };
}
importedFiberUpdate.makeReadyReferences = [11, 12].map(index => ({
  poleId: `P${index}`,
  makeReadyId: `MR-P${index}`,
  makeReadyIndex: String(index),
  attachmentFiber: "24CT Fiber",
  attachmentSizeRaw: "6.6M 24CT Fiber (E)",
  attachmentDirectionTokens: ["E"],
  attachmentType: "New"
}));
const preservedFiberUpdate = merge(previousFiberUpdate, importedFiberUpdate);
assert.equal(preservedFiberUpdate.makeReadyReferences.length, 12, "adding two poles in a partial Update Data must retain the ten existing fiber rows");
assert.equal(preservedFiberUpdate.makeReadyReferences.filter(row => row.attachmentFiber === "48CT Fiber").length, 10);
assert.equal(preservedFiberUpdate.updateDiagnostics.fiberMakeReadyReferencesPreserved, 10);

const clearedFiberPrevious = JSON.parse(JSON.stringify(previous));
clearedFiberPrevious.makeReadyReferences = [{
  poleId: "P1", makeReadyId: "MR-P1", makeReadyIndex: "1",
  attachmentFiber: "48CT Fiber", attachmentSizeRaw: "6.6M 48CT Fiber (N)",
  attachmentDirectionTokens: ["N"], attachmentType: "New"
}];
const clearedFiberImport = JSON.parse(JSON.stringify(imported));
clearedFiberImport.makeReadyReferences = [{
  poleId: "P1", makeReadyId: "MR-P1", makeReadyIndex: "1",
  attachmentFiber: "", attachmentSizeRaw: "", attachmentDirectionTokens: [], attachmentType: "New"
}];
const clearedFiberResult = merge(clearedFiberPrevious, clearedFiberImport);
assert.equal(clearedFiberResult.makeReadyReferences[0].attachmentFiber, "", "a matching updated Excel row with no fiber explicitly clears the old fiber");
assert.equal(clearedFiberResult.updateDiagnostics.fiberMakeReadyReferencesRemoved, 1);

const unidentifiedFiberPrevious = JSON.parse(JSON.stringify(previous));
unidentifiedFiberPrevious.makeReadyReferences = [{
  poleId: "P1", attachmentFiber: "48CT Fiber", attachmentSizeRaw: "6.6M 48CT Fiber (N)", attachmentDirection: "N"
}];
const unidentifiedFiberImport = JSON.parse(JSON.stringify(imported));
unidentifiedFiberImport.makeReadyReferences = [{ poleId: "P1", attachmentFiber: "", attachmentSizeRaw: "", attachmentDirection: "" }];
const unidentifiedFiberResult = merge(unidentifiedFiberPrevious, unidentifiedFiberImport);
assert.equal(unidentifiedFiberResult.makeReadyReferences[0].attachmentFiber, "48CT Fiber", "a blank row without stable MR identity is not enough evidence to delete a prior fiber");

const manualAncPrevious = JSON.parse(JSON.stringify(previous));
manualAncPrevious.poles.P1.ancActive = true;
manualAncPrevious.poles.P1.ancDistance = "8'";
manualAncPrevious.poles.P1.ancDistanceManual = true;
const refreshedAncImport = JSON.parse(JSON.stringify(imported));
refreshedAncImport.poles.P1.ancActive = true;
refreshedAncImport.poles.P1.ancDistance = "10'";
refreshedAncImport.poles.P1.ancDistanceManual = false;
const ancUpdateResult = merge(manualAncPrevious, refreshedAncImport);
assert.equal(ancUpdateResult.poles.P1.ancDistance, "8'", "Update Data must preserve a manually entered ANC Distance over a newly imported/default distance");
assert.equal(ancUpdateResult.poles.P1.ancDistanceManual, true, "Update Data must retain the manual ANC Distance marker");

const automaticAncPrevious = JSON.parse(JSON.stringify(previous));
automaticAncPrevious.poles.P1.ancActive = true;
automaticAncPrevious.poles.P1.ancDistance = "9'";
automaticAncPrevious.poles.P1.ancDistanceManual = false;
const automaticAncUpdate = merge(automaticAncPrevious, refreshedAncImport);
assert.equal(automaticAncUpdate.poles.P1.ancDistance, "10'", "a non-manual ANC Distance may update from refreshed/default data");

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

const revisionPrevious = {
  poles: {
    "P3-313602": { poleId: "P3-313602" },
    "P4-313603": { poleId: "P4-313603" },
    "P4-313603 RECOLLECT": { poleId: "P4-313603 RECOLLECT" }
  },
  spans: { S34: { spanId: "S34", fromPole: "P3-313602", toPole: "P4-313603 RECOLLECT" } },
  spanSides: {},
  spanComms: {
    REVISION_BASELINE: { spanId: "S34", poleId: "P4-313603 RECOLLECT", owner: "Eagle West LLC", ownerCanonical: "Eagle West LLC", wireId: "WIRE-P34", existingHOA: "22'8\"", existingHOAChange: "", midspan: "21'9\"" },
    OLD_ID_MOVED: { spanId: "S34", poleId: "P4-313603", owner: "Eagle West LLC", ownerCanonical: "Eagle West LLC", wireId: "WIRE-P34", existingHOA: "22'8\"", existingHOAChange: "22'", midspan: "" }
  },
  spanPower: {}, makeReadyReferences: [], poleClassChecks: [], settings: {}, ui: {},
  excelReviewSource: {}
};
const revisionImported = {
  poles: {
    "P3-313602": { poleId: "P3-313602" },
    "P4-313603 RECOLLECT": { poleId: "P4-313603 RECOLLECT" }
  },
  spans: { S34: { spanId: "S34", fromPole: "P3-313602", toPole: "P4-313603 RECOLLECT" } },
  spanSides: {},
  spanComms: {
    NEW_IMPORT: { spanId: "S34", poleId: "P4-313603 RECOLLECT", owner: "Eagle West LLC", ownerCanonical: "Eagle West LLC", wireId: "WIRE-P34", existingHOA: "22'8\"", existingHOAChange: "", midspan: "21'9\"" }
  },
  spanPower: {}, makeReadyReferences: [], poleClassChecks: [], settings: {}, ui: {},
  excelReviewSource: {}
};
const revisionResult = merge(revisionPrevious, revisionImported);
assert.deepEqual(Object.keys(revisionResult.poles).sort(), ["P3-313602", "P4-313603 RECOLLECT"], "Update Data should merge old suffix variants into the latest Collection pole");
assert.equal(Object.values(revisionResult.spanComms).filter(row => row.spanId === "S34" && row.poleId === "P4-313603 RECOLLECT").length, 1, "aliased duplicate endpoint cables must reconcile to one row");
assert.equal(Object.values(revisionResult.spanComms).find(row => row.wireId === "WIRE-P34")?.existingHOAChange, "22'", "Update Data must retain movement from the old suffix variant when merging duplicate pole rows");

const renamedPrevious = JSON.parse(JSON.stringify(previous));
renamedPrevious.poleIdAliases = { P1: "P1-RENAMED" };
renamedPrevious.poles["P1-RENAMED"] = { ...renamedPrevious.poles.P1, poleId: "P1-RENAMED" };
delete renamedPrevious.poles.P1;
Object.values(renamedPrevious.spans).forEach(span => {
  if (span.fromPole === "P1") span.fromPole = "P1-RENAMED";
  if (span.toPole === "P1") span.toPole = "P1-RENAMED";
});
renamedPrevious.spanSides = Object.fromEntries(Object.values(renamedPrevious.spanSides).map(row => {
  if (row.poleId === "P1") row.poleId = "P1-RENAMED";
  return [`${row.spanId}__${row.poleId}`, row];
}));
["spanComms", "spanPower"].forEach(mapName => Object.values(renamedPrevious[mapName]).forEach(row => {
  if (row.poleId === "P1") row.poleId = "P1-RENAMED";
}));
const renamedUpdate = merge(renamedPrevious, JSON.parse(JSON.stringify(imported)));
assert.ok(renamedUpdate.poles["P1-RENAMED"], "Update Data must retain the operator's renamed pole ID");
assert.equal(renamedUpdate.poles.P1, undefined, "Update Data must not recreate the old pole ID");
assert.equal(renamedUpdate.spans.S1.fromPole, "P1-RENAMED", "Update Data must remap imported span endpoints to the renamed ID");
assert.equal(renamedUpdate.spanSides["S1__P1-RENAMED"].proposedHOA, "22'", "saved span-side work must follow the renamed ID through Update Data");
assert.equal(Object.values(renamedUpdate.spanComms).find(row => row.wireId === "W1").poleId, "P1-RENAMED", "imported communication rows must use the renamed ID");
assert.equal(renamedUpdate.poleIdAliases.P1, "P1-RENAMED", "the original-to-current ID mapping must survive updates");
assert.equal(renamedUpdate.excelReviewSource.collection.rows[0].Id, "P1", "the raw workbook audit snapshot must remain unchanged");

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

const deletedPrevious = {
  poles: { P2: { poleId: "P2" } },
  spans: {}, spanSides: {}, spanComms: {}, spanPower: {},
  makeReadyReferences: [], poleClassChecks: [], settings: {}, ui: {},
  deletedPoleIds: ["P1"]
};
const deletedImported = {
  poles: { "P1 STEEL": { poleId: "P1 STEEL" }, P2: { poleId: "P2" } },
  spans: { RESTORED: { spanId: "RESTORED", fromPole: "P1 STEEL", toPole: "P2" } },
  spanSides: { RESTORED_SIDE: { spanId: "RESTORED", poleId: "P1 STEEL", proposedHOA: "22'" } },
  spanComms: { RESTORED_COMM: { spanId: "RESTORED", poleId: "P2", owner: "CATV", wireId: "W1", midspan: "18'" } },
  spanPower: { RESTORED_POWER: { spanId: "RESTORED", poleId: "P1 STEEL", owner: "UTILITY", midspan: "25'" } },
  makeReadyReferences: [], poleClassChecks: [], settings: {}, ui: {}, excelReviewSource: {}
};
const deletedResult = merge(deletedPrevious, deletedImported);
assert.equal(deletedResult.poles["P1 STEEL"], undefined, "Update Data must not restore a canonically deleted pole");
assert.equal(deletedResult.spans.RESTORED, undefined, "spans connected to a deleted pole must remain suppressed");
assert.equal(Object.keys(deletedResult.spanComms).length, 0, "dependent comm rows must be suppressed with the deleted span");
assert.equal(deletedResult.updateDiagnostics.deletedPolesSuppressed, 1);

console.log("Update Data non-destructive merge tests passed.");
