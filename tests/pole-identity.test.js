"use strict";

const assert = require("node:assert/strict");

global.window = global;
global.localStorage = { getItem: () => null, setItem: () => {} };
global.Calculations = { recalculateAll: () => {} };

require("../js/domain/height-math.js");
require("../js/core/project-config.js");
require("../js/domain/pole-resolution.js");
require("../js/core/state-compat.js");
require("../js/core/store-identity.js");
require("../js/core/store.js");
require("../js/import/import-normalizers.js");
require("../js/import/excel-import.js");

assert.equal(AppStore.canonicalPoleIdentity("P01-LX339927 STEEL UG"), "P01-LX339927");
assert.equal(AppStore.canonicalPoleIdentity("P01-LX339927 PCO STEEL"), "P01-LX339927");
assert.equal(AppStore.canonicalPoleIdentity("P4-313603 RECOLLECT"), "P4-313603");
assert.equal(AppStore.canonicalPoleIdentity("P4-313603 T STEEL UG PCO"), "P4-313603");
assert.equal(AppStore.canonicalPoleIdentity("P4-T313603"), "P4-T313603", "annotation matching must not strip descriptor-like text inside the ID");

const workbook = {
  SheetNames: ["Collection", "Span", "Span.Wire", "Equipment", "Anchor", "Anchor.Guys", "Make Ready"],
  Sheets: {
    Collection: [
      ["collectionId", "Id", "Sequence", "Lowest Power.display", "Low Power Attachment.display"],
      ["C1", "P01-LX339927 STEEL", "P01", "30'", ""],
      ["C2", "P02-X339926 STEEL", "P02", "", "29'"]
    ],
    Span: [
      ["Id", "Span Id", "Type", "Linked Collection.Title", "Environment"],
      ["P01-LX339927 STEEL UG", "S1", "Fore Span", "P02-X339926", "STREET"],
      ["P02-X339926 STEEL PCO", "S2", "Back Span", "P01-LX339927", "STREET"]
    ],
    "Span.Wire": [
      ["Id", "Span Id", "Owner", "Size", "Attachment Height.display", "Mid Span Height.display", "Wire Id"],
      ["P01-LX339927 UG", "S1", "COMMUNICATION > CATV", "CATV Bundle", "20'", "16'", "W1"],
      ["P02-X339926", "S2", "COMMUNICATION > Telco", "Telco Bundle", "19'", "16'", "W2"]
    ],
    Equipment: [
      ["Id", "Equipment Id", "Type", "Owner", "Orientation", "Attachment Height.display", "Bottom Height.display", "Drip Loop Height.display"],
      ["P01-LX339927 STEEL", "E1", "Transformer_Single Phase", "UTILITY > APS", "E", "27'", "", ""],
      ["P01-LX339927 STEEL", "E2", "Streetlight > SMALL", "UTILITY > APS", "S", "25'", "24'6\"", "24'4\""],
      ["P01-LX339927 STEEL", "E3", "Generic Equipment > Riser 90.0°", "UTILITY > APS", "N", "19'", "", ""],
      ["P01-LX339927 STEEL", "E4", "Generic Equipment > Riser 90.0°", "COMMUNICATION > CATV", "W", "18'", "", ""]
    ],
    Anchor: [
      ["collectionId", "Id", "Anchor Index", "Anchor Id", "Type", "Lead Length.display", "Lead Length.provider", "Lead Length.bearing.display", "Lead Length.pitch.display", "Owner", "Guys", 'Anchor Violation: <6" or >18" from GL'],
      ["C1", "P01-LX339927 STEEL", 1, "A1", "Single - 14\" - Soil Class 4", "20'", "manual", "90°", "-10°", "UTILITY > APS", 1, "Yes"]
    ],
    "Anchor.Guys": [
      ["collectionId", "Id", "Anchor Index", "Anchor Id", "Guys Index", "Guys Id", "Size", "Owner", "Attachment Height.display"],
      ["C2", "P02-X339926 STEEL", 1, "A2", 1, "G2", "Down > EHS 1/4 DG", "UTILITY > OPPD", "19'"],
      ["C1", "P01-LX339927 STEEL", 1, "A1", 1, "G1", "Down > EHS 5/16 DG", "UTILITY > APS", "23'8\""]
    ],
    "Make Ready": [
      ["Id", "Attachment Size", "Attachment Height.display"],
      ["P01-LX339927 PCO", "6.6M 24CT Fiber (S) + 72CT Fiber (E/W)", "22'"]
    ]
  }
};

const state = ExcelImport.importOriginalWorkbook(workbook, "identity-test.xlsx");
assert.deepEqual(Object.keys(state.poles).sort(), ["P01-LX339927 STEEL", "P02-X339926 STEEL"]);
assert.equal(state.poles["P01-LX339927 STEEL"].lowPower, "30'", "new MidAm Lowest Power.display must import");
assert.equal(state.poles["P02-X339926 STEEL"].lowPower, "29'", "legacy Low Power Attachment.display must remain supported");
assert.deepEqual(
  state.poles["P01-LX339927 STEEL"].metadata.powerEquipment.map(row => row.category),
  ["TRANSFORMER", "STREETLIGHT", "RISER"],
  "power Equipment must include transformers, streetlights and Utility risers but exclude communication risers"
);
assert.equal(state.excelReviewSource.anchors.rows.length, 1, "Anchor and Anchor.Guys must be preserved as separate review sheets");
assert.equal(state.poles["P01-LX339927 STEEL"].metadata.anchors[0].leadLength, "20'", "anchor lead length should be normalized for the ANC editor");
assert.equal(state.poles["P01-LX339927 STEEL"].metadata.anchors[0].bearing, "90°", "anchor bearing should be editable in the ANC panel");
assert.equal(state.poles["P01-LX339927 STEEL"].metadata.anchors[0].groundlineViolation, true, "groundline violation should import as an editable yes/no value");
assert.equal(state.poles["P01-LX339927 STEEL"].metadata.anchors[0].guys[0].attachmentHeight, "23'8\"", "guys should be associated with their Anchor Id");
assert.equal(state.spans.S1.fromPole, "P01-LX339927 STEEL");
assert.equal(state.spans.S1.toPole, "P02-X339926 STEEL");
assert.equal(state.spans.S2.fromPole, "P02-X339926 STEEL");
assert.equal(state.spans.S2.toPole, "P01-LX339927 STEEL");
assert.ok(AppStore.getSpanCommsForPole("P01-LX339927 STEEL").some(row => row.spanId === "S1" && row.existingHOA === "20'"));
assert.ok(AppStore.getSpanCommsForPole("P02-X339926 STEEL").some(row => /Telco/i.test(row.owner) && row.downGuy), "a utility-owned DG at the only non-service comm HOA must be retained");
assert.equal(state.makeReadyReferences[0].poleId, "P01-LX339927 STEEL");
assert.deepEqual(
  state.makeReadyReferences[0].attachmentDirectionTokens,
  ["S", "E", "W"],
  "Make Ready attachment sizes must retain every directional component"
);

const ambiguousCollectionWorkbook = {
  SheetNames: ["Collection", "Span"],
  Sheets: {
    Collection: [
      ["Collection ID", "Id"],
      ["C902", "P01"],
      ["C902", "P02"]
    ],
    Span: [
      ["Id", "Span Id", "Type"],
      ["P01", "S1", "Fore Span"]
    ]
  }
};
const ambiguousCollectionState = ExcelImport.importOriginalWorkbook(ambiguousCollectionWorkbook, "ambiguous-collection.xlsx");
assert.deepEqual(Object.keys(ambiguousCollectionState.poles).filter(id => /^P0[12]$/.test(id)).sort(), ["P01", "P02"], "a reused collection ID must not merge different base pole IDs");
assert.equal(
  ambiguousCollectionState.spans.S1.toPole,
  "Unknown-S1",
  "a repeated Collection ID without a linked pole must not resolve every span to the last collection row"
);
assert.ok(ambiguousCollectionState.poles["Unknown-S1"].isGenerated, "unresolved linked poles must remain generated Unknown poles");

const revisedPoleWorkbook = {
  SheetNames: ["Collection", "Span", "Span.Wire"],
  Sheets: {
    Collection: [
      ["collectionId", "Id", "Sequence", "updatedAt", "Tip.display", "Lowest Power.display", "Type", "Status"],
      ["C-P3", "P3-313602", "P3", "2026-11-15T10:00:00Z", "33'", "25'", "Douglas Fir > 3 > 40", "300 - HOA"],
      ["C-P4", "P4-313603", "P4", "2026-11-15T10:00:00Z", "33'", "26'4\"", "Southern Pine > 1 > 40", "300 - HOA"],
      ["C-P4", "P4-313603 RECOLLECT", "P4", "2026-11-16T10:00:00Z", "33'", "26'4\"", "Southern Pine > 1 > 40", "500 - Final QC"]
    ],
    Span: [
      ["Id", "Span Id", "Type", "Linked Collection.Title", "Environment"],
      ["P3-313602", "S-P34", "Fore Span", "P4-313603 RECOLLECT", "STREET"]
    ],
    "Span.Wire": [
      ["Id", "Span Id", "Owner", "Size", "Attachment Height.display", "Mid Span Height.display", "Wire Id"],
      ["P3-313602", "S-P34", "COMMUNICATION > Eagle West LLC", "Telco Bundle", "21'", "21'9\"", "WIRE-P34"],
      ["P4-313603", "S-P34", "COMMUNICATION > Eagle West LLC", "Telco Bundle", "22'8\"", "", "WIRE-P34"]
    ]
  }
};
const revisedPoleState = ExcelImport.importOriginalWorkbook(revisedPoleWorkbook, "revised-pole.xlsx");
assert.deepEqual(Object.keys(revisedPoleState.poles).sort(), ["P3-313602", "P4-313603 RECOLLECT"], "Collection revisions that only add/remove a status suffix must import as one pole");
assert.equal(revisedPoleState.spans["S-P34"].toPole, "P4-313603 RECOLLECT", "span links to an older suffix variant must resolve to the latest Collection pole ID");
assert.equal(revisedPoleState.poleClassChecks.filter(row => AppStore.canonicalPoleIdentity(row.poleId) === "P4-313603").length, 1, "pole type checks must not duplicate Collection revisions of one pole");
assert.equal(AppStore.getSpanCommsForPole("P4-313603 RECOLLECT").find(row => row.wireId === "WIRE-P34")?.existingHOA, "22'8\"", "Span.Wire rows using an older suffix variant must attach to the one normalized pole");

(async () => {
  const restored = await ExcelImport.importJsonFile({
    name: "self-linked-debug.json",
    text: async () => JSON.stringify({
      state: {
        poles: { P13: { poleId: "P13" } },
        spans: {
          S1: { spanId: "S1", fromPole: "P13", toPole: "P14", rawSpanId: "S1" },
          S2: { spanId: "S2", fromPole: "P13", toPole: "P13", rawSpanId: "S2" }
        },
        excelReviewSource: {
          collection: { rows: [{ "Collection ID": "C902", Id: "P13" }, { "Collection ID": "C902", Id: "P14" }] },
          spans: { rows: [{ Id: "P13", "Span Id": "S1", Type: "Fore Span" }] }
        },
        spanSides: {},
        spanComms: {}
      }
    })
  });
  assert.equal(restored.spans.S1.toPole, "Unknown-S1", "debug JSON must repair stale ambiguous links");
  assert.equal(restored.spans.S2.toPole, "Unknown-S2", "debug JSON must repair a stale self-link even without a source row");
  assert.ok(restored.poles["Unknown-S1"].isGenerated, "repaired JSON endpoints must remain generated Unknown poles");
  console.log("Canonical pole identity import tests passed.");
})().catch(error => {
  console.error(error);
  process.exitCode = 1;
});
