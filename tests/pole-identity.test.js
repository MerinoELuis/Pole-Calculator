"use strict";

const assert = require("node:assert/strict");

global.window = global;
global.localStorage = { getItem: () => null, setItem: () => {} };
global.Calculations = { recalculateAll: () => {} };

require("../js/height-utils.js");
require("../js/state.js");
require("../js/excel-import.js");

assert.equal(AppStore.canonicalPoleIdentity("P01-LX339927 STEEL UG"), "P01-LX339927");
assert.equal(AppStore.canonicalPoleIdentity("P01-LX339927 PCO STEEL"), "P01-LX339927");

const workbook = {
  SheetNames: ["Collection", "Span", "Span.Wire", "Make Ready"],
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
      ["P01-LX339927 UG", "S1", "COMMUNICATION > CATV", "CATV Bundle", "20'", "16'", "W1"]
    ],
    "Make Ready": [
      ["Id", "Attachment Size", "Attachment Height.display"],
      ["P01-LX339927 PCO", "6.6M 24CT Fiber (E/W)", "22'"]
    ]
  }
};

const state = ExcelImport.importOriginalWorkbook(workbook, "identity-test.xlsx");
assert.deepEqual(Object.keys(state.poles).sort(), ["P01-LX339927 STEEL", "P02-X339926 STEEL"]);
assert.equal(state.poles["P01-LX339927 STEEL"].lowPower, "30'", "new MidAm Lowest Power.display must import");
assert.equal(state.poles["P02-X339926 STEEL"].lowPower, "29'", "legacy Low Power Attachment.display must remain supported");
assert.equal(state.spans.S1.fromPole, "P01-LX339927 STEEL");
assert.equal(state.spans.S1.toPole, "P02-X339926 STEEL");
assert.equal(state.spans.S2.fromPole, "P02-X339926 STEEL");
assert.equal(state.spans.S2.toPole, "P01-LX339927 STEEL");
assert.ok(AppStore.getSpanCommsForPole("P01-LX339927 STEEL").some(row => row.spanId === "S1" && row.existingHOA === "20'"));
assert.equal(state.makeReadyReferences[0].poleId, "P01-LX339927 STEEL");

console.log("Canonical pole identity import tests passed.");
