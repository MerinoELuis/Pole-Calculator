"use strict";

const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");

const state = {
  settings: { projectProfile: "INTEC", proposedOwner: "Wecom", mrCase: "LOWER" },
  poles: {
    P1: { poleId: "P1" },
    P2: { poleId: "P2", ugActive: true, ugReason: "proposed pole overloaded" },
    P3: { poleId: "P3" }
  },
  spans: {
    OTHER: { spanId: "OTHER", fromPole: "P1", toPole: "P2", type: "Other", rawType: "Other", direction: "S" },
    BACK: { spanId: "BACK", fromPole: "P1", toPole: "P2", type: "Back Span", rawType: "Back Span", direction: "SE" },
    PROP: { spanId: "PROP", fromPole: "P1", toPole: "P3", type: "Fore Span", rawType: "Fore Span", direction: "N" }
  },
  spanSides: {
    PROP__P1: { spanId: "PROP", poleId: "P1", proposedHOA: "19'" }
  },
  spanComms: {
    A: { spanId: "PROP", poleId: "P1", owner: "CATV", ownerBase: "CATV", existingHOA: "21'6\"", transferToNewPole: true, downGuy: true },
    B: { spanId: "BACK", poleId: "P1", owner: "CATV", ownerBase: "CATV", existingHOA: "21'10\"", transferToNewPole: true },
    C: { spanId: "OTHER", poleId: "P1", owner: "Century Link Communications", ownerBase: "Century Link Communications", existingHOA: "19'10\"", transferToNewPole: true },
    D: { spanId: "BACK", poleId: "P1", owner: "Century Link Communications", ownerBase: "Century Link Communications", existingHOA: "20'2\"", transferToNewPole: true },
    E: { spanId: "PROP", poleId: "P1", owner: "Century Link Communications", ownerBase: "Century Link Communications", existingHOA: "20'6\"", transferToNewPole: true }
  },
  makeReadyReferences: [{ poleId: "P1", makeReadyNotes: "Pl riser W at HOA 18'.", raw: {} }],
  mr: []
};

const AppStore = {
  getState: () => state,
  getPole: poleId => state.poles[poleId] || null,
  getSpan: spanId => state.spans[spanId] || null,
  getConnectedSpans: poleId => Object.values(state.spans).filter(span => span.fromPole === poleId || span.toPole === poleId),
  getOtherPoleId: (span, poleId) => span.fromPole === poleId ? span.toPole : span.fromPole,
  getSpanSidesForPole: poleId => Object.values(state.spanSides).filter(side => side.poleId === poleId),
  getSpanSidesForSpan: spanId => Object.values(state.spanSides).filter(side => side.spanId === spanId),
  getSpanSide: (spanId, poleId) => Object.values(state.spanSides).find(side => side.spanId === spanId && side.poleId === poleId) || null,
  getSpanCommsForPole: poleId => Object.values(state.spanComms).filter(row => row.poleId === poleId),
  getSpanCommsForSpan: spanId => Object.values(state.spanComms).filter(row => row.spanId === spanId)
};

const sandbox = { window: { AppStore }, console, Set, Map };
const heightPath = path.join(__dirname, "..", "js", "height-utils.js");
const mrPath = path.join(__dirname, "..", "js", "mr-logic.js");
vm.runInNewContext(fs.readFileSync(heightPath, "utf8"), sandbox, { filename: heightPath });
vm.runInNewContext(fs.readFileSync(mrPath, "utf8"), sandbox, { filename: mrPath });

const defaultUGTemplate = sandbox.window.MRLogic.getEditableUGTemplate(state.poles.P2);
assert.equal(defaultUGTemplate.split("\n").length, 6, "the UG editor must start with all six template lines");
assert.match(defaultUGTemplate, /^Unable to attach due to proposed pole overloaded\./, "legacy UG Reason must populate the first template line");

sandbox.window.MRLogic.generateMRForPole("P1");
const p1Text = state.mr.find(item => item.poleId === "P1").text;
assert.match(p1Text, /Backspan to go UG SE due to on adj pole proposed pole overloaded\./);
assert.match(p1Text, /Pl riser W at HOA 18'\./);
assert.match(p1Text, /Transfer CATV to new pole at HOA 21'6" and 21'10" with DG\./);
assert.match(p1Text, /Transfer CTL to new pole at HOA 19'10", 20'2" and 20'6"\./);
assert.ok(p1Text.indexOf("Transfer CATV") < p1Text.indexOf("Transfer CTL"), "transfer groups must follow the pole from highest to lowest comm");
assert.equal(p1Text.trim().split("\n").at(-1), "Pl riser W at HOA 18'.", "riser must be final and use the imported IO direction");

state.spanComms.MOVE = {
  spanId: "PROP",
  poleId: "P1",
  owner: "Cable One, Prescott",
  ownerBase: "Cable One, Prescott",
  existingHOA: "22'6\"",
  existingHOAChange: "21'",
  transferToNewPole: false,
  serviceDrop: false,
  downGuy: false
};
state.spanComms.MOVE_DUPLICATE = {
  spanId: "OTHER",
  poleId: "P1",
  owner: "Cable One, Prescott",
  ownerBase: "Cable One, Prescott",
  existingHOA: "22’6”",
  existingHOAChange: "21’",
  transferToNewPole: false,
  serviceDrop: false,
  downGuy: false
};
sandbox.window.MRLogic.generateMRForPole("P1");
const movementText = state.mr.find(item => item.poleId === "P1").text;
assert.match(
  movementText,
  /At HOA 22'6\" lower Cable One to HOA 21'\./,
  "an HOA Change must produce a regular comm movement in MR"
);
assert.equal(
  movementText.split("\n").filter(line => /lower Cable One to HOA 21'\./i.test(line)).length,
  1,
  "the same movement must be emitted only once when source rows use different quote characters"
);
state.spanComms.MOVE.mr = "Review attachment and maintain existing route.";
sandbox.window.MRLogic.generateMRForPole("P1");
const customMovementText = state.mr.find(item => item.poleId === "P1").text;
assert.match(customMovementText, /Review attachment and maintain existing route\./);
assert.match(
  customMovementText,
  /At HOA 22'6\" lower Cable One to HOA 21'\./,
  "an imported custom MR note must not hide its HOA movement"
);
state.spanComms.WECOM_MOVE = {
  spanId: "PROP",
  poleId: "P1",
  owner: "COMMUNICATION > Wecom Inc",
  ownerBase: "COMMUNICATION > Wecom Inc",
  existingHOA: "23'6\"",
  existingHOAChange: "23'",
  transferToNewPole: false,
  serviceDrop: false,
  downGuy: false
};
sandbox.window.MRLogic.generateMRForPole("P1");
const normalizedOwnerMR = state.mr.find(item => item.poleId === "P1").text;
assert.match(normalizedOwnerMR, /lower Wecom to HOA 23'\./, "MR movement owners must omit the Inc suffix");
assert.doesNotMatch(normalizedOwnerMR, /Wecom Inc/, "MR text must not include the Wecom Inc suffix");
delete state.spanComms.WECOM_MOVE;
delete state.spanComms.MOVE;
delete state.spanComms.MOVE_DUPLICATE;

state.settings.showServiceDrop = false;
state.spanComms.DROP = {
  spanId: "PROP",
  poleId: "P1",
  owner: "DropCo",
  ownerBase: "DropCo",
  existingHOA: "18'",
  existingHOAChange: "17'",
  serviceDrop: true,
  transferToNewPole: true,
  resagServiceDrop: true
};
sandbox.window.MRLogic.generateMRForPole("P1");
const noDropMR = state.mr.find(item => item.poleId === "P1")?.text || "";
assert.doesNotMatch(noDropMR, /DropCo|Re-sag/i, "disabled service drops must not generate movement or re-sag MR");
delete state.spanComms.DROP;
state.settings.showServiceDrop = true;

state.makeReadyReferences = [];
sandbox.window.MRLogic.generateMRForPole("P1");
let generatedWithoutDirection = state.mr.find(item => item.poleId === "P1").text;
assert.match(generatedWithoutDirection, /Pl riser SE at HOA 18'\./, "a Back Span riser must use the UG span direction");
assert.equal(generatedWithoutDirection.trim().split("\n").at(-1), "Pl riser SE at HOA 18'.", "the defaulted riser must remain the final instruction");
assert.equal(sandbox.window.MRLogic.getDefaultRiserDirection("Forespan", "SE"), "SE", "Fore Span keeps the span direction");
assert.equal(sandbox.window.MRLogic.getDefaultRiserDirection("Backspan", "SE"), "SE", "Back Span also keeps the UG span direction");
assert.equal(sandbox.window.MRLogic.getDefaultRiserDirection("Otherspan", "SE"), "", "Other Span remains pending for manual review");

state.poles.P1.riserActive = false;
sandbox.window.MRLogic.generateMRForPole("P1");
assert.doesNotMatch(state.mr.find(item => item.poleId === "P1").text, /Pl riser/i, "the Riser action can suppress the automatic Back Span riser");

state.poles.P1.riserActive = true;
state.poles.P1.ugRiserDirection = "E";
sandbox.window.MRLogic.generateMRForPole("P1");
const generatedWithOverride = state.mr.find(item => item.poleId === "P1").text;
assert.match(generatedWithOverride, /Pl riser E at HOA 18'\./, "the editable riser direction must override a missing imported direction");
assert.equal(generatedWithOverride.trim().split("\n").at(-1), "Pl riser E at HOA 18'.", "the edited riser must remain the final instruction");

state.poles.P1.ugRiserDirection = "";
state.poles.P2.ugActive = false;
state.poles.P3.ugActive = false;
sandbox.window.MRLogic.generateMRForPole("P1");
const foreSpanRiser = state.mr.find(item => item.poleId === "P1").text;
assert.doesNotMatch(foreSpanRiser, /to go UG/i, "manual Riser must not require or invent an adjacent UG instruction");
assert.equal(foreSpanRiser.trim().split("\n").at(-1), "Pl riser N at HOA 18'.", "a manually enabled Riser must use the primary Proposed span direction");
state.poles.P1.riserActive = null;
state.poles.P3.ugActive = true;
sandbox.window.MRLogic.generateMRForPole("P1");
const automaticForeSpanRiser = state.mr.find(item => item.poleId === "P1").text;
assert.doesNotMatch(automaticForeSpanRiser, /Forespan to go UG/i, "an adjacent UG without a specific reason must not export a placeholder relation instruction");
assert.equal(automaticForeSpanRiser.trim().split("\n").at(-1), "Pl riser N at HOA 18'.", "Fore Span toward an adjacent UG pole must receive an automatic riser");
state.poles.P3.ugActive = false;
state.poles.P2.ugActive = true;
state.poles.P1.riserActive = true;
state.poles.P1.ugRiserDirection = "E";

state.poles.P2.ugMRText = [
  "Unable to attach due to red tag.",
  "Red tag",
  "Inability to place ANC"
].join("\n");
const ugChangePoleIds = sandbox.window.MRLogic.generateMRForUGChange("P2");
assert.ok(ugChangePoleIds.includes("P1"), "editing UG MR must include the adjacent previous pole");
assert.match(
  state.mr.find(item => item.poleId === "P1").text,
  /Backspan to go UG SE due to on adj pole red tag\./,
  "connected UG instructions must extract only the reason from the editable template"
);

state.poles.P1.ugActive = true;
sandbox.window.MRLogic.generateMRForPole("P1");
assert.doesNotMatch(state.mr.find(item => item.poleId === "P1").text, /Pl riser/i, "Riser must be disabled on the pole's own UG replacement MR");
assert.equal(sandbox.window.MRLogic.isRiserAvailable("P1"), false, "Riser control must be unavailable while the pole is UG");
state.poles.P1.ugActive = false;
state.poles.P1.pcoActive = true;
const defaultPCOTemplate = sandbox.window.MRLogic.getEditablePCOTemplate(state.poles.P1);
assert.equal(defaultPCOTemplate.split("\n").length, 2, "the INTEC PCO editor must start with both replacement options");
state.poles.P1.pcoMRText = "Replace pole due to CATV overload.\nTransfer CATV to new pole.";
sandbox.window.MRLogic.generateMRForPole("P1");
assert.equal(
  state.mr.find(item => item.poleId === "P1").text,
  "Replace pole due to CATV overload.\nTransfer CATV to new pole.",
  "PCO Make Ready must use the operator-edited multiline text"
);
assert.doesNotMatch(state.mr.find(item => item.poleId === "P1").text, /Pl riser/i, "Riser must be disabled on the pole's own PCO replacement MR");
assert.equal(sandbox.window.MRLogic.isRiserAvailable("P1"), false, "Riser control must be unavailable while the pole is PCO");
state.poles.P1.pcoActive = false;

state.poles.P1.poleInsetActive = true;
state.poles.P1.poleInsetReason = "OVERLOADED";
sandbox.window.MRLogic.generateMRForPole("P1");
assert.equal(
  state.mr.find(item => item.poleId === "P1").text.split("\n")[0],
  "Pole overloaded unless comm inset pole placed at midspan to reduce span length.",
  "Pole Inset must generate the overloaded Make Ready wording"
);
state.poles.P1.poleInsetReason = "FAILING_CLEARANCES";
sandbox.window.MRLogic.generateMRForPole("P1");
assert.equal(
  state.mr.find(item => item.poleId === "P1").text.split("\n")[0],
  "Pole failing clearances unless comm inset pole placed at midspan to reduce span length.",
  "Pole Inset must generate the failing-clearances Make Ready wording"
);
state.poles.P1.poleInsetActive = false;

state.spanSides.OTHER__P1 = { spanId: "OTHER", poleId: "P1", proposedHOA: "22'6\"" };
state.spanComms.OVERLASH_MESSENGER = {
  spanId: "OTHER",
  poleId: "P1",
  owner: "COMMUNICATION > Wecom Inc",
  rawOwner: "COMMUNICATION > Wecom Inc",
  existingHOA: "22'6\"",
  size: "Telco Bundles > 0.5\" Communication Bundle Msgr:0.242\""
};
state.makeReadyReferences = [{
  poleId: "P2",
  attachmentType: "Overlash",
  attachmentSizeRaw: "72CT Fiber (N)",
  attachmentDirectionTokens: ["N"]
}];
state.spanSides.BACK__P1 = { spanId: "BACK", poleId: "P1", proposedHOA: "24'8\"" };
state.spanComms.OVERLASH_BACK_MESSENGER = {
  spanId: "BACK",
  poleId: "P1",
  owner: "COMMUNICATION > Wecom Inc",
  rawOwner: "COMMUNICATION > Wecom Inc",
  existingHOA: "24'8\"",
  size: "Telco Bundles > 0.5\" Communication Bundle Msgr:0.242\""
};
state.makeReadyReferences.push({
  poleId: "P2",
  attachmentType: "Overlash",
  attachmentSizeRaw: "72CT Fiber (N)",
  attachmentDirectionTokens: []
});
sandbox.window.MRLogic.generateMRForPole("P1");
const mixedNewOverlashMR = state.mr.find(item => item.poleId === "P1").text;
assert.match(mixedNewOverlashMR, /Overlash Wecom at HOA 22'6"\./, "Overlash must generate its own MR instruction");
assert.doesNotMatch(mixedNewOverlashMR, /Overlash Wecom .*24'8"/, "a reciprocal Back Span must not add an Overlash height to this pole's MR");
assert.match(mixedNewOverlashMR, /Attach Wecom at HOA 19'\./, "a New proposal must still generate its Attach MR on the same pole");
assert.doesNotMatch(mixedNewOverlashMR, /^Attach Wecom at HOA .*22'6"/m, "Overlash must not be merged into the New Attach instruction");
delete state.spanComms.OVERLASH_MESSENGER;
delete state.spanComms.OVERLASH_BACK_MESSENGER;
delete state.spanSides.OTHER__P1;
delete state.spanSides.BACK__P1;

sandbox.window.MRLogic.generateMRForPole("P2");
assert.equal(state.mr.find(item => item.poleId === "P2").text, [
  "Unable to attach due to red tag.",
  "Red tag",
  "Inability to place ANC"
].join("\n"));

state.settings = { projectProfile: "METRONET", metronetWI: "CSU", proposedOwner: "MNT", mrCase: "UPPER" };
state.poles.P2.ugMRText = "";
sandbox.window.MRLogic.generateMRForPole("P2");
assert.equal(
  state.mr.find(item => item.poleId === "P2").text,
  "SUGGEST GOING UG DUE TO [CLEARANCE VIOLATION / 500FT/1500FT AERIAL REQUIREMENT].",
  "Metronet UG must use the editable CSU/MetroNet template"
);
state.poles.P2.ugMRText = "CUSTOM CSU UG INSTRUCTION.";
sandbox.window.MRLogic.generateMRForPole("P2");
assert.equal(state.mr.find(item => item.poleId === "P2").text, "CUSTOM CSU UG INSTRUCTION.", "Metronet UG text must be editable");
state.poles.P2.ugMRText = "";
state.poles.P2.ugActive = true;
state.poles.P1.riserActive = true;
state.poles.P1.ugActive = false;
sandbox.window.MRLogic.generateMRForPole("P1");
assert.match(state.mr.find(item => item.poleId === "P1").text, /PL RISER .* AT HOA 18'/, "Metronet must generate the enabled Riser MR");

console.log("Make Ready logic tests passed.");
