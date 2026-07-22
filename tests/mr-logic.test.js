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
    A: { spanId: "PROP", poleId: "P1", owner: "CATV", ownerBase: "CATV", existingHOA: "20'10\"", transferToNewPole: true, downGuy: true },
    B: { spanId: "BACK", poleId: "P1", owner: "CATV", ownerBase: "CATV", existingHOA: "21'2\"", transferToNewPole: true }
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
assert.match(p1Text, /Backspan to go UG SE due to proposed pole overloaded on adj pole\./);
assert.match(p1Text, /Pl riser W at HOA 18'\./);
assert.match(p1Text, /Transfer CATV to new pole at HOA 20'10" and 21'2" with DG\./);
assert.equal(p1Text.trim().split("\n").at(-1), "Pl riser W at HOA 18'.", "riser must be final and use the imported IO direction");

state.makeReadyReferences = [];
sandbox.window.MRLogic.generateMRForPole("P1");
let generatedWithoutDirection = state.mr.find(item => item.poleId === "P1").text;
assert.match(generatedWithoutDirection, /Pl riser NW at HOA 18'\./, "a Back Span riser must default opposite the span direction");
assert.equal(generatedWithoutDirection.trim().split("\n").at(-1), "Pl riser NW at HOA 18'.", "the defaulted riser must remain the final instruction");
assert.equal(sandbox.window.MRLogic.getDefaultRiserDirection("Forespan", "SE"), "SE", "Fore Span keeps the span direction");
assert.equal(sandbox.window.MRLogic.getDefaultRiserDirection("Backspan", "SE"), "NW", "Back Span reverses the span direction");
assert.equal(sandbox.window.MRLogic.getDefaultRiserDirection("Otherspan", "SE"), "", "Other Span remains pending for manual review");

state.poles.P1.ugRiserDirection = "E";
sandbox.window.MRLogic.generateMRForPole("P1");
const generatedWithOverride = state.mr.find(item => item.poleId === "P1").text;
assert.match(generatedWithOverride, /Pl riser E at HOA 18'\./, "the editable riser direction must override a missing imported direction");
assert.equal(generatedWithOverride.trim().split("\n").at(-1), "Pl riser E at HOA 18'.", "the edited riser must remain the final instruction");

state.poles.P2.ugMRText = [
  "Unable to attach due to red tag.",
  "Red tag",
  "Inability to place ANC"
].join("\n");
sandbox.window.MRLogic.generateMRForPole("P1");
assert.match(
  state.mr.find(item => item.poleId === "P1").text,
  /Backspan to go UG SE due to red tag on adj pole\./,
  "connected UG instructions must extract only the reason from the editable template"
);

sandbox.window.MRLogic.generateMRForPole("P2");
assert.equal(state.mr.find(item => item.poleId === "P2").text, [
  "Unable to attach due to red tag.",
  "Red tag",
  "Inability to place ANC"
].join("\n"));

console.log("Make Ready logic tests passed.");
