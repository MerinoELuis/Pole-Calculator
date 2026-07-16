"use strict";

const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");

const state = {
  settings: { projectProfile: "INTEC", proposedOwner: "Wecom", mrCase: "LOWER" },
  poles: {
    P1: { poleId: "P1" },
    P2: { poleId: "P2", ugActive: true },
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

sandbox.window.MRLogic.generateMRForPole("P1");
const p1Text = state.mr.find(item => item.poleId === "P1").text;
assert.match(p1Text, /Backspan to go UG SE due to existing pole overloaded\./);
assert.match(p1Text, /Pl riser S at HOA 18'\./);
assert.match(p1Text, /Transfer CATV to new pole at HOA 20'10" and 21'2" with DG\./);
assert.equal(p1Text.trim().split("\n").at(-1), "Pl riser S at HOA 18'.", "riser must be the final MR instruction");

sandbox.window.MRLogic.generateMRForPole("P2");
assert.equal(state.mr.find(item => item.poleId === "P2").text, "Unable to attach due to proposed pole overloaded.");

console.log("Make Ready logic tests passed.");
