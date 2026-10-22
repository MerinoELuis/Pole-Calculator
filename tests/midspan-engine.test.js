"use strict";

const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");

const window = {};
const sandbox = { window, String, Array, Object, Number, Set, RegExp, Math, Boolean };
const modulePath = path.join(__dirname, "..", "js", "domain", "midspan-engine.js");
vm.runInNewContext(fs.readFileSync(modulePath, "utf8"), sandbox, { filename: modulePath });

function parseHeight(value) {
  const match = String(value || "").match(/^(\d+)'(?:\s*(\d+)")?$/);
  return match ? Number(match[1]) * 12 + Number(match[2] || 0) : null;
}
function formatHeight(inches) {
  const feet = Math.floor(inches / 12);
  const remainder = inches % 12;
  return `${feet}'${remainder ? ` ${remainder}"` : ""}`;
}

const remote = { poleId: "P2", existingHOA: "20'", currentHOA: "19'" };
const comm = { spanId: "S1", poleId: "P1", owner: "CATV", wireId: "W1", existingHOA: "20'", currentHOA: "18'" };
const calculated = window.MidspanEngine.calculateCommMidspanDetails(comm, {
  store: { getSpan: () => ({ spanId: "S1" }) },
  heightUtils: { parseHeight },
  format: formatHeight,
  getEffectiveCommHOA: row => row.currentHOA,
  findRemoteComm: () => remote,
  ownerCanonicalOf: () => "CATV",
  findMidspanSourceComm: row => row,
  getImportedMidspanInchesForComm: () => 20 * 12
});

assert.equal(calculated.calculated, "18' 6\"");
assert.equal(calculated.localAdjustment, -12);
assert.equal(calculated.remoteAdjustment, -6);
assert.equal(calculated.remote, remote);
assert.equal(calculated.remoteHOA, "19'");
assert.equal(window.MidspanLogic, window.MidspanEngine);

console.log("midspan engine tests passed");
