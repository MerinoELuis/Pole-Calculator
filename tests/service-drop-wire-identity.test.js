"use strict";

const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");

const window = {};
const sandbox = { window, console, Date, Set, Map, Math, JSON, Object, Array, String, Number, Boolean, RegExp };
for (const file of ["domain/height-math.js", "core/project-config.js", "domain/pole-resolution.js", "core/state-compat.js", "core/store-identity.js", "core/store.js"]) {
  const modulePath = path.join(__dirname, "..", "js", file);
  vm.runInNewContext(fs.readFileSync(modulePath, "utf8"), sandbox, { filename: modulePath });
}

const owner = "COMMUNICATION > Cox Communications";
const normalWire = "AzUsBNmNyN";
const dropWire = "Gy07a47z1n";
const state = window.AppStore.setState({
  version: "1.8.177",
  settings: { projectProfile: "INTEC" },
  poles: {
    P7: {
      poleId: "P7",
      comms: [
        { owner, existingHOA: "20'2\"", wireId: normalWire, size: "CATV Bundles > 0.5\" Communication Bundle" },
        { owner, existingHOA: "20'2\"", wireId: dropWire, size: "Communication Drops > 0.25in CATV" }
      ]
    },
    P8: { poleId: "P8", comms: [] }
  },
  spans: {
    S1: { spanId: "S1", fromPole: "P7", toPole: "P8", type: "Fore Span" }
  },
  spanComms: {
    [`S1__P7__${owner}__${normalWire}`]: {
      spanId: "S1",
      poleId: "P7",
      owner,
      wireId: normalWire,
      existingHOA: "20'2\"",
      midspan: "14'4\"",
      serviceDrop: true,
      size: "CATV Bundles > 0.5\" Communication Bundle"
    }
  },
  excelReviewSource: {
    spanWires: {
      rows: [
        {
          Id: "P7",
          "Span Id": "S1",
          "Wire Id": normalWire,
          Owner: owner,
          Size: "CATV Bundles > 0.5\" Communication Bundle",
          "Attachment Height.display": "20'2\"",
          "Mid Span Height.display": "14'4\""
        },
        {
          Id: "P7",
          "Span Id": "S1",
          "Wire Id": dropWire,
          Owner: owner,
          Size: "Communication Drops > 0.25in CATV",
          "Attachment Height.display": "20'2\"",
          "Mid Span Height.display": "13'5\""
        }
      ]
    }
  }
});

const rows = Object.values(state.spanComms).filter(row => row.spanId === "S1" && row.poleId === "P7");
assert.equal(rows.length, 2, "each physical CATV wire must retain its own span row");
assert.equal(rows.find(row => row.wireId === normalWire)?.serviceDrop, false, "normal CATV must not inherit a stale drop flag");
assert.equal(rows.find(row => row.wireId === dropWire)?.serviceDrop, true, "CATV drop must remain a service drop");
assert.equal(rows.find(row => row.wireId === dropWire)?.midspan, "13'5\"", "the drop midspan must remain tied to its wire id");

console.log("Service Drop wire identity tests passed.");
