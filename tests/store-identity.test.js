"use strict";

const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");

const window = {};
const modulePath = path.join(__dirname, "..", "js", "core", "store-identity.js");
vm.runInNewContext(fs.readFileSync(modulePath, "utf8"), { window }, { filename: modulePath });
const identity = window.AppStoreIdentity;

assert.equal(identity.canonicalPoleIdentity("p-12 steel UG PCO"), "P-12");
assert.equal(identity.canonicalPoleIdentity(" Unknown-P3 "), "UNKNOWN-P3");
assert.equal(identity.isUnknownPoleId("Unknown-P3"), true);
assert.equal(identity.isAddressableUnknownPoleId("P3-Unknown-Span"), true);
assert.equal(identity.keyForSpanSide(" S1 ", " P1 "), "S1__P1");
assert.equal(identity.keyForSpanComm("S1", "P1", "Century Link", "W1"), "S1__P1__Century Link__W1");
assert.ok(Object.isFrozen(identity), "identity primitives should be exposed as an immutable dependency surface");

console.log("Store identity tests passed");
