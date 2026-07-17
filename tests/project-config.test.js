"use strict";

const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");

const sandbox = { window: {} };
const modulePath = path.join(__dirname, "..", "js", "project-config.js");
vm.runInNewContext(fs.readFileSync(modulePath, "utf8"), sandbox, { filename: modulePath });

const profiles = sandbox.window.ProjectProfiles;
const metronet = profiles.applyProfileSettings({}, "METRONET");

assert.equal(metronet.projectProfile, "METRONET");
assert.equal(metronet.proposedOwner, "MidAm", "Metronet WI must default to MidAm");
assert.equal(metronet.position, "LOW_COMM", "existing Metronet calculation defaults must remain unchanged");
assert.equal(profiles.normalizeProfileId("WI"), "INTEC", "WI is a Metronet field, not a separate project profile");

console.log("Project profile tests passed.");
