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
assert.equal(metronet.calculateBackspanMidspan, true, "MidAm must calculate an imported Back Span midspan");
assert.equal(metronet.environmentClearances.RAILROAD, "23'6\"", "MidAm railroad crossing must use 23 feet 6 inches");
assert.equal(metronet.environmentClearances.WATER_WITHOUT_SAILBOATS, "14'", "MidAm non-sailboat water crossing must use 14 feet");
assert.equal(metronet.streetlightBracketCommClearance, "20\"", "MidAm streetlight bracket clearance must be configured");
assert.equal(metronet.streetlightDripLoopCommClearance, "12\"", "MidAm uncovered drip-loop clearance must be configured");
assert.equal(metronet.powerGuyCommClearance, "", "MidAm must not apply the removed power guy-to-comm clearance");
assert.equal(profiles.detectProfile({ owners: ["UTILITY > MidAm"] }), "METRONET", "MidAm utility ownership must select Metronet automatically");
assert.equal(profiles.normalizeProfileId("WI"), "INTEC", "WI is a Metronet field, not a separate project profile");

console.log("Project profile tests passed.");
