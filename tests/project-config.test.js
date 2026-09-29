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
const intec = profiles.applyProfileSettings({}, "INTEC");
const csu = profiles.applyProfileSettings({}, "CSU");
const comed = profiles.applyProfileSettings({}, "COMED");
const oppd = profiles.applyProfileSettings({}, "OLSSON_OPPD");

assert.equal(metronet.projectProfile, "METRONET");
assert.equal(metronet.proposedOwner, "MidAm", "Metronet WI must default to MidAm");
assert.equal(metronet.position, "LOW_COMM", "existing Metronet calculation defaults must remain unchanged");
assert.equal(metronet.calculateBackspanMidspan, true, "MidAm must calculate an imported Back Span midspan");
assert.equal(intec.calculateBackspanMidspan, true, "INTEC must calculate a Back Span only when that row owns a midspan");
assert.equal(intec.midspanPrimaryPowerCommClearance, "33\"", "INTEC must use 33 inches from Primary midspan to comm");
assert.equal(metronet.midspanPrimaryPowerCommClearance, "30\"", "Metronet must retain its generic Primary midspan clearance");
assert.equal(metronet.environmentClearances.RAILROAD, "23'6\"", "MidAm railroad crossing must use 23 feet 6 inches");
assert.equal(metronet.environmentClearances.WATER_WITHOUT_SAILBOATS, "14'", "MidAm non-sailboat water crossing must use 14 feet");
assert.equal(metronet.streetlightBracketCommClearance, "20\"", "MidAm streetlight bracket clearance must be configured");
assert.equal(metronet.streetlightDripLoopCommClearance, "12\"", "MidAm uncovered drip-loop clearance must be configured");
assert.equal(metronet.powerGuyCommClearance, "", "MidAm must not apply the removed power guy-to-comm clearance");
assert.equal(profiles.detectProfile({ owners: ["UTILITY > MidAm"] }), "METRONET", "MidAm utility ownership must select Metronet automatically");
assert.equal(profiles.normalizeProfileId("WI"), "INTEC", "WI is a Metronet field, not a separate project profile");
assert.equal(csu.projectProfile, "METRONET", "CSU must be represented as a MetroNet work instruction");
assert.equal(csu.metronetWI, "CSU", "CSU must be selected through the MetroNet WI control");
assert.equal(csu.proposedOwner, "MNT", "CSU must use Proposed MNT as its proposed owner");
assert.equal(csu.position, "LOW_COMM", "CSU must place the proposed attachment at bottom comm");
assert.equal(csu.showServiceDrop, false, "CSU must not expose service-drop movements");
assert.equal(csu.showResagServiceDrop, false, "CSU must not generate service-drop re-sag MR");
assert.equal(csu.polePowerCommsClearance, "52\"", "CSU must reserve 52 inches below low power");
assert.equal(csu.environmentClearances.RAILROAD, "23'6\"", "CSU railroad clearance must use 23 feet 6 inches");
assert.equal(profiles.detectProfile({ fileName: "EXCEL_COCS174 92026.xlsx", owners: ["UTILITY > CSU"] }), "CSU", "CSU workbook should select the CSU profile automatically");
assert.equal(comed.projectProfile, "METRONET", "COMED must use the MetroNet calculation profile");
assert.equal(comed.metronetWI, "COMED", "COMED must select the COMED work instruction");
assert.equal(comed.proposedOwner, "Power", "COMED must use Power as the proposed owner");
assert.equal(comed.polePowerCommsClearance, "40\"", "COMED must inherit MetroNet clearance defaults");
assert.equal(profiles.detectProfile({ fileName: "Metronet WI COMED.xlsx", owners: ["UTILITY > Power"] }), "COMED", "COMED workbook should select the COMED work instruction automatically");
assert.equal(profiles.detectProfile({ fileName: "job.xlsx", owners: ["UTILITY > Power", "Primary > ACSR 1/0 AWG 6/1 RAVEN > Static"] }), "COMED", "COMED conductor data should select the COMED work instruction automatically");
assert.equal(oppd.projectProfile, "OLSSON_OPPD");
assert.equal(oppd.proposedOwner, "Cox", "Olsson OPPD should start with Cox as the proposed owner");
assert.equal(oppd.primaryPowerCommsClearance, "72\"", "Olsson OPPD should use 72 inches from primary to top comm");
assert.equal(oppd.midspanPrimaryPowerCommClearance, "60\"", "Olsson OPPD should use 60 inches from primary midspan to top comm midspan");
assert.equal(oppd.environmentClearances.HIGHWAY, "18'", "Olsson OPPD highway clearance should use 18 feet");
assert.equal(oppd.showServiceDrop, true, "Olsson OPPD must expose imported Service Drop rows and their relocation MR");
assert.equal(oppd.showResagServiceDrop, false, "Olsson OPPD should not inherit INTEC's optional Re-sag Service Drop action");
assert.equal(profiles.detectProfile({ fileName: "job.xlsx", owners: ["UTILITY > OPPD"] }), "OLSSON_OPPD", "OPPD ownership must select the Olsson OPPD profile automatically");

console.log("Project profile tests passed.");
