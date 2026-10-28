"use strict";

const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");

const rows = [
  { Owner: "Communication > Century Link", Size: "0.5in CATV", Insulator: "Single Bolt", comm: true },
  { Owner: "Communication > Cox Communications", Size: "0.5in CATV", Insulator: "Single Bolt", comm: true },
  { Owner: "Communication > MCI Metro", Size: `2\" Communication Bundle Msgr:0.306\"`, comm: true },
  { Owner: "Communication > Century Link", Size: `4\" Communication Bundle Msgr:0.386\"`, comm: true },
  { Owner: "Communication > Unknown ISP", Size: "0.5in CATV", Construction: "Davit", Insulator: "Unknown", comm: true }
];
const window = {};
const excelImporter = {
  isPowerWire: row => Boolean(row.power),
  isCommunicationWire: row => Boolean(row.comm)
};
const modulePath = path.join(__dirname, "..", "js", "import", "excel-review-wire-checks.js");
vm.runInNewContext(fs.readFileSync(modulePath, "utf8"), { window }, { filename: modulePath });

const checks = window.ExcelReviewWireChecks.create({
  add: (result, check) => result.push(check),
  I: () => excelImporter,
  isIntecProject: () => true,
  normalizedInsulator: value => String(value || "").trim().toLowerCase().replace(/\s+/g, " "),
  normalizedText: value => String(value || "").trim().toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/\s+/g, " "),
  pick: (row, names) => names.map(name => row[name]).find(value => value !== undefined),
  rowsForPole: () => rows,
  text: value => String(value ?? "").trim()
});

const result = [];
checks.addCommunicationSizeChecks(result, "P1");
assert.equal(result.filter(item => item.code === "COMM_SIZE_OWNER_MISMATCH").length, 1,
  "Century Link CATV size should fail while Cox CATV size passes");
assert.equal(result.filter(item => item.code === "INVALID_TELCO_BUNDLE_SIZE").length, 1,
  "a Telco bundle outside the approved bundle/messenger pairs should fail");

const intecResult = [];
checks.addIntecWireChecks(intecResult, "P1");
assert.ok(intecResult.some(item => item.code === "UNKNOWN_COMM_OWNER"));
assert.ok(intecResult.some(item => item.code === "INVALID_COMM_INSULATOR"));
assert.ok(intecResult.some(item => item.code === "DAVIT_NOT_ALLOWED"));

console.log("Excel Review wire checks tests passed");
