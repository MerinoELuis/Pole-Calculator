"use strict";

const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");

const logs = [];
const window = {
  performance: { now: (() => { let value = 0; return () => (value += 5); })() },
  console: { info: (...args) => logs.push(args), debug() {}, table() {} }
};
vm.runInNewContext(
  fs.readFileSync(path.join(__dirname, "..", "js", "core", "performance-debug.js"), "utf8"),
  { window, console: window.console },
  { filename: "performance-debug.js" }
);

assert.equal(window.PoleCalcPerformance.isEnabled(), true);
window.PoleCalcPerformance.measure("automatic", () => 42);
window.PoleCalcPerformance.enable();
window.PoleCalcPerformance.measure("sample", () => "ok", { poles: 3 });
const report = window.PoleCalcPerformance.snapshot();
assert.equal(report.totals.sample.count, 1);
assert.equal(report.totals.sample.totalMs, 5);
assert.equal(report.recent[0].details.poles, 3);
window.PoleCalcPerformance.disable();
assert.equal(window.PoleCalcPerformance.isEnabled(), false);

console.log("performance debug tests passed");
