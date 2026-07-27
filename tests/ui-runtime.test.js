"use strict";

const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");

const calls = [];
const document = {
  readyState: "complete",
  getElementById: () => ({ name: "overview" }),
  body: {},
  documentElement: {}
};
function MutationObserver(callback) {
  this.callback = callback;
  this.observe = () => calls.push("observe");
  this.disconnect = () => calls.push("disconnect");
}
const window = {
  document,
  MutationObserver,
  queueMicrotask: callback => callback(),
  UiDomContract: { apply: () => { calls.push("dom"); return 2; } },
  CommTableUI: { refresh: () => { calls.push("comm"); return 1; } },
  SpanColorUI: { refresh: () => { calls.push("colors"); return 3; } },
  AppStore: {}
};
const sourcePath = path.join(__dirname, "..", "js", "ui", "runtime.js");
vm.runInNewContext(fs.readFileSync(sourcePath, "utf8"), { window }, { filename: sourcePath });

calls.length = 0;
const result = window.PoleCalculatorUI.runRefresh(document);
assert.deepEqual(JSON.parse(JSON.stringify(result)), { cards: 2, tablesChanged: 1, coloredSpans: 3 });
assert.deepEqual(calls, ["dom", "comm", "dom", "colors"], "DOM contract runs before table cleanup and color assignment");

console.log("UI runtime tests passed");
