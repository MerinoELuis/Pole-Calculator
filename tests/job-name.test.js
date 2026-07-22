"use strict";

const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");

const sandbox = {
  window: { localStorage: { getItem: () => null, setItem: () => {} } },
  console,
  Date,
  Set,
  Map,
  JSON,
  Object,
  Array,
  String,
  Number,
  Boolean,
  RegExp
};
const modulePath = path.join(__dirname, "..", "js", "state.js");
vm.runInNewContext(fs.readFileSync(modulePath, "utf8"), sandbox, { filename: modulePath });

const store = sandbox.window.AppStore;
assert.equal(
  store.jobNameFromFileName("EXCEL_Wecom-SUPRAZ3.3_2026-07-21.xlsx"),
  "Wecom-SUPRAZ3.3",
  "raw Excel prefix and date must not become part of the editable job name"
);

let state = store.setState({ importedFileName: "EXCEL_Wecom-SUPRAZ3.3_2026-07-21.xlsx" });
assert.equal(state.jobName, "Wecom-SUPRAZ3.3", "older JSON without jobName must migrate from its source filename");

state = store.setState({
  importedFileName: "EXCEL_Wecom-SUPRAZ3.3_2026-07-21.xlsx",
  jobName: "SUPRAZ 3.3 - Revised"
});
assert.equal(state.jobName, "SUPRAZ 3.3 - Revised", "an operator-edited job name must remain authoritative");

console.log("Job name tests passed.");
