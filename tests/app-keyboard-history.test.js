"use strict";

const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");

const source = fs.readFileSync(path.join(__dirname, "..", "js", "app.js"), "utf8");

assert.match(source, /const redoHistory = \[\]/, "app must retain redo snapshots");
assert.match(source, /function redoLastAction\(\)/, "app must provide a redo action");
assert.match(source, /redoHistory\.push\(/, "undo must preserve the current state for redo");
assert.match(source, /redoHistory\.length = 0/, "a new edit must clear stale redo history");
assert.match(source, /event\.key\.toLowerCase\(\) === "y"/, "Ctrl+Y must be bound to redo");
assert.match(source, /event\.shiftKey && event\.key\.toLowerCase\(\) === "z"/, "Ctrl+Shift+Z must be bound to redo");

console.log("application undo/redo keyboard contract tests passed");
