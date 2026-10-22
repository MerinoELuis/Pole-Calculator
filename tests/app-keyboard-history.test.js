"use strict";

const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");

const source = fs.readFileSync(path.join(__dirname, "..", "js", "core", "history.js"), "utf8");
const app = fs.readFileSync(path.join(__dirname, "..", "js", "ui", "app.js"), "utf8");
const poleRenderer = fs.readFileSync(path.join(__dirname, "..", "js", "ui", "pole-renderer.js"), "utf8");

assert.match(source, /const redoHistory = \[\]/, "app must retain redo snapshots");
assert.match(source, /function redoLastAction\([^)]*\)/, "history module must provide a redo action");
assert.match(source, /redoHistory\.push\(/, "undo must preserve the current state for redo");
assert.match(source, /redoHistory\.length = 0/, "a new edit must clear stale redo history");
assert.match(app, /event\.key\.toLowerCase\(\) === "y"/, "Ctrl+Y must be bound to redo");
assert.match(app, /event\.shiftKey && event\.key\.toLowerCase\(\) === "z"/, "Ctrl+Shift+Z must be bound to redo");
assert.match(source, /const undoTrace = \[\]/, "app must retain the undo/redo debug trace");
assert.match(source, /global\.AppUndoTrace/, "history module must expose the undo/redo debug trace to Debug export");
assert.match(app, /history\.restoreTransaction\(undoStateBefore\)/, "Auto Calculate cancellation must restore the undo stack");
assert.doesNotMatch(source, /commitActiveEditableInputBeforeUndo/, "Ctrl+Z must not commit the active input first");
assert.match(poleRenderer, /function poleHasProposed\(poleId\)/, "pole badges must detect an applied Proposed value");
assert.match(poleRenderer, /<span class="badge changed">Proposed<\/span>/, "pole header must show Proposed instead of Changed");
assert.match(poleRenderer, /const hasMR = .*S\.getState\(\)\.mr/, "pole header must detect Make Ready for the current pole");
assert.match(poleRenderer, /<span class="badge warning" title="Make Ready exists for this pole">MR<\/span>/, "pole header must show MR when Make Ready exists");
assert.match(poleRenderer, /<span class="mini-dot warning" title="Make Ready exists for this pole">MR<\/span>/, "pole index must show MR when Make Ready exists");
assert.doesNotMatch(poleRenderer, /<span class="badge changed">Changed<\/span>/, "pole header must not show the obsolete Changed badge");

console.log("application undo/redo keyboard contract tests passed");
