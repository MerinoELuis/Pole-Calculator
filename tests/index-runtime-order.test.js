"use strict";

const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");

const html = fs.readFileSync(path.join(__dirname, "..", "index.html"), "utf8");
const packageJson = JSON.parse(fs.readFileSync(path.join(__dirname, "..", "package.json"), "utf8"));
const deploymentVersion = fs.readFileSync(path.join(__dirname, "..", "js", "core", "deployment-version.js"), "utf8");
const poleRenderer = fs.readFileSync(path.join(__dirname, "..", "js", "ui", "pole-renderer.js"), "utf8");
const scripts = Array.from(html.matchAll(/<script src="([^"]+)"/g), match => match[1]);

function position(filename) {
  const index = scripts.findIndex(src => src.split("?", 1)[0] === filename);
  assert.notEqual(index, -1, `${filename} must be loaded`);
  return index;
}

assert.ok(position("js/core/deployment-version.js") < position("js/ui/app.js"));
assert.ok(position("js/export/compact-autoproposed.js") < position("js/rules/auto-calculate-solver.js"));
assert.ok(position("js/rules/auto-calculate-solver.js") < position("js/ui/app.js"));
assert.ok(position("js/import/update-data-reconciler.js") < position("js/ui/app.js"));
assert.ok(position("js/import/import-normalizers.js") < position("js/import/excel-import.js"));
assert.ok(position("js/import/excel-review-wire-checks.js") < position("js/import/excel-review.js"));
assert.ok(position("js/domain/comm-movements.js") < position("js/domain/clearance-rules.js"));
assert.ok(position("js/domain/pole-inset-calculation.js") < position("js/domain/clearance-rules.js"));
assert.ok(position("js/domain/midspan-engine.js") < position("js/domain/clearance-rules.js"));
assert.ok(position("js/domain/pole-resolution.js") < position("js/core/store.js"));
assert.ok(position("js/core/state-compat.js") < position("js/core/store.js"));
assert.ok(position("js/core/store-identity.js") < position("js/core/store.js"));
assert.ok(position("js/ui/pole-renderer.js") < position("js/ui/app.js"));
assert.ok(position("js/ui/controls.js") < position("js/ui/app.js"));
assert.ok(position("js/core/persistence.js") < position("js/ui/app.js"));
assert.ok(position("js/core/history.js") < position("js/ui/app.js"));
assert.ok(position("js/ui/dialogs.js") < position("js/ui/app.js"));
assert.ok(position("js/ui/theme.js") < position("js/ui/app.js"));
assert.ok(position("js/ui/mobile-keyboard.js") < position("js/ui/app.js"));
assert.ok(position("js/ui/excel-review-renderer.js") < position("js/ui/pole-renderer.js"));
assert.ok(position("js/core/history.js") < position("js/ui/pole-navigation.js"));
assert.ok(position("js/ui/pole-navigation.js") < position("js/ui/app.js"));
assert.ok(position("js/ui/app.js") < position("js/ui/dom-contract.js"));
assert.ok(position("js/ui/dom-contract.js") < position("js/ui/comm-table.js"));
assert.ok(position("js/ui/comm-table.js") < position("js/ui/span-colors.js"));
assert.ok(position("js/ui/span-colors.js") < position("js/ui/runtime.js"));
assert.doesNotMatch(html, /auto-calculate-status/);
assert.doesNotMatch(html, /compact-autoproposed-ug-guard\.js/);
assert.doesNotMatch(html, /comm-table-columns\.js/);
assert.doesNotMatch(html, /span-color-reset\.js/);
assert.match(html, /id="deploymentVersion"/);
assert.match(packageJson.version, /^\d+\.\d+\.\d+$/, "package version must use numeric semantic versioning");
assert.match(deploymentVersion, new RegExp(`version:\\s*"${packageJson.version.replace(/\./g, "\\.")}"`));
assert.match(poleRenderer, /const compassDirections = \["", "N", "NE", "E", "SE", "S", "SW", "W", "NW"\]/);
assert.match(poleRenderer, /const anchorDirectionOptions = compassDirections/);

console.log("runtime script order tests passed");
