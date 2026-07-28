"use strict";

const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");

const html = fs.readFileSync(path.join(__dirname, "..", "index.html"), "utf8");
const scripts = Array.from(html.matchAll(/<script src="([^"]+)"/g), match => match[1]);

function position(filename) {
  const index = scripts.indexOf(filename);
  assert.notEqual(index, -1, `${filename} must be loaded`);
  return index;
}

assert.ok(position("js/compact-autoproposed.js") < position("js/auto-calculate-solver.js"));
assert.ok(position("js/auto-calculate-solver.js") < position("js/auto-calculate-source-compat.js"));
assert.ok(position("js/auto-calculate-source-compat.js") < position("js/app.js"));
assert.ok(position("js/app.js") < position("js/ui/dom-contract.js"));
assert.ok(position("js/ui/dom-contract.js") < position("js/ui/comm-table.js"));
assert.ok(position("js/ui/comm-table.js") < position("js/ui/span-colors.js"));
assert.ok(position("js/ui/span-colors.js") < position("js/ui/auto-calculate-status.js"));
assert.ok(position("js/ui/auto-calculate-status.js") < position("js/ui/runtime.js"));
assert.match(html, /css\/auto-calculate-status\.css/);
assert.doesNotMatch(html, /compact-autoproposed-ug-guard\.js/);
assert.doesNotMatch(html, /comm-table-columns\.js/);
assert.doesNotMatch(html, /span-color-reset\.js/);

console.log("runtime script order tests passed");
