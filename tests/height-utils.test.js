"use strict";

const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");

const sandbox = { window: {}, console };
const modulePath = path.join(__dirname, "..", "js", "height-utils.js");
vm.runInNewContext(fs.readFileSync(modulePath, "utf8"), sandbox, { filename: modulePath });

const parseHeight = sandbox.window.HeightUtils.parseHeight;
assert.equal(parseHeight("15'6´´"), 186, "acute-accent double mark must parse as inches");
assert.equal(parseHeight("15'6′′"), 186, "prime double mark must parse as inches");
assert.equal(parseHeight("6´´"), 6, "inch-only acute-accent input must parse");

console.log("Height input tests passed.");
