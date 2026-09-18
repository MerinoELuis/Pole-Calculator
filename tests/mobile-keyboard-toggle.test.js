"use strict";

const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");

const root = path.join(__dirname, "..");
const html = fs.readFileSync(path.join(root, "index.html"), "utf8");
const app = fs.readFileSync(path.join(root, "js", "app.js"), "utf8");
const floating = fs.readFileSync(path.join(root, "js", "floating-calculator.js"), "utf8");
const css = fs.readFileSync(path.join(root, "css", "styles.css"), "utf8");

assert.match(html, /id="toggleMobileKeyboardBtn"[^>]*aria-pressed="false"/);
assert.match(app, /let mobileKeyboardEnabled = false/);
assert.match(app, /setMobileKeyboardEnabled\(!mobileKeyboardEnabled\)/);
assert.match(floating, /keyboardControl \? keyboardControl\.isEnabled\(\)/);
assert.match(css, /\.mobile-keyboard-toggle \{ display: none; \}/);
assert.match(css, /\.mobile-keyboard-toggle \{ display: inline-flex; \}/);

console.log("mobile keyboard toggle tests passed");
