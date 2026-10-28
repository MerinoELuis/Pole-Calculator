"use strict";

const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");

const indexPath = path.join(__dirname, "..", "index.html");
const html = fs.readFileSync(indexPath, "utf8");

assert.match(
  html,
  /<button id="openCalculatorBtn"(?=[^>]*aria-label="Open or close calculator")[^>]*>\s*<svg\b/,
  "The floating calculator button must display an accessible calculator icon."
);

assert.match(html, /<div class="floating-header">\s*<p class="eyebrow">Calculator<\/p>/, "The floating calculator header must center its Calculator label independently of the close button.");
assert.match(html, /<label class="calc-result-field">Result<output class="calc-result" id="calcResult"/, "The calculator output must have an accessible English Result label.");
assert.doesNotMatch(html, /Feet' Inches/, "The redundant Feet' Inches title must not be shown.");

assert.match(
  html,
  /id="calcExpression"[^>]*placeholder="12'6&quot; \+ 3'8&quot; - 1' \+ 2&quot;"/,
  "The expression input must show a mixed multi-operation placeholder."
);

assert.doesNotMatch(
  html,
  /id="calcExpression"[^>]*placeholder="Example:/,
  "The placeholder must not include the Example prefix."
);

assert.doesNotMatch(
  html,
  /id="calcExpression"[^>]*\svalue=/,
  "The example must be a placeholder, not a real input value."
);

console.log("floating calculator UI tests passed");
