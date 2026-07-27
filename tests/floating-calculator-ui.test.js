"use strict";

const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");

const indexPath = path.join(__dirname, "..", "index.html");
const html = fs.readFileSync(indexPath, "utf8");

assert.match(
  html,
  /id="openCalculatorBtn"[^>]*>ft'in(?:&quot;|")<\/button>/,
  "The floating calculator button must display ft'in\"."
);

assert.match(
  html,
  /id="calcExpression"[^>]*placeholder="12'6&quot; \+ 3'8&quot;"/,
  "The expression input must show a feet-and-inches addition placeholder."
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
