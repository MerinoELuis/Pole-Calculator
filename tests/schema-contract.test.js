"use strict";

const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");

const root = path.join(__dirname, "..");
const schema = JSON.parse(fs.readFileSync(path.join(root, "schemas", "autoproposed.schema.json"), "utf8"));

assert.equal(schema.$schema, "https://json-schema.org/draft/2020-12/schema");
assert.equal(schema.additionalProperties, false);
assert.deepEqual(schema.required, ["sizes", "owner", "poles"]);
assert.equal(schema.$defs.pole.additionalProperties, false);
assert.equal(schema.$defs.span.additionalProperties, false);
assert.equal(schema.$defs.move.additionalProperties, false);

function exactKeys(value, allowed, context) {
  Object.keys(value || {}).forEach(key => assert.ok(allowed.includes(key), `${context}: unexpected property ${key}`));
}

function validatePayload(payload, filename) {
  exactKeys(payload, ["sizes", "owner", "poles"], filename);
  assert.equal(typeof payload.owner, "string");
  assert.ok(payload.owner.trim());
  assert.ok(Array.isArray(payload.poles));
  assert.equal(typeof payload.sizes, "object");
  exactKeys(payload.sizes, ["messenger", "fiber"], `${filename}.sizes`);
  assert.ok(payload.sizes.messenger === null || (typeof payload.sizes.messenger === "number" && payload.sizes.messenger > 0));
  assert.equal(typeof payload.sizes.fiber, "object");
  Object.entries(payload.sizes.fiber).forEach(([count, size]) => {
    assert.match(count, /^\d+$/);
    assert.ok(size === null || (typeof size === "number" && size > 0));
  });

  payload.poles.forEach((pole, poleIndex) => {
    const context = `${filename}.poles[${poleIndex}]`;
    exactKeys(pole, ["id", "pco", "terminalHoa", "spans", "moves"], context);
    assert.equal(typeof pole.id, "string");
    assert.ok(pole.id.trim());
    if ("terminalHoa" in pole) assert.equal(Number.isInteger(pole.terminalHoa), true);
    if ("pco" in pole) assert.equal(typeof pole.pco, "boolean");
    if (pole.pco === true) {
      assert.equal("moves" in pole, false, `${context}: PCO cannot contain moves`);
      assert.equal("terminalHoa" in pole, false, `${context}: PCO cannot contain terminalHoa`);
    }

    (pole.moves || []).forEach((move, moveIndex) => {
      exactKeys(move, ["owner", "from", "to", "service", "dg"], `${context}.moves[${moveIndex}]`);
      assert.equal(typeof move.owner, "string");
      assert.equal(Number.isInteger(move.from), true);
      assert.equal(Number.isInteger(move.to), true);
    });

    (pole.spans || []).forEach((span, spanIndex) => {
      const spanContext = `${context}.spans[${spanIndex}]`;
      exactKeys(span, ["to", "kind", "bearing", "length", "hoa", "fiber", "endDrop", "nextHoa", "ug"], spanContext);
      assert.equal(typeof span.to, "string");
      assert.ok(["F", "B", "O"].includes(span.kind));
      ["length", "hoa", "endDrop", "nextHoa"].forEach(field => {
        if (field in span) assert.equal(Number.isInteger(span[field]), true, `${spanContext}.${field} must be integer inches`);
      });
      if (span.ug === true) {
        ["hoa", "fiber", "endDrop", "nextHoa"].forEach(field => {
          assert.equal(field in span, false, `${spanContext}: UG cannot contain ${field}`);
        });
      }
    });
  });
}

for (const filename of ["autoproposed-ug.json", "autoproposed-pco.json", "autoproposed-multi-span.json"]) {
  validatePayload(JSON.parse(fs.readFileSync(path.join(__dirname, "fixtures", filename), "utf8")), filename);
}
validatePayload(JSON.parse(fs.readFileSync(path.join(__dirname, "expected", "compact-normal-output.json"), "utf8")), "compact-normal-output.json");

console.log("AutoProposed schema and fixture contract tests passed");
