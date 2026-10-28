"use strict";

const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");

const window = {
  localStorage: { getItem: () => null, setItem: () => {} },
  Validations: { validateAll: () => {}, validatePole: () => {}, validateSpan: () => {} }
};
const sandbox = { window, console, Date, Set, Map, Math, JSON, Object, Array, String, Number, Boolean, RegExp };

for (const file of ["domain/height-math.js", "core/project-config.js", "domain/pole-resolution.js", "core/state-compat.js", "core/store-identity.js", "core/store.js", "rules/mr-logic.js", "domain/comm-movements.js", "domain/midspan-engine.js", "domain/pole-inset-calculation.js", "domain/clearance-rules.js", "ui/excel-review-renderer.js", "ui/pole-renderer.js"]) {
  const modulePath = path.join(__dirname, "..", "js", file);
  vm.runInNewContext(fs.readFileSync(modulePath, "utf8"), sandbox, { filename: modulePath });
}

const S = window.AppStore;
const C = window.Calculations;

S.resetState();
S.applyProjectProfile("INTEC");
S.upsertPole(S.createPole({ poleId: "P1", lowPower: "25'", metadata: { lowPowerBaseline: "25'" } }));
S.upsertPole(S.createPole({ poleId: "P2", lowPower: "25'", metadata: { lowPowerBaseline: "25'" } }));
S.upsertSpan(S.createSpan("S1", "P1", "P2", "E", "", { type: "Fore Span" }));
S.addSpanPower({ spanId: "S1", poleId: "P1", wireId: "POWER-1", label: "Secondary", attachmentHeight: "24'", midspan: "20'" });
S.addSpanPower({ spanId: "S1", poleId: "P2", wireId: "POWER-1", label: "Secondary", attachmentHeight: "24'", midspan: "20'" });
const p1Key = S.getSpanPowerForPole("P1")[0].key;
const p2Key = S.getSpanPowerForPole("P2")[0].key;

S.updateSpanPowerField(p1Key, "attachmentHeightChange", "26'");
S.updateSpanPowerField(p2Key, "attachmentHeightChange", "23'");
C.recalculateSpansForPole("P1");

assert.equal(S.getState().spanPower[p1Key].attachmentHeight, "24'", "HOA Change must not overwrite imported Attachment Height");
assert.equal(S.getState().spanPower[p1Key].midspan, "20'", "movement calculations must preserve the imported Midspan baseline");
assert.equal(S.getState().spanPower[p1Key].calculatedMidspan, "20'6\"", "both endpoint height changes must adjust the same physical wire Midspan by half their total delta");
assert.equal(S.getSpan("S1").midspanLowPower, "20'6\"", "span Low Power Midspan must use the adjusted power-wire Midspan");
assert.equal(S.getSpan("S1").midspanMaxCommHeight, "18'", "Midspan Max Height must recalculate from the adjusted power-wire Midspan");
assert.equal(S.getPole("P2").lowPower, "23'", "lowering a non-primary Power wire must lower effective pole Low Power");
assert.equal(S.getPole("P2").maxCommHeight, "19'8\"", "pole maximum height must follow the moved Low Power wire");
S.updatePoleField("P1", "poleInsetSpanId", "S1");
assert.equal(C.getPoleInsetDetails("P1").maxHeightAtMidspan, "17'2\"", "Pole Inset must also use the adjusted Midspan for its selected span");
S.setState(JSON.parse(JSON.stringify(S.getState())));
assert.equal(S.getSpanPowerForPole("P1")[0].attachmentHeightChange, "26'", "Save/Load state normalization must retain Power HOA Change");
assert.equal(S.getSpanPowerForPole("P1")[0].calculatedMidspan, "20'6\"", "Save/Load state normalization must retain the derived Power Midspan");

S.resetState();
S.applyProjectProfile("INTEC");
S.upsertPole(S.createPole({ poleId: "A", lowPower: "32'", metadata: { lowPowerBaseline: "32'" } }));
S.upsertPole(S.createPole({ poleId: "B", lowPower: "32'", metadata: { lowPowerBaseline: "32'" } }));
S.upsertPole(S.createPole({ poleId: "C", lowPower: "32'", metadata: { lowPowerBaseline: "32'" } }));
S.upsertSpan(S.createSpan("AB", "A", "B", "E", "", { type: "Fore Span" }));
S.upsertSpan(S.createSpan("BC", "B", "C", "E", "", { type: "Fore Span" }));
S.addSpanPower({ spanId: "AB", poleId: "A", wireId: "REUSED-ID", label: "Secondary", attachmentHeight: "24'", midspan: "20'" });
S.addSpanPower({ spanId: "BC", poleId: "B", wireId: "REUSED-ID", label: "Secondary", attachmentHeight: "24'", midspan: "20'" });
const aKey = S.getSpanPowerForPole("A")[0].key;
const bcKey = S.getSpanPowerForPole("B")[0].key;
S.updateSpanPowerField(aKey, "attachmentHeightChange", "26'");
S.updateSpanPowerField(bcKey, "attachmentHeightChange", "30'");
C.recalculateSpansForPole("A");
assert.equal(S.getState().spanPower[aKey].calculatedMidspan, "21'", "a wire ID reused on a neighboring physical span must not contribute remote movement");

S.resetState();
S.applyProjectProfile("INTEC");
S.upsertPole(S.createPole({ poleId: "PRIMARY-A", lowPower: "32'", metadata: { lowPowerBaseline: "32'" } }));
S.upsertPole(S.createPole({ poleId: "PRIMARY-B", lowPower: "32'", metadata: { lowPowerBaseline: "32'" } }));
S.upsertSpan(S.createSpan("PRIMARY-SPAN", "PRIMARY-A", "PRIMARY-B", "E", "", { type: "Fore Span" }));
S.addSpanPower({ spanId: "PRIMARY-SPAN", poleId: "PRIMARY-A", wireId: "PRIMARY-1", label: "Primary", attachmentHeight: "30'", midspan: "30'" });
const primaryKey = S.getSpanPowerForPole("PRIMARY-A")[0].key;
S.updateSpanPowerField(primaryKey, "attachmentHeightChange", "31'");
C.recalculateSpansForPole("PRIMARY-A");
assert.equal(S.getPole("PRIMARY-A").maxCommHeight, "27'5\"", "Primary HOA Change must update the pole's Primary-to-comm maximum");
assert.equal(S.getSpan("PRIMARY-SPAN").midspanMaxCommHeight, "27'9\"", "Primary midspan movement must update Midspan Max Height using the INTEC Primary rule");

S.resetState();
S.applyProjectProfile("INTEC");
S.upsertPole(S.createPole({ poleId: "UI", lowPower: "25'" }));
S.upsertPole(S.createPole({ poleId: "UI-REMOTE", lowPower: "25'" }));
S.upsertSpan(S.createSpan("UI-SPAN", "UI", "UI-REMOTE", "E", "", { type: "Fore Span" }));
S.addSpanPower({ spanId: "UI-SPAN", poleId: "UI", wireId: "UI-POWER", label: "Secondary", attachmentHeight: "24'", attachmentHeightChange: "26'", midspan: "19'10\"", calculatedMidspan: "20'10\"" });
const renderer = window.PoleWorkspaceRenderer.create({
  H: window.HeightUtils,
  S,
  escapeHtml: value => String(value ?? "").replaceAll("&", "&amp;").replaceAll('"', "&quot;").replaceAll("<", "&lt;").replaceAll(">", "&gt;"),
  poleResolutionOf: () => "AERIAL",
  spanRelationOf: span => S.getSpanRelation(span)
});
const powerHtml = renderer.renderPowerTable("UI");
assert.match(powerHtml, /<th>HOA Change<\/th><th>Midspan<\/th>/, "Power table must place HOA Change between Attachment Height and Midspan");
assert.match(powerHtml, /data-field="attachmentHeightChange" value="26'/, "Power HOA Change must be editable per wire");
assert.match(powerHtml, /Original Midspan before movements: 19'10&quot;/, "hovering the power Midspan must expose its original value");
assert.match(powerHtml, /Adjusted: 20'10&quot;/, "the resulting Midspan must remain visible next to its imported baseline");
assert.doesNotMatch(powerHtml, /Max Height at MS/, "Power table must omit the maximum-height column while retaining its internal calculations");

console.log("Power wire movement tests passed.");
