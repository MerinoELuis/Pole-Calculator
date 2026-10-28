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
for (const poleId of ["T", "T-E", "T-N"]) {
  S.upsertPole(S.createPole({ poleId, lowPower: "32'", metadata: { lowPowerBaseline: "32'" } }));
}
S.upsertSpan(S.createSpan("T-E-SPAN", "T", "T-E", "E", "", { type: "Fore Span" }));
S.upsertSpan(S.createSpan("T-N-SPAN", "T", "T-N", "N", "", { type: "Fore Span" }));
S.addSpanPower({ spanId: "T-E-SPAN", poleId: "T", wireId: "T-PRIMARY-E", label: "Primary", attachmentHeight: "29'", midspan: "29'11\"" });
S.addSpanPower({ spanId: "T-N-SPAN", poleId: "T", wireId: "T-PRIMARY-N", label: "Primary", attachmentHeight: "29'", midspan: "29'3\"" });
S.addSpanPower({ spanId: "T-E-SPAN", poleId: "T", wireId: "T-SECONDARY-E", label: "Secondary", attachmentHeight: "29'", midspan: "25'" });
S.addSpanPower({ spanId: "T-N-SPAN", poleId: "T", wireId: "T-PRIMARY-LOW", label: "Primary", attachmentHeight: "28'", midspan: "26'" });
S.addSpanPower({ spanId: "T-E-SPAN", poleId: "T-E", wireId: "T-PRIMARY-E", label: "Primary", attachmentHeight: "29'", midspan: "29'11\"" });
S.addSpanPower({ spanId: "T-N-SPAN", poleId: "T-N", wireId: "T-PRIMARY-N", label: "Primary", attachmentHeight: "29'", midspan: "29'3\"" });
const localPower = S.getSpanPowerForPole("T");
const primaryEast = localPower.find(row => row.wireId === "T-PRIMARY-E");
const primaryNorth = localPower.find(row => row.wireId === "T-PRIMARY-N");
const secondaryEast = localPower.find(row => row.wireId === "T-SECONDARY-E");
const primaryLower = localPower.find(row => row.wireId === "T-PRIMARY-LOW");
const remoteEast = S.getSpanPowerForPole("T-E")[0];
S.updateSpanPowerField(primaryEast.key, "attachmentHeight", "28'10\"");
assert.equal(S.getState().spanPower[primaryNorth.key].attachmentHeight, "28'10\"", "same-height Primary runs must share one Attachment Height edit");
assert.equal(S.getState().spanPower[secondaryEast.key].attachmentHeight, "28'10\"", "all power types at the same pole-side height must share one Attachment Height edit");
S.updateSpanPowerField(primaryEast.key, "attachmentHeight", "29'");
S.updateSpanPowerField(primaryEast.key, "attachmentHeightChange", "30'");
C.recalculateSpansForPole("T");
assert.equal(S.getState().spanPower[primaryNorth.key].attachmentHeightChange, "30'", "same-height Primary runs leaving one pole must share the pole-side HOA Change");
assert.equal(S.getState().spanPower[secondaryEast.key].attachmentHeightChange, "30'", "a Secondary wire at the same pole-side height must join the shared-height group");
assert.equal(S.getState().spanPower[primaryLower.key].attachmentHeightChange, "", "a Primary at a different attachment height must remain independent");
assert.equal(S.getState().spanPower[remoteEast.key].attachmentHeightChange, "", "the matching Primary at the other pole must remain an independent endpoint movement");
assert.equal(S.getState().spanPower[primaryEast.key].calculatedMidspan, "30'5\"", "the first linked span must retain and calculate its own Midspan");
assert.equal(S.getState().spanPower[primaryNorth.key].calculatedMidspan, "29'9\"", "the second linked span must calculate its distinct Midspan");

S.resetState();
S.applyProjectProfile("INTEC");
S.upsertPole(S.createPole({ poleId: "UI", lowPower: "32'" }));
S.upsertPole(S.createPole({ poleId: "UI-E", lowPower: "32'" }));
S.upsertPole(S.createPole({ poleId: "UI-N", lowPower: "32'" }));
S.upsertPole(S.createPole({ poleId: "UI-W", lowPower: "32'" }));
S.upsertSpan(S.createSpan("UI-SPAN-E", "UI", "UI-E", "E", "", { type: "Fore Span" }));
S.upsertSpan(S.createSpan("UI-SPAN-N", "UI", "UI-N", "N", "", { type: "Fore Span" }));
S.upsertSpan(S.createSpan("UI-SPAN-W", "UI", "UI-W", "W", "", { type: "Fore Span" }));
S.addSpanPower({ spanId: "UI-SPAN-E", poleId: "UI", wireId: "UI-POWER-E", label: "Neutral", attachmentHeight: "22'10\"", attachmentHeightChange: "23'", midspan: "23'" });
S.addSpanPower({ spanId: "UI-SPAN-N", poleId: "UI", wireId: "UI-POWER-N", label: "Secondary", attachmentHeight: "22'10\"", attachmentHeightChange: "23'", midspan: "17'7\"" });
S.addSpanPower({ spanId: "UI-SPAN-W", poleId: "UI", wireId: "UI-POWER-W", label: "Neutral", attachmentHeight: "22'10\"", attachmentHeightChange: "23'", midspan: "21'3\"" });
const renderer = window.PoleWorkspaceRenderer.create({
  H: window.HeightUtils,
  S,
  escapeHtml: value => String(value ?? "").replaceAll("&", "&amp;").replaceAll('"', "&quot;").replaceAll("<", "&lt;").replaceAll(">", "&gt;"),
  poleResolutionOf: () => "AERIAL",
  spanRelationOf: span => S.getSpanRelation(span)
});
const powerHtml = renderer.renderPowerTable("UI");
assert.match(powerHtml, /<th>HOA Change<\/th><th>Midspan<\/th>/, "Power table must place HOA Change between Attachment Height and Midspan");
assert.equal((powerHtml.match(/data-field="attachmentHeight"/g) || []).length, 1, "different power types at one height should share one Attachment Height input");
assert.equal((powerHtml.match(/data-field="attachmentHeightChange"/g) || []).length, 1, "different power types at one height should share one HOA Change input");
assert.equal((powerHtml.match(/data-field="midspan"/g) || []).length, 3, "each span should retain its own Midspan field and calculation");
assert.equal((powerHtml.match(/data-delete-span-power/g) || []).length, 3, "each linked span should retain its own delete action");
assert.match(powerHtml, />UI-E<\/button>/, "first linked span should remain visible in the shared row");
assert.match(powerHtml, />UI-N<\/button>/, "second linked span should remain visible in the shared row");
assert.match(powerHtml, />UI-W<\/button>/, "third linked span should remain visible in the shared row");
assert.equal((powerHtml.match(/class="badge warning">Neutral<\/span>/g) || []).length, 2, "each Neutral power row should retain its own aligned type badge");
assert.match(powerHtml, /class="badge warning">Secondary<\/span>/, "the Secondary power row should retain its own aligned type badge");
assert.match(powerHtml, /data-span="UI-SPAN-E" data-field="midspan"/);
assert.match(powerHtml, /data-span="UI-SPAN-N" data-field="midspan"/);
assert.match(powerHtml, /data-span="UI-SPAN-W" data-field="midspan"/);
assert.match(powerHtml, /data-field="midspan" value="23'" title="Original Midspan before movements: 23'"/, "the first Midspan should remain separate and preserve its imported value on hover");
assert.match(powerHtml, /data-field="midspan" value="17'7&quot;" title="Original Midspan before movements: 17'7&quot;"/, "the second Midspan should display its distinct imported value");
assert.match(powerHtml, /data-field="midspan" value="21'3&quot;" title="Original Midspan before movements: 21'3&quot;"/, "the third Midspan should display its distinct imported value");
assert.doesNotMatch(powerHtml, /Adjusted:/, "Power Midspan should no longer render a separate Adjusted value");
assert.doesNotMatch(powerHtml, /Max Height at MS/, "Power table must omit the maximum-height column while retaining its internal calculations");

S.upsertPole(S.createPole({ poleId: "ORDER", lowPower: "32'" }));
S.upsertPole(S.createPole({ poleId: "ORDER-LOW", lowPower: "32'" }));
S.upsertPole(S.createPole({ poleId: "ORDER-HIGH", lowPower: "32'" }));
S.upsertSpan(S.createSpan("ORDER-SPAN-LOW", "ORDER", "ORDER-LOW", "E", "", { type: "Fore Span" }));
S.upsertSpan(S.createSpan("ORDER-SPAN-HIGH", "ORDER", "ORDER-HIGH", "N", "", { type: "Fore Span" }));
S.addSpanPower({ spanId: "ORDER-SPAN-LOW", poleId: "ORDER", wireId: "ORDER-LOW-WIRE", label: "Neutral", attachmentHeight: "22'10\"", midspan: "20'" });
S.addSpanPower({ spanId: "ORDER-SPAN-HIGH", poleId: "ORDER", wireId: "ORDER-HIGH-WIRE", label: "Primary", attachmentHeight: "30'", midspan: "28'" });
const orderedRows = renderer.renderPowerTable("ORDER").match(/<tbody>([\s\S]*?)<\/tbody>/)?.[1]
  ?.match(/<tr>([\s\S]*?)<\/tr>/g) || [];
assert.equal(orderedRows.length, 2, "different attachment heights should remain separate groups");
assert.match(orderedRows[0], /data-field="attachmentHeight" value="30'"/, "the highest attachment group should render first regardless of import order");
assert.match(orderedRows[1], /data-field="attachmentHeight" value="22'10&quot;"/, "lower attachment groups should render below higher groups");

console.log("Power wire movement tests passed.");
