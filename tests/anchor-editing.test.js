"use strict";

const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");

const window = { localStorage: { getItem: () => null, setItem: () => {} } };
const sandbox = { window, console, Date, Set, Map, Math, JSON, Object, Array, String, Number, Boolean, RegExp };
for (const file of ["domain/height-math.js", "core/project-config.js", "domain/pole-resolution.js", "core/state-compat.js", "core/store-identity.js", "core/store.js"]) {
  const modulePath = path.join(__dirname, "..", "js", file);
  vm.runInNewContext(fs.readFileSync(modulePath, "utf8"), sandbox, { filename: modulePath });
}

const S = window.AppStore;
S.resetState();
S.upsertPole(S.createPole({
  poleId: "P1",
  metadata: {
    anchors: [{
      anchorId: "A1",
      anchorIndex: "1",
      owner: "UTILITY > APS",
      leadLength: "18'",
      bearing: "90°",
      groundlineViolation: false,
      manualFields: [],
      guys: [{ guyId: "G1", guyIndex: "1", size: "EHS 5/16 DG", owner: "UTILITY > APS", attachmentHeight: "23'", manualFields: [] }]
    }]
  }
}));

S.updateAnchorField("P1", 0, "leadLength", "19'6\"");
S.updateAnchorField("P1", 0, "bearing", "180°");
S.updateAnchorField("P1", 0, "groundlineViolation", true);
S.updateAnchorField("P1", 0, "attachmentHeight", "24'", 0);

const anchor = S.getPole("P1").metadata.anchors[0];
assert.equal(anchor.leadLength, "19'6\"", "lead length should be editable");
assert.equal(anchor.bearing, "180°", "bearing should be editable");
assert.equal(anchor.groundlineViolation, true, "the groundline violation should be an editable boolean");
assert.equal(anchor.guys[0].attachmentHeight, "24'", "the associated Guy HOA should be editable");
assert.deepEqual(Array.from(anchor.manualFields).sort(), ["bearing", "groundlineViolation", "leadLength"].sort());
assert.deepEqual(Array.from(anchor.guys[0].manualFields), ["attachmentHeight"]);
assert.equal(S.poleHasChanges("P1"), true, "manual ANC/Guys edits should count as saved pole work");

const secondANC = S.addProposedANC("P1", { ancDistance: "18'", dgHoa: "22'", ancSize: '10"' });
assert.ok(secondANC?.id, "adding an extra proposed ANC should create a persisted row");
S.updateProposedANCField("P1", 0, "ancDirection", "w");
S.updateProposedANCField("P1", 0, "dgHoa", "23'6\"");
assert.equal(S.getPole("P1").metadata.proposedANCs[0].ancDirection, "W", "proposed ANC directions should normalize to uppercase");
assert.equal(S.getPole("P1").metadata.proposedANCs[0].dgHoa, "23'6\"");
assert.deepEqual(Array.from(S.getPole("P1").metadata.proposedANCs[0].manualFields).sort(), ["ancDirection", "dgHoa"].sort());
assert.equal(S.poleHasChanges("P1"), true, "additional proposed anchors should count as saved pole work");

const serialized = JSON.stringify(S.getState());
S.setState(JSON.parse(serialized));
assert.equal(S.getPole("P1").metadata.anchors[0].groundlineViolation, true, "ANC edits should survive save/load state roundtrips");
assert.equal(S.getPole("P1").metadata.anchors[0].guys[0].attachmentHeight, "24'", "Guy edits should survive save/load state roundtrips");
assert.equal(S.getPole("P1").metadata.proposedANCs[0].dgHoa, "23'6\"", "additional proposed ANC edits should survive save/load state roundtrips");
assert.equal(S.removeProposedANC("P1", 0).ancDirection, "W", "the selected additional proposed ANC should be removable");
assert.equal(S.getPole("P1").metadata.proposedANCs.length, 0, "removing the additional ANC should leave the original ANC data untouched");

const rendererPath = path.join(__dirname, "..", "js", "ui", "poles", "anchor-renderer.js");
vm.runInNewContext(fs.readFileSync(rendererPath, "utf8"), sandbox, { filename: rendererPath });
const markup = window.AnchorRenderer.create({
  S,
  escapeHtml: value => String(value ?? "").replace(/[&<>\"]/g, char => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", "\"": "&quot;" })[char]),
  renderCollapseEyeIcon: () => "<span class=\"eye\"></span>"
}).renderAnchorSection("P1");
assert.match(markup, /<details class="subsection wide anchor-section" open>/, "ANC disclosure should start expanded when anchor data exists");
assert.match(markup, /Groundline Violation/, "the editable groundline toggle should be visible in ANC details");
assert.match(markup, /EHS 5\/16 DG/, "associated Guys should be listed in the same anchor section");
assert.doesNotMatch(markup, /Anchor Type|Pitch|Anchor ·/, "ANC details should omit fields and counts the operator excluded");
const emptyMarkup = window.AnchorRenderer.create({
  S,
  escapeHtml: value => String(value ?? ""),
  renderCollapseEyeIcon: () => "<span class=\"eye\"></span>"
}).renderAnchorSection("NO-ANC");
assert.match(emptyMarkup, /<details class="subsection wide anchor-section">/, "an ANC-free pole should keep its section collapsed");

console.log("Anchor/Guys editing tests passed.");
