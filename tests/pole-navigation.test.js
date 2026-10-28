"use strict";

const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");

const calls = [];
const state = {
  ui: { activeView: "calculator", hiddenPoleIds: ["Unknown-P2"], revealedUnknownPoleIds: [] },
  selectedPoleId: "",
  selectedSpanId: ""
};
const card = { scrollIntoView: options => calls.push(["scroll", options]) };
const document = {
  body: { classList: { toggle: (...args) => calls.push(["bodyClass", ...args]) } },
  activeElement: { blur: () => calls.push(["blur"]) },
  querySelector: selector => (calls.push(["query", selector]), card)
};
const window = { CSS: { escape: value => String(value).replaceAll(" ", "\\ ") } };
const modulePath = path.join(__dirname, "..", "js", "ui", "pole-navigation.js");
vm.runInNewContext(fs.readFileSync(modulePath, "utf8"), { window }, { filename: modulePath });

const elements = {
  poleIndexDrawer: { classList: { toggle: (...args) => calls.push(["drawerClass", ...args]) }, setAttribute() {} },
  poleIndexToggle: { classList: { add: (...args) => calls.push(["toggleAdd", ...args]), remove: (...args) => calls.push(["toggleRemove", ...args]) }, setAttribute() {} },
  poleIndexBackdrop: { classList: { toggle: (...args) => calls.push(["backdropClass", ...args]) } }
};
const navigation = window.PoleNavigation.create({
  document,
  els: elements,
  mobileKeyboard: { setOpen: open => calls.push(["keyboard", open]) },
  recordUndoSnapshot: () => calls.push(["undo"]),
  render: () => calls.push(["render"]),
  S: { getState: () => state }
});

navigation.setOpen(true);
assert.ok(calls.some(call => call[0] === "keyboard" && call[1] === false));
assert.ok(calls.some(call => call[0] === "blur"));
navigation.updateToggleVisibility();
assert.ok(calls.some(call => call[0] === "toggleAdd" && call[1] === "visible"));

navigation.selectPole("Unknown-P2");
assert.deepEqual(Array.from(state.ui.hiddenPoleIds), []);
assert.deepEqual(Array.from(state.ui.revealedUnknownPoleIds), ["Unknown-P2"]);
assert.equal(state.selectedPoleId, "Unknown-P2");
assert.equal(state.selectedSpanId, "");
assert.equal(state.ui.activeView, "calculator");
assert.ok(calls.some(call => call[0] === "undo"));
assert.ok(calls.some(call => call[0] === "render"));
assert.ok(calls.some(call => call[0] === "scroll"));

console.log("Pole navigation tests passed");
