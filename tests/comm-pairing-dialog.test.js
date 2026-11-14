"use strict";

const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");

const local = {
  spanId: "span-1",
  poleId: "P1",
  owner: "Uniti",
  ownerCanonical: "UNITI",
  ownerDisplay: "Uniti",
  wireId: "wire-local",
  size: "Telco > 1\" Messenger > Static",
  existingHOA: "20'6\"",
  midspan: "19'2\""
};
const remote = {
  spanId: "span-1",
  poleId: "P2",
  owner: "Uniti",
  ownerCanonical: "UNITI",
  ownerDisplay: "Uniti",
  wireId: "wire-remote",
  size: "Telco > 1\" Messenger > Static",
  existingHOA: "17'4\"",
  midspan: "16'"
};
const span = { spanId: "span-1", fromPole: "P1", toPole: "P2", relation: "Fore Span" };
const calls = [];
let changeHandler = null;
let remoteMidspan = { textContent: "" };
let select = null;
let options = null;

const store = {
  getConnectedSpans: poleId => poleId === "P1" ? [span] : [],
  getOtherPoleId: (_span, poleId) => poleId === "P1" ? "P2" : "P1",
  getSpanRelation: value => value?.relation || "",
  getOwnerCanonical: row => row.ownerCanonical,
  getOwnerDisplay: row => row.ownerDisplay,
  keyForSpanComm: (spanId, poleId, owner, wireId) => [spanId, poleId, owner, wireId].join("|")
};
const calculations = {
  getCommPairingWarnings: () => [],
  getCommPairingRows: () => ({ localRows: [local], remoteRows: [remote] }),
  getRemoteCommMatchResult: () => ({ status: "NO_MATCH", row: null }),
  getCommMatchMidspanDisplay: row => row?.midspan || "",
  canMatchCommRows: () => true
};
const form = {
  addEventListener: (_type, handler) => { changeHandler = handler; },
  querySelectorAll: selector => selector === "[data-match-local]" ? [select] : [],
  querySelector(selector) {
    if (selector === '[data-match-local="0"]') return select;
    if (selector === '[data-match-remote-midspan="0"]') return remoteMidspan;
    if (selector === "[data-match-table-body]") return { innerHTML: "" };
    return null;
  }
};
let dialogResult = null;
async function openAppDialog(dialogOptions) {
  calls.push(dialogOptions);
  options = dialogOptions;
  if (dialogOptions.onOpen) {
    dialogOptions.onOpen({ form });
    changeHandler?.({
      target: {
        matches: selector => selector === "[data-match-local]",
        dataset: { matchLocal: "0" }
      }
    });
  }
  return dialogResult === null ? dialogOptions.getResult(form) : dialogResult;
}

const window = {};
const sourcePath = path.join(__dirname, "..", "js", "ui", "comm-pairing-dialog.js");
vm.runInNewContext(fs.readFileSync(sourcePath, "utf8"), { window }, { filename: sourcePath });

async function run() {
  select = { value: "0", options: [{ value: "" }, { value: "0" }] };
  remoteMidspan = { textContent: "" };
  const dialog = window.CommPairingDialogModule.create({
    S: store,
    H: { parseHeight: value => Number.parseInt(value, 10) || 0 },
    getCalculations: () => calculations,
    openAppDialog,
    escapeHtml: value => String(value)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&#039;")
  });

  const matches = await dialog.open("P1");
  assert.equal(calls.length, 1);
  assert.equal(options.title, "Match Cables");
  assert.match(options.contentHtml, /Fore Span · P1 → P2/);
  assert.match(options.contentHtml, /Uniti · Telco &gt; 1&quot; Messenger/);
  assert.match(options.contentHtml, /Match cable/);
  assert.equal(remoteMidspan.textContent, remote.midspan, "changing a pairing should refresh the remote Midspan display");
  assert.deepEqual(JSON.parse(JSON.stringify(matches)), [{
    otherPoleId: "P2",
    spanId: "span-1",
    pairs: [{
      local: { spanId: "span-1", poleId: "P1", ownerCanonical: "UNITI", wireId: "wire-local" },
      remote: { spanId: "span-1", poleId: "P2", ownerCanonical: "UNITI", wireId: "wire-remote" }
    }]
  }]);

  const noSpanDialog = window.CommPairingDialogModule.create({
    S: { ...store, getConnectedSpans: () => [] },
    H: { parseHeight: () => null },
    getCalculations: () => calculations,
    openAppDialog: async dialogOptions => { calls.push(dialogOptions); return []; },
    escapeHtml: String
  });
  assert.equal(await noSpanDialog.open("P9"), null, "a pole without connected spans should not produce match data");
  assert.match(calls[1].contentHtml, /no connected spans yet/);

  console.log("Comm pairing dialog tests passed");
}

run().catch(error => {
  console.error(error);
  process.exitCode = 1;
});
