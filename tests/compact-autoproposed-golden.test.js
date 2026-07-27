"use strict";

const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");

function parseHeight(value) {
  if (typeof value === "number" && Number.isFinite(value)) return value;
  const raw = String(value ?? "").trim();
  if (!raw) return null;
  if (/^-?\d+(?:\.\d+)?$/.test(raw)) return Number(raw);
  const sign = raw.startsWith("-") ? -1 : 1;
  const clean = raw.replace(/^-/, "");
  const feet = Number(clean.match(/(\d+)\s*'/)?.[1] || 0);
  const inches = Number(clean.match(/(\d+)\s*"/)?.[1] || 0);
  return feet || inches ? sign * (feet * 12 + inches) : null;
}

const root = path.join(__dirname, "..");
const state = JSON.parse(fs.readFileSync(path.join(__dirname, "fixtures", "compact-normal-state.json"), "utf8"));
const expected = JSON.parse(fs.readFileSync(path.join(__dirname, "expected", "compact-normal-output.json"), "utf8"));
const window = {
  AppStore: {
    getState: () => state,
    getPole: poleId => state.poles[poleId] || null,
    canonicalPoleIdentity: value => String(value || "").trim().toUpperCase(),
    setState() {},
    updateSetting() {}
  },
  HeightUtils: {
    parseHeight,
    formatHeight: value => String(value)
  },
  Calculations: { recalculateAll() {} },
  ProjectExport: { downloadJson() {} },
  MRLogic: { generateMRForPole() { return []; }, generateAllMR() { return []; } },
  console,
  setTimeout,
  queueMicrotask,
  alert() {}
};

const source = fs.readFileSync(path.join(root, "js", "compact-autoproposed.js"), "utf8");
vm.runInNewContext(source, { window }, { filename: "compact-autoproposed.js" });
const actual = window.CompactAutoProposed.buildCompactPayload(state);
assert.deepStrictEqual(JSON.parse(JSON.stringify(actual)), expected);

console.log("compact-autoproposed golden fixture test passed");
