const assert = require("assert");
const fs = require("fs");
const vm = require("vm");
const path = require("path");

const source = fs.readFileSync(
  path.join(__dirname, "../js/span-color-reset.js"),
  "utf8"
);
const window = {
  document: null,
  setTimeout,
  queueMicrotask
};
vm.runInNewContext(source, { window, console, Set, Map, Number, String, Array });

const spans = {
  "A-1": { spanId: "A-1", spanIndex: "1" },
  "A-2": { spanId: "A-2", spanIndex: "2" },
  "A-hidden": { spanId: "A-hidden", spanIndex: "0" },
  "B-1": { spanId: "B-1", spanIndex: "1" },
  "B-2": { spanId: "B-2", spanIndex: "2" },
  "manual-B": { spanId: "manual-B", sourceSpanId: "B-2", spanIndex: "99" }
};
const connected = {
  P01: [spans["A-hidden"], spans["A-1"], spans["A-2"]],
  P02: [spans["B-1"], spans["B-2"]]
};
const store = {
  getSpan: id => spans[id] || null,
  getConnectedSpans: poleId => connected[poleId] || []
};

const helper = window.SpanColorReset;

assert.deepStrictEqual(
  Array.from(helper.sortVisibleSpanIds(store, "P01", ["A-2", "A-1"])),
  ["A-1", "A-2"],
  "Hidden spans must not consume a visible color position."
);
assert.strictEqual(helper.classForVisibleIndex(0), "span-color-0");
assert.strictEqual(helper.classForVisibleIndex(1), "span-color-1");
assert.strictEqual(helper.classForVisibleIndex(2), "span-color-2");
assert.strictEqual(helper.classForVisibleIndex(5), "span-color-0");

const p01 = helper.sortVisibleSpanIds(store, "P01", ["A-2", "A-1"]);
const p02 = helper.sortVisibleSpanIds(store, "P02", ["manual-B", "B-1"]);
assert.strictEqual(helper.classForVisibleIndex(p01.indexOf("A-1")), "span-color-0");
assert.strictEqual(helper.classForVisibleIndex(p02.indexOf("B-1")), "span-color-0");
assert.strictEqual(helper.physicalSpanId(store, "manual-B"), "B-2");

function spanRow(spanId) {
  const marker = { dataset: { span: spanId } };
  const dot = { id: `dot-${spanId}` };
  return {
    querySelector(selector) {
      return selector === "[data-span]" ? marker : null;
    },
    querySelectorAll(selector) {
      return selector === ".span-color-dot" ? [dot] : [];
    }
  };
}

function blankMidspan(id) {
  const display = { id: `display-${id}` };
  return {
    id,
    dataset: {},
    querySelector() {
      return null;
    },
    querySelectorAll(selector) {
      return selector === ".midspan-highlight-input, .midspan-highlight-display"
        ? [display]
        : [];
    }
  };
}

const blueSpanRow = spanRow("B-2");
const blueBlankMidspan = blankMidspan("blank-blue-midspan");
const groupedCommTableRow = {
  querySelectorAll(selector) {
    if (selector === ".comm-span-list .comm-span-row") {
      return [spanRow("B-1"), blueSpanRow];
    }
    if (
      selector ===
      ".comm-midspan-list .comm-midspan-value.colored-midspan"
    ) {
      return [blankMidspan("first-midspan"), blueBlankMidspan];
    }
    return [];
  }
};

const rowPairs = helper.collectCommRowPairs(groupedCommTableRow);
assert.strictEqual(rowPairs.length, 2);
assert.strictEqual(rowPairs[1].spanId, "B-2");
assert.strictEqual(
  rowPairs[1].midspanContainer,
  blueBlankMidspan,
  "A blank REF Midspan must inherit the color of the span row at the same index."
);
assert.strictEqual(
  helper.firstSpanId({ dataset: { span: "B-2" } }),
  "B-2",
  "Elements carrying data-span directly must be supported."
);

console.log("span color reset tests passed");
