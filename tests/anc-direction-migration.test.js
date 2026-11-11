"use strict";

const assert = require("node:assert/strict");

global.window = global;
global.localStorage = { getItem: () => null, setItem: () => {} };
global.Calculations = { recalculateAll: () => {} };

require("../js/domain/height-math.js");
require("../js/core/project-config.js");
require("../js/domain/pole-resolution.js");
require("../js/core/state-compat.js");
require("../js/core/store-identity.js");
require("../js/core/store.js");

const normalized = AppStore.normalizeState({
  settings: { projectProfile: "INTEC", proposedOwner: "Wecom" },
  poles: {
    P1: { poleId: "P1", ancActive: true, ancDirection: "W", ancDirectionManual: false },
    P2: { poleId: "P2", ancActive: true, ancDirection: "", ancDirectionManual: false }
  }
});

assert.equal(normalized.settings.proposedOwner, "Wecom", "an existing INTEC proposed-owner selection must be preserved");
assert.equal(normalized.poles.P1.ancDirectionManual, true, "a populated saved ANC direction should be restored as the selected direction");
assert.equal(normalized.poles.P2.ancDirectionManual, false, "an empty ANC direction should remain unselected");

AppStore.setState({ settings: { projectProfile: "INTEC", proposedOwner: "Wecom" } });
AppStore.updateSetting("proposedOwner", "Cable One");
assert.equal(AppStore.getState().settings.proposedOwner, "Cable One", "INTEC must retain the owner chosen from its selector");

console.log("ANC direction migration tests passed.");
