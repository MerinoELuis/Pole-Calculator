(function (global) {
  "use strict";

  const DEFAULT_MESSENGER_144 = "0.242";
  const DEFAULT_FIBER_144 = "0.51";
  const SUPPORTED_PROFILES = new Set(["INTEC", "METRONET"]);

  const S = () => global.AppStore;
  const H = () => global.HeightUtils;

  function text(value) {
    return String(value ?? "").trim();
  }

  function escapeHtml(value) {
    return String(value ?? "")
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&#039;");
  }

  function fiberCount(value) {
    const match = text(value).match(/\b(\d+)\s*CT\b/i);
    return match ? String(Number(match[1])) : "";
  }

  function canonicalFiberName(count) {
    return count ? `${Number(count)}CT Fiber` : "";
  }

  function profileName(state) {
    return text(state?.settings?.projectProfile || "INTEC").toUpperCase();
  }

  function detectedReferenceFiberCounts(state) {
    const counts = new Set();
    (state?.makeReadyReferences || []).forEach(ref => {
      const count = fiberCount(`${ref?.attachmentFiber || ""} ${ref?.attachmentSizeRaw || ""}`);
      if (count) counts.add(count);
    });
    return counts;
  }

  function fiberEntries(state) {
    const entries = new Map();
    const configured = state?.settings?.fiberSizes && typeof state.settings.fiberSizes === "object"
      ? state.settings.fiberSizes
      : {};

    Object.keys(configured).forEach(name => {
      const count = fiberCount(name);
      if (count && !entries.has(count)) entries.set(count, name);
    });

    (state?.makeReadyReferences || []).forEach(ref => {
      const source = `${ref?.attachmentFiber || ""} ${ref?.attachmentSizeRaw || ""}`;
      const count = fiberCount(source);
      if (count && !entries.has(count)) entries.set(count, canonicalFiberName(count));
    });

    return Array.from(entries.entries())
      .map(([count, name]) => ({ count, name }))
      .sort((a, b) => Number(a.count) - Number(b.count));
  }

  function configuredFiberKey(state, count) {
    const configured = state?.settings?.fiberSizes && typeof state.settings.fiberSizes === "object"
      ? state.settings.fiberSizes
      : {};
    return Object.keys(configured).find(name => fiberCount(name) === String(Number(count))) || canonicalFiberName(count);
  }

  function applyAttachmentDefaults(state = S()?.getState?.()) {
    if (!state || !SUPPORTED_PROFILES.has(profileName(state))) return false;
    if (!detectedReferenceFiberCounts(state).has("144")) return false;

    state.settings = state.settings || {};
    state.settings.fiberSizes = state.settings.fiberSizes && typeof state.settings.fiberSizes === "object"
      ? state.settings.fiberSizes
      : {};

    let changed = false;
    if (!text(state.settings.attachmentMessengerSize)) {
      state.settings.attachmentMessengerSize = DEFAULT_MESSENGER_144;
      changed = true;
    }

    const key = configuredFiberKey(state, "144");
    if (!text(state.settings.fiberSizes[key])) {
      state.settings.fiberSizes[key] = DEFAULT_FIBER_144;
      changed = true;
    }
    return changed;
  }

  function numericSize(value) {
    const parsed = Number(text(value));
    return Number.isFinite(parsed) && parsed > 0 ? parsed : null;
  }

  function inches(value) {
    if (typeof value === "number" && Number.isFinite(value)) return Math.round(value);
    const parsed = H()?.parseHeight?.(value);
    return parsed === null || parsed === undefined || !Number.isFinite(Number(parsed))
      ? null
      : Math.round(Number(parsed));
  }

  function spanLengthInches(span) {
    const display = inches(span?.lengthDisplay || "");
    if (display !== null) return display;
    const raw = Number(span?.length);
    if (!Number.isFinite(raw) || raw <= 0) return null;
    return Math.round(raw * 12);
  }

  function ownerForMove(row) {
    const raw = text(row?.rawOwner || row?.ownerBase || row?.owner || "COMM");
    if (/century\s*link|centurylink|\bctl\b/i.test(raw)) return "CTL";
    return raw.replace(/^COMMUNICATION\s*>\s*/i, "").replace(/,\s*.*$/, "").trim() || "COMM";
  }

  function directionFromPole(span, poleId) {
    if (!span) return "";
    if (span.fromPole === poleId) return text(span.direction).toUpperCase();
    return ({ N: "S", NE: "SW", E: "W", SE: "NW", S: "N", SW: "NE", W: "E", NW: "SE" })[text(span.direction).toUpperCase()] || "";
  }

  function referenceForSpan(state, poleId, span) {
    const refs = (state.makeReadyReferences || []).filter(ref => ref.poleId === poleId);
    if (!refs.length) return null;
    const direction = directionFromPole(span, poleId);
    const matches = refs.filter(ref => {
      const tokens = Array.isArray(ref.attachmentDirectionTokens)
        ? ref.attachmentDirectionTokens.map(value => text(value).toUpperCase())
        : text(ref.attachmentDirection).split(/[\s/,;-]+/).map(value => value.toUpperCase()).filter(Boolean);
      return !direction || !tokens.length || tokens.includes(direction);
    });
    return matches[0] || refs[0] || null;
  }

  function fiberForSpan(state, poleId, span) {
    const ref = referenceForSpan(state, poleId, span);
    const fromRef = fiberCount(`${ref?.attachmentFiber || ""} ${ref?.attachmentSizeRaw || ""}`);
    if (fromRef) return fromRef;
    return fiberEntries(state)[0]?.count || "";
  }

  function spanKind(span) {
    const type = text(span?.type || span?.rawType).toLowerCase();
    if (/fore\s*span|forespan/.test(type)) return "F";
    if (/back\s*span|backspan/.test(type)) return "B";
    return "O";
  }

  function isUgSpan(state, span) {
    const targetId = text(span?.toPole);
    const targetPole = state?.poles?.[targetId] || S()?.getPole?.(targetId);
    const source = `${targetId} ${span?.rawType || ""} ${span?.notes || ""}`;
    return Boolean(targetPole?.ugActive || /(?:^|\s)UG(?:\s|$)/i.test(source));
  }

  function sideForSpan(state, poleId, spanId) {
    return Object.values(state.spanSides || {}).find(side => side.poleId === poleId && side.spanId === spanId) || null;
  }

  function primaryProposalForPole(state, poleId) {
    const sides = Object.values(state.spanSides || {})
      .filter(side => side.poleId === poleId && inches(side.proposedHOA) !== null)
      .sort((a, b) => Number(Boolean(a.isAdditionalProposed)) - Number(Boolean(b.isAdditionalProposed)));
    return sides[0] || null;
  }

  function bearingDegrees(span) {
    const raw = text(span?.bearingDegrees);
    const value = Number(raw);
    return raw && Number.isFinite(value) ? value : null;
  }

  function hasSpanGeometry(span) {
    return Boolean(text(span?.toPole) && (
      bearingDegrees(span) !== null || spanLengthInches(span) !== null || text(span?.direction)
    ));
  }

  function compactSpan(state, poleId, span, primaryProposal) {
    const item = {
      to: text(span.toPole),
      kind: spanKind(span)
    };
    const bearing = bearingDegrees(span);
    const length = spanLengthInches(span);
    if (bearing !== null) item.bearing = bearing;
    if (length !== null) item.length = length;

    if (isUgSpan(state, span)) {
      item.ug = true;
      return item;
    }

    const exactSide = sideForSpan(state, poleId, span.spanId);
    const proposal = inches(exactSide?.proposedHOA) !== null ? exactSide : primaryProposal;
    const hoa = inches(proposal?.proposedHOA);
    const fiber = fiberForSpan(state, poleId, span);
    if (hoa !== null && fiber) {
      item.hoa = hoa;
      item.fiber = Number(fiber);
    }

    if (proposal === exactSide) {
      const endDrop = inches(exactSide?.endDrop);
      const nextHoa = inches(exactSide?.proposedHOAChange);
      if (endDrop !== null) item.endDrop = endDrop;
      if (nextHoa !== null) item.nextHoa = nextHoa;
    }
    return item;
  }

  function movementIdentity(row, index, service) {
    const owner = ownerForMove(row).toLowerCase();
    const wireId = text(row?.wireId);
    if (wireId) return `${owner}|wire:${wireId}|service:${service}`;
    const wireIndex = text(row?.wireIndex);
    if (wireIndex) return `${owner}|index:${wireIndex}|service:${service}`;
    return `${owner}|row:${text(row?.spanId)}:${index}|service:${service}`;
  }

  function movementCandidates(state, poleId) {
    const candidates = [];
    const byIdentity = new Map();
    Object.values(state.spanComms || {}).forEach((row, index) => {
      if (row.poleId !== poleId) return;
      const from = inches(row.existingHOA);
      const to = inches(row.existingHOAChange || (row.transferToNewPole ? row.existingHOA : ""));
      if (from === null || to === null) return;
      if (from === to && !row.transferToNewPole) return;

      const service = Boolean(row.serviceDrop);
      const candidate = {
        owner: ownerForMove(row),
        from,
        to,
        service,
        dg: Boolean(row.downGuy),
        transfer: Boolean(row.transferToNewPole),
        identity: movementIdentity(row, index, service)
      };

      const existing = byIdentity.get(candidate.identity);
      if (existing) {
        existing.dg = existing.dg || candidate.dg;
        existing.transfer = existing.transfer || candidate.transfer;
        if (existing.to === existing.from && candidate.to !== candidate.from) existing.to = candidate.to;
        return;
      }
      byIdentity.set(candidate.identity, candidate);
      candidates.push(candidate);
    });
    return candidates;
  }

  function compactMoves(state, poleId) {
    return movementCandidates(state, poleId).map(candidate => {
      const item = { owner: candidate.owner, from: candidate.from, to: candidate.to };
      if (candidate.service) item.service = true;
      if (candidate.dg) item.dg = true;
      return item;
    });
  }

  function buildCompactPayload(state = S()?.getState?.()) {
    if (!state) return { sizes: { messenger: null, fiber: {} }, owner: "", poles: [] };
    applyAttachmentDefaults(state);

    const poles = [];
    Object.keys(state.poles || {}).sort((a, b) => a.localeCompare(b, undefined, { numeric: true })).forEach(poleId => {
      const primaryProposal = primaryProposalForPole(state, poleId);
      const spans = Object.values(state.spans || {})
        .filter(span => span.fromPole === poleId && hasSpanGeometry(span))
        .map(span => compactSpan(state, poleId, span, primaryProposal))
        .sort((a, b) => `${a.kind}|${a.to}`.localeCompare(`${b.kind}|${b.to}`, undefined, { numeric: true }));
      const moves = compactMoves(state, poleId);
      const terminalHoa = inches(state.poles[poleId]?.standaloneProposedHOA);
      if (!spans.length && !moves.length && terminalHoa === null) return;

      const pole = { id: poleId };
      if (terminalHoa !== null) pole.terminalHoa = terminalHoa;
      if (spans.length) pole.spans = spans;
      if (moves.length) pole.moves = moves;
      poles.push(pole);
    });

    const usedFibers = new Set();
    poles.forEach(pole => (pole.spans || []).forEach(span => {
      if (!span.ug && Number.isFinite(Number(span.fiber))) usedFibers.add(String(Number(span.fiber)));
    }));

    const fiberSizes = {};
    usedFibers.forEach(count => {
      const key = configuredFiberKey(state, count);
      fiberSizes[count] = numericSize(state.settings?.fiberSizes?.[key]);
    });

    return {
      sizes: {
        messenger: numericSize(state.settings?.attachmentMessengerSize),
        fiber: fiberSizes
      },
      owner: text(state.settings?.proposedOwner || "Wecom") || "Wecom",
      poles
    };
  }

  function validationErrors(payload) {
    const usedFibers = new Set();
    (payload?.poles || []).forEach(pole => (pole.spans || []).forEach(span => {
      if (!span.ug && Number.isFinite(Number(span.fiber))) usedFibers.add(String(Number(span.fiber)));
    }));
    if (!usedFibers.size) return [];

    const errors = [];
    if (numericSize(payload?.sizes?.messenger) === null) errors.push("Missing Messenger Size");
    Array.from(usedFibers).sort((a, b) => Number(a) - Number(b)).forEach(count => {
      if (numericSize(payload?.sizes?.fiber?.[count]) === null) errors.push(`Missing ${count}CT Fiber Size`);
    });
    return errors;
  }

  function safeJobFilePart(value) {
    const raw = text(value || "pole_job")
      .replace(/\.[^.]+$/, "")
      .replace(/_(?:Pole_Calculator|AutoProposed|Debug)$/i, "")
      .replace(/^excel[_\-\s]*/i, "")
      .replace(/[_\-\s]*20\d{2}[-_]\d{2}[-_]\d{2}$/i, "")
      .trim();
    return (raw || "pole_job")
      .replace(/[<>:"/\\|?*\u0000-\u001F]+/g, "_")
      .replace(/\s+/g, "_");
  }

  function notify(message, type = "warning") {
    if (global.document) {
      const host = global.document.getElementById("toastHost");
      if (host) {
        const item = global.document.createElement("div");
        item.className = `toast ${type}`;
        item.textContent = message;
        host.appendChild(item);
        global.setTimeout?.(() => item.remove(), 4200);
        return;
      }
    }
    global.alert?.(message);
  }

  function exportCompactProposedJson() {
    global.Calculations?.recalculateAll?.();
    const state = S()?.getState?.();
    const payload = buildCompactPayload(state);
    const errors = validationErrors(payload);
    if (errors.length) {
      errors.forEach(error => notify(error, "warning"));
      return false;
    }
    const jobName = safeJobFilePart(state?.jobName || state?.importedFileName || "pole_job");
    global.ProjectExport?.downloadJson?.(`${jobName}_AutoProposed.json`, payload);
    return true;
  }

  function movementMrLine(candidate) {
    if (candidate.transfer) return "";
    const from = H()?.formatHeight?.(candidate.from) || String(candidate.from);
    const to = H()?.formatHeight?.(candidate.to) || String(candidate.to);
    const dg = candidate.dg ? " with DG" : "";
    if (candidate.service) return `Relocate ${candidate.owner} drop at HOA ${from} to HOA ${to}${dg}.`;
    const verb = candidate.to > candidate.from ? "raise" : "lower";
    return `At HOA ${from} ${verb} ${candidate.owner} to HOA ${to}${dg}.`;
  }

  function applyMrCase(value, state) {
    return text(state?.settings?.mrCase).toUpperCase() === "UPPER" ? value.toUpperCase() : value;
  }

  function augmentPoleMakeReady(poleId) {
    const state = S()?.getState?.();
    if (!state) return [];
    const generated = movementCandidates(state, poleId).map(movementMrLine).filter(Boolean).map(line => applyMrCase(line, state));
    if (!generated.length) return state.mr?.filter(item => item.poleId === poleId) || [];

    state.mr = Array.isArray(state.mr) ? state.mr : [];
    let item = state.mr.find(row => row.poleId === poleId);
    if (!item) {
      item = { poleId, spanId: "", owner: "MR", text: "", imported: false };
      state.mr.push(item);
    }
    const lines = text(item.text).split(/\n+/).map(line => line.trim()).filter(Boolean);
    const seen = new Set(lines.map(line => line.toLowerCase()));
    generated.forEach(line => {
      const key = line.toLowerCase();
      if (seen.has(key)) return;
      const withoutDg = key.replace(/ with dg\.$/, ".");
      const baseIndex = lines.findIndex(existing => existing.toLowerCase() === withoutDg);
      if (baseIndex >= 0 && withoutDg !== key) {
        seen.delete(lines[baseIndex].toLowerCase());
        lines[baseIndex] = line;
        seen.add(key);
        return;
      }
      seen.add(key);
      lines.push(line);
    });
    item.text = lines.join("\n");
    return state.mr.filter(row => row.poleId === poleId);
  }

  function patchMrLogic() {
    const logic = global.MRLogic;
    if (!logic || logic.__compactAutoProposedPatched || typeof logic.generateMRForPole !== "function") return;
    const originalGenerateMRForPole = logic.generateMRForPole.bind(logic);
    logic.generateMRForPole = function (poleId) {
      originalGenerateMRForPole(poleId);
      return augmentPoleMakeReady(poleId);
    };
    logic.generateAllMR = function () {
      const state = S()?.getState?.();
      if (!state) return [];
      state.mr = [];
      Object.keys(state.poles || {}).forEach(poleId => logic.generateMRForPole(poleId));
      return state.mr;
    };
    logic.__compactAutoProposedPatched = true;
  }

  function markSaveDirty() {
    if (!global.document) return;
    const button = global.document.getElementById("saveLocalBtn");
    if (button && !/\*$/.test(button.textContent || "")) button.textContent = `${text(button.textContent) || "Save"} *`;
  }

  function updateInjectedSetting(input) {
    const state = S()?.getState?.();
    if (!state) return;
    state.settings = state.settings || {};
    if (input.dataset.field === "attachmentMessengerSize") {
      state.settings.attachmentMessengerSize = text(input.value);
    } else if (input.dataset.field === "fiberSize" && input.dataset.fiber) {
      state.settings.fiberSizes = state.settings.fiberSizes && typeof state.settings.fiberSizes === "object" ? state.settings.fiberSizes : {};
      state.settings.fiberSizes[input.dataset.fiber] = text(input.value);
    }
    markSaveDirty();
  }

  function createFiberSection(state, fibers) {
    const section = global.document.createElement("div");
    section.className = "settings-section settings-section-fiber";
    section.dataset.compactFiberSection = "true";
    section.innerHTML = `<h3>Fiber</h3><div class="settings-grid fiber-grid">
      <label class="clearance-row"><span>Messenger Size</span><input class="input" data-scope="attachmentSettings" data-field="attachmentMessengerSize" value="${escapeHtml(state.settings?.attachmentMessengerSize || "")}" /></label>
      ${fibers.map(item => {
        const key = configuredFiberKey(state, item.count);
        return `<label class="clearance-row"><span>${escapeHtml(item.name)} Size</span><input class="input" data-scope="attachmentSettings" data-field="fiberSize" data-fiber="${escapeHtml(key)}" value="${escapeHtml(state.settings?.fiberSizes?.[key] || "")}" /></label>`;
      }).join("")}
    </div>`;
    section.querySelectorAll("input[data-scope='attachmentSettings']").forEach(input => {
      input.addEventListener("input", () => updateInjectedSetting(input));
      input.addEventListener("change", () => updateInjectedSetting(input));
    });
    return section;
  }

  let syncingUi = false;
  function syncFiberSettingsUi() {
    if (syncingUi || !global.document || !S()?.getState) return;
    syncingUi = true;
    try {
      const state = S().getState();
      const changed = applyAttachmentDefaults(state);
      if (changed) markSaveDirty();
      const container = global.document.getElementById("clearanceSettings");
      if (!container) return;
      const fibers = fiberEntries(state);
      const supported = SUPPORTED_PROFILES.has(profileName(state));
      let section = container.querySelector(".settings-section-fiber");
      if (!supported || !fibers.length) {
        if (section?.dataset.compactFiberSection === "true") section.remove();
        return;
      }
      if (!section) {
        section = createFiberSection(state, fibers);
        container.appendChild(section);
      }
      const messenger = section.querySelector('[data-field="attachmentMessengerSize"]');
      if (messenger && global.document.activeElement !== messenger) messenger.value = state.settings?.attachmentMessengerSize || "";
      fibers.forEach(item => {
        const key = configuredFiberKey(state, item.count);
        const input = Array.from(section.querySelectorAll('[data-field="fiberSize"]')).find(row => fiberCount(row.dataset.fiber) === item.count);
        if (input && global.document.activeElement !== input) input.value = state.settings?.fiberSizes?.[key] || "";
      });
    } finally {
      syncingUi = false;
    }
  }

  function patchStore() {
    const store = S();
    if (!store || store.__compactAutoProposedPatched) return;
    if (typeof store.setState === "function") {
      const originalSetState = store.setState.bind(store);
      store.setState = function (nextState) {
        const result = originalSetState(nextState);
        applyAttachmentDefaults(store.getState?.());
        return result;
      };
    }
    if (typeof store.updateSetting === "function") {
      const originalUpdateSetting = store.updateSetting.bind(store);
      store.updateSetting = function (field, value) {
        const result = originalUpdateSetting(field, value);
        if (field === "projectProfile") applyAttachmentDefaults(store.getState?.());
        return result;
      };
    }
    store.__compactAutoProposedPatched = true;
  }

  function installUiSync() {
    if (!global.document) return;
    const start = () => {
      syncFiberSettingsUi();
      const container = global.document.getElementById("clearanceSettings");
      if (container && global.MutationObserver) {
        const observer = new global.MutationObserver(() => {
          if (global.queueMicrotask) global.queueMicrotask(syncFiberSettingsUi);
          else global.setTimeout?.(syncFiberSettingsUi, 0);
        });
        observer.observe(container, { childList: true, subtree: true });
      }
      global.document.addEventListener("input", event => {
        if (event.target?.dataset?.scope === "attachmentSettings") markSaveDirty();
      });
      global.document.addEventListener("change", event => {
        if (event.target?.dataset?.field === "projectProfile" || event.target?.dataset?.scope === "attachmentSettings") {
          global.setTimeout?.(syncFiberSettingsUi, 0);
        }
      });
    };
    if (global.document.readyState === "loading") global.document.addEventListener("DOMContentLoaded", start);
    else start();
  }

  patchStore();
  patchMrLogic();
  applyAttachmentDefaults();
  if (global.ProjectExport) global.ProjectExport.exportProposedJson = exportCompactProposedJson;
  installUiSync();

  global.CompactAutoProposed = {
    applyAttachmentDefaults,
    detectedReferenceFiberCounts,
    fiberEntries,
    buildCompactPayload,
    validationErrors,
    compactMoves,
    movementCandidates,
    augmentPoleMakeReady,
    exportCompactProposedJson,
    isUgSpan,
    spanKind,
    spanLengthInches
  };
})(window);
