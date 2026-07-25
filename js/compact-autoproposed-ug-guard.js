(function (global) {
  "use strict";

  const api = global.CompactAutoProposed;
  const projectExport = global.ProjectExport;
  const store = global.AppStore;

  if (!api || !projectExport || !store || projectExport.__compactAutoProposedUgGuardPatched) return;

  function text(value) {
    return String(value ?? "").trim();
  }

  function hasUgToken(value) {
    return /(?:^|\s)UG(?:\s|$)/i.test(text(value));
  }

  function hasPcoToken(value) {
    return /(?:^|\s)PCO(?:\s|$)/i.test(text(value));
  }

  function findPole(state, poleId) {
    if (!poleId) return null;
    if (state?.poles?.[poleId]) return state.poles[poleId];

    const canonical = store.canonicalPoleIdentity?.(poleId);
    if (canonical) {
      const match = Object.values(state?.poles || {}).find(pole =>
        store.canonicalPoleIdentity?.(pole?.poleId || "") === canonical);
      if (match) return match;
    }

    return store.getPole?.(poleId) || null;
  }

  function isPoleFullyUg(state, poleId) {
    if (typeof api.isPoleFullyUg === "function") return api.isPoleFullyUg(state, poleId);
    const pole = findPole(state, poleId);
    return Boolean(pole?.ugActive || hasUgToken(poleId));
  }

  function isPolePco(state, poleId) {
    const pole = findPole(state, poleId);
    return Boolean(pole?.pcoActive || hasPcoToken(poleId));
  }

  function blocksLocalActions(state, poleId) {
    return isPoleFullyUg(state, poleId) || isPolePco(state, poleId);
  }

  function geometryOnlySpan(span, markUg) {
    const item = {
      to: span.to,
      kind: span.kind
    };
    if (Object.prototype.hasOwnProperty.call(span, "bearing")) item.bearing = span.bearing;
    if (Object.prototype.hasOwnProperty.call(span, "length")) item.length = span.length;
    if (markUg) item.ug = true;
    return item;
  }

  function sanitizeBlockedLocalPole(pole, state) {
    if (!pole) return pole;

    const fullyUg = isPoleFullyUg(state, pole.id);
    const pco = !fullyUg && isPolePco(state, pole.id);
    if (!fullyUg && !pco) return pole;

    const sanitized = { id: pole.id };
    if (Array.isArray(pole.spans) && pole.spans.length) {
      sanitized.spans = pole.spans.map(span => geometryOnlySpan(span, fullyUg));
    }
    return sanitized;
  }

  function hasExportablePoleData(pole) {
    return Boolean(
      pole && (
        (Array.isArray(pole.spans) && pole.spans.length) ||
        (Array.isArray(pole.moves) && pole.moves.length) ||
        Object.prototype.hasOwnProperty.call(pole, "terminalHoa")
      )
    );
  }

  function sanitizeBlockedLocalPayload(payload, state = store.getState?.()) {
    if (!payload || !Array.isArray(payload.poles)) return payload;
    return {
      ...payload,
      poles: payload.poles
        .map(pole => sanitizeBlockedLocalPole(pole, state))
        .filter(hasExportablePoleData)
    };
  }

  function movementMrLine(candidate) {
    if (!candidate || candidate.transfer) return "";
    const formatter = global.HeightUtils?.formatHeight;
    const from = typeof formatter === "function" ? formatter(candidate.from) : String(candidate.from);
    const to = typeof formatter === "function" ? formatter(candidate.to) : String(candidate.to);
    const owner = text(candidate.owner || "COMM") || "COMM";
    const dg = candidate.dg ? " with DG" : "";
    if (candidate.service) return `Relocate ${owner} drop at HOA ${from} to HOA ${to}${dg}.`;
    const verb = candidate.to > candidate.from ? "raise" : "lower";
    return `At HOA ${from} ${verb} ${owner} to HOA ${to}${dg}.`;
  }

  function normalizeMrLine(value) {
    return text(value).replace(/\s+/g, " ").toLowerCase();
  }

  function generatedMovementLineKeys(state, poleId) {
    if (typeof api.movementCandidates !== "function") return new Set();
    return new Set(
      api.movementCandidates(state, poleId)
        .map(movementMrLine)
        .map(normalizeMrLine)
        .filter(Boolean)
    );
  }

  function sanitizeBlockedLocalMr(state, poleId) {
    if (!state || !Array.isArray(state.mr) || !blocksLocalActions(state, poleId)) {
      return state?.mr?.filter(item => item.poleId === poleId) || [];
    }

    const generatedMovementLines = generatedMovementLineKeys(state, poleId);
    state.mr = state.mr.reduce((items, item) => {
      if (item?.poleId !== poleId) {
        items.push(item);
        return items;
      }

      const remainingLines = text(item.text)
        .split(/\r?\n/)
        .map(line => line.trim())
        .filter(Boolean)
        .filter(line => !generatedMovementLines.has(normalizeMrLine(line)));

      if (remainingLines.length) items.push({ ...item, text: remainingLines.join("\n") });
      return items;
    }, []);

    return state.mr.filter(item => item.poleId === poleId);
  }

  const originalBuildCompactPayload = api.buildCompactPayload.bind(api);
  api.buildCompactPayload = function (state = store.getState?.()) {
    return sanitizeBlockedLocalPayload(originalBuildCompactPayload(state), state);
  };

  if (typeof projectExport.exportProposedJson === "function") {
    const originalExportProposedJson = projectExport.exportProposedJson.bind(projectExport);
    projectExport.exportProposedJson = function () {
      const originalDownloadJson = projectExport.downloadJson;
      if (typeof originalDownloadJson !== "function") return originalExportProposedJson();

      const state = store.getState?.();
      projectExport.downloadJson = function (filename, payload) {
        return originalDownloadJson.call(
          projectExport,
          filename,
          sanitizeBlockedLocalPayload(payload, state)
        );
      };

      try {
        return originalExportProposedJson();
      } finally {
        projectExport.downloadJson = originalDownloadJson;
      }
    };
  }

  const mrLogic = global.MRLogic;
  if (mrLogic && typeof mrLogic.generateMRForPole === "function" && !mrLogic.__compactBlockedLocalMrPatched) {
    const originalGenerateMRForPole = mrLogic.generateMRForPole.bind(mrLogic);
    mrLogic.generateMRForPole = function (poleId) {
      originalGenerateMRForPole(poleId);
      return sanitizeBlockedLocalMr(store.getState?.(), poleId);
    };

    if (typeof mrLogic.generateAllMR === "function") {
      const originalGenerateAllMR = mrLogic.generateAllMR.bind(mrLogic);
      mrLogic.generateAllMR = function () {
        originalGenerateAllMR();
        const state = store.getState?.();
        Object.keys(state?.poles || {}).forEach(poleId => sanitizeBlockedLocalMr(state, poleId));
        return state?.mr || [];
      };
    }

    mrLogic.__compactBlockedLocalMrPatched = true;
  }

  api.isPolePco = isPolePco;
  api.blocksLocalActions = blocksLocalActions;
  api.sanitizeBlockedLocalPayload = sanitizeBlockedLocalPayload;
  api.sanitizeBlockedLocalMr = sanitizeBlockedLocalMr;
  api.sanitizeFullyUgPayload = sanitizeBlockedLocalPayload;
  projectExport.__compactAutoProposedUgGuardPatched = true;
})(window);
