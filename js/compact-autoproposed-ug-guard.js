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

  api.isPolePco = isPolePco;
  api.sanitizeBlockedLocalPayload = sanitizeBlockedLocalPayload;
  api.sanitizeFullyUgPayload = sanitizeBlockedLocalPayload;
  projectExport.__compactAutoProposedUgGuardPatched = true;
})(window);
