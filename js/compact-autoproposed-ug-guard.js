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

  function isPoleFullyUg(state, poleId) {
    if (typeof api.isPoleFullyUg === "function") return api.isPoleFullyUg(state, poleId);
    const pole = state?.poles?.[poleId];
    return Boolean(pole?.ugActive || hasUgToken(poleId));
  }

  function sanitizeFullyUgPole(pole, state) {
    if (!pole || !isPoleFullyUg(state, pole.id)) return pole;

    const sanitized = { id: pole.id };
    if (Array.isArray(pole.spans) && pole.spans.length) {
      sanitized.spans = pole.spans.map(span => {
        const item = {
          to: span.to,
          kind: span.kind
        };
        if (Object.prototype.hasOwnProperty.call(span, "bearing")) item.bearing = span.bearing;
        if (Object.prototype.hasOwnProperty.call(span, "length")) item.length = span.length;
        item.ug = true;
        return item;
      });
    }
    return sanitized;
  }

  function sanitizeFullyUgPayload(payload, state = store.getState?.()) {
    if (!payload || !Array.isArray(payload.poles)) return payload;
    return {
      ...payload,
      poles: payload.poles.map(pole => sanitizeFullyUgPole(pole, state))
    };
  }

  const originalBuildCompactPayload = api.buildCompactPayload.bind(api);
  api.buildCompactPayload = function (state = store.getState?.()) {
    return sanitizeFullyUgPayload(originalBuildCompactPayload(state), state);
  };

  if (typeof projectExport.exportProposedJson === "function") {
    const originalExportProposedJson = projectExport.exportProposedJson.bind(projectExport);
    projectExport.exportProposedJson = function () {
      const originalDownloadJson = projectExport.downloadJson;
      if (typeof originalDownloadJson !== "function") return originalExportProposedJson();

      const state = store.getState?.();
      projectExport.downloadJson = function (filename, payload) {
        return originalDownloadJson.call(projectExport, filename, sanitizeFullyUgPayload(payload, state));
      };

      try {
        return originalExportProposedJson();
      } finally {
        projectExport.downloadJson = originalDownloadJson;
      }
    };
  }

  api.sanitizeFullyUgPayload = sanitizeFullyUgPayload;
  projectExport.__compactAutoProposedUgGuardPatched = true;
})(window);
