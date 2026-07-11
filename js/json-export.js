(function (global) {
  "use strict";

  // JSON export is the lightweight save-point format. It stores the full state,
  // including user edits and calculated fields, so work can be resumed exactly.
  const S = () => global.AppStore;
  const H = () => global.HeightUtils;

  function downloadJson(filename, data) {
    const blob = new Blob([JSON.stringify(data, null, 2)], { type: "application/json;charset=utf-8" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    a.remove();
    URL.revokeObjectURL(url);
  }

  function safeJobFilePart(value) {
    const raw = String(value || "pole_job")
      .replace(/\.[^.]+$/, "")
      .replace(/^excel[_\-\s]*/i, "")
      .replace(/[_\-\s]*20\d{2}[-_]\d{2}[-_]\d{2}$/i, "")
      .replace(/_Pole_Calculator$/i, "")
      .replace(/_AutoProposed$/i, "")
      .trim();
    return (raw || "pole_job")
      .replace(/[<>:"/\\|?*\u0000-\u001F]+/g, "_")
      .replace(/\s+/g, "_");
  }

  function exportJson() {
    global.Calculations.recalculateAll();
    const state = S().getState();
    const date = new Date().toISOString().slice(0, 10);
    downloadJson(`pole-calculator-data-${date}.json`, {
      app: "pole-calculator",
      exportedAt: new Date().toISOString(),
      version: state.version || S().CURRENT_VERSION,
      state
    });
  }

  function hasSpanLength(span) {
    if (!span) return "";
    const displayInches = H().parseHeight(span.lengthDisplay || "");
    if (displayInches !== null) return true;
    const raw = Number(span.length);
    return Number.isFinite(raw) && raw > 0;
  }

  function oppositeDirection(direction) {
    return {
      N: "S",
      NE: "SW",
      E: "W",
      SE: "NW",
      S: "N",
      SW: "NE",
      W: "E",
      NW: "SE"
    }[direction] || direction || "";
  }

  function directionFromPole(span, poleId) {
    if (!span) return "";
    return span.fromPole === poleId ? (span.direction || "") : oppositeDirection(span.direction || "");
  }

  function makeReadyRefsForProposal(state, poleId, span) {
    const direction = directionFromPole(span, poleId).toUpperCase();
    const refs = (state.makeReadyReferences || []).filter(ref => ref.poleId === poleId);
    const directionMatches = refs.filter(ref => {
      const tokens = Array.isArray(ref.attachmentDirectionTokens) ? ref.attachmentDirectionTokens.map(String) : [];
      if (!direction || !tokens.length) return true;
      return tokens.map(token => token.toUpperCase()).includes(direction);
    });
    return directionMatches.length ? directionMatches : refs;
  }

  function spanExportInfo(span, poleId) {
    const fromPole = span?.fromPole || "";
    const toPole = span?.toPole || "";
    return {
      label: span ? `${fromPole} -> ${toPole}` : "",
      toPole,
      type: span?.type || "",
      direction: directionFromPole(span, poleId),
      bearingDegrees: span?.bearingDegrees ?? "",
      lengthDisplay: span?.lengthDisplay || ""
    };
  }

  function hasSpanExportGeometry(span, poleId) {
    if (!span) return false;
    const direction = directionFromPole(span, poleId);
    const bearing = Number(span.bearingDegrees);
    return Boolean(
      direction ||
      Number.isFinite(bearing) ||
      span.lengthDisplay ||
      hasSpanLength(span)
    );
  }

  function shouldExportSpanForPole(span, poleId) {
    // AutoProposed is organized by the pole where work is being proposed. For
    // that reason P29 only exports P29 -> ... spans; incoming P30 -> P29 spans
    // belong to P30 and would duplicate/confuse the O-Calc import payload.
    return Boolean(span && span.fromPole === poleId && hasSpanExportGeometry(span, poleId));
  }

  function attachmentExportInfo(ref) {
    if (!ref) return null;
    return {
      attachmentSize: ref.attachmentSizeRaw || "",
      messenger: ref.attachmentMessenger || "",
      fiber: ref.attachmentFiber || "",
      direction: ref.attachmentDirection || "",
      attachmentType: ref.attachmentType || ""
    };
  }

  function ensurePoleExport(map, poleId, state) {
    const key = poleId || "UNKNOWN_POLE";
    if (!map.has(key)) {
      map.set(key, {
        poleId: key,
        proposedOwner: state.settings?.proposedOwner || "",
        proposed: [],
        attachments: [],
        spans: [],
        commMakeReady: []
      });
    }
    return map.get(key);
  }

  function ensureSpanExport(poleItem, span, poleId) {
    if (!shouldExportSpanForPole(span, poleId)) return null;
    const info = spanExportInfo(span, poleId);
    const key = info.label || `${poleId || ""}->${info.toPole || ""}`;
    let item = poleItem.spans.find(row => row.label === key);
    if (!item) {
      item = { ...info };
      poleItem.spans.push(item);
    }
    return item;
  }

  function addOutgoingSpansForPole(poleItem, poleId) {
    S().getConnectedSpans(poleId)
      .filter(span => shouldExportSpanForPole(span, poleId))
      .forEach(span => ensureSpanExport(poleItem, span, poleId));
  }

  function addAttachmentForPole(poleItem, attachment) {
    if (!attachment) return;
    const key = JSON.stringify(attachment);
    const exists = poleItem.attachments.some(row => JSON.stringify(row) === key);
    if (!exists) poleItem.attachments.push(attachment);
  }

  function addProposedForPole(poleItem, proposed) {
    if (!proposed) return;
    const key = proposed.spanLabel || JSON.stringify(proposed);
    const existing = poleItem.proposed.find(row => (row.spanLabel || JSON.stringify(row)) === key);
    if (existing) {
      Object.assign(existing, proposed);
    } else {
      poleItem.proposed.push(proposed);
    }
  }

  function exportProposedJson() {
    global.Calculations.recalculateAll();
    const state = S().getState();
    const date = new Date().toISOString().slice(0, 10);
    const polesById = new Map();

    Object.values(state.spanSides || {})
      .filter(side => side.proposedHOA)
      .forEach(side => {
        const span = S().getSpan(side.spanId);
        const poleItem = ensurePoleExport(polesById, side.poleId, state);
        addOutgoingSpansForPole(poleItem, side.poleId);
        const spanItem = ensureSpanExport(poleItem, span, side.poleId);
        if (!spanItem) return;
        const primaryAttachmentReference = makeReadyRefsForProposal(state, side.poleId, span)[0] || null;
        const attachment = attachmentExportInfo(primaryAttachmentReference);
        addAttachmentForPole(poleItem, attachment);
        addProposedForPole(poleItem, {
          spanLabel: spanItem.label,
          fromPole: span?.fromPole || side.poleId || "",
          toPole: spanItem.toPole,
          proposed: side.proposedHOA || "",
          endDrop: side.endDrop || "",
          nextPoleProposed: side.proposedHOAChange || ""
        });
      });

    Object.values(state.spanComms || {}).forEach(sc => {
      const mrLine = global.MRLogic?.generateMRForComm(sc) || "";
      if (!sc.existingHOAChange || !mrLine) return;
      const poleItem = ensurePoleExport(polesById, sc.poleId, state);
      addOutgoingSpansForPole(poleItem, sc.poleId);
      poleItem.commMakeReady.push(mrLine);
      poleItem.commMakeReady = Array.from(new Set(poleItem.commMakeReady));
    });

    const poles = Array.from(polesById.values())
      .map(pole => ({
        ...pole,
        proposed: pole.proposed
          .sort((a, b) => (a.spanLabel || "").localeCompare(b.spanLabel || "", undefined, { numeric: true })),
        attachments: pole.attachments
          .sort((a, b) => (a.spanLabel || "").localeCompare(b.spanLabel || "", undefined, { numeric: true })),
        spans: pole.spans
          .sort((a, b) => a.label.localeCompare(b.label, undefined, { numeric: true }))
      }))
      .filter(pole => pole.proposed.length || pole.attachments.length || pole.spans.length || pole.commMakeReady.length)
      .sort((a, b) => a.poleId.localeCompare(b.poleId, undefined, { numeric: true }));

    const jobName = safeJobFilePart(state.importedFileName || `pole_job_${date}`);
    downloadJson(`${jobName}_AutoProposed.json`, {
      app: "pole-calculator",
      exportType: "proposed-for-ocalc",
      exportedAt: new Date().toISOString(),
      version: state.version || S().CURRENT_VERSION,
      sourceFile: state.importedFileName || "",
      settings: {
        proposedOwner: state.settings?.proposedOwner || "",
        position: state.settings?.position || "",
        poleCommCommClearance: state.settings?.commClearance || "",
        poleBoltBoltClearance: state.settings?.boltClearance || "",
        midspanPowerCommClearance: state.settings?.midspanPowerCommClearance || "",
        midspanCommCommClearance: state.settings?.midspanCommCommClearance || ""
      },
      poles
    });
  }

  global.ProjectExport = {
    exportJson,
    exportProposedJson,
    downloadJson
  };
})(window);
