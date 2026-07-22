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
      // Strip an existing calculator/export suffix first so a preceding date
      // becomes the end of the job name and can be removed consistently.
      .replace(/_(?:Pole_Calculator|AutoProposed|Debug)$/i, "")
      .replace(/^excel[_\-\s]*/i, "")
      .replace(/[_\-\s]*20\d{2}[-_]\d{2}[-_]\d{2}$/i, "")
      .trim();
    return (raw || "pole_job")
      .replace(/[<>:"/\\|?*\u0000-\u001F]+/g, "_")
      .replace(/\s+/g, "_");
  }

  function normalizedFiberName(value) {
    const match = String(value || "").match(/\b(\d+)\s*CT\b/i);
    return match ? `${match[1]}CT Fiber` : "";
  }

  function attachmentSizeSettings(state) {
    const configured = state.settings?.fiberSizes && typeof state.settings.fiberSizes === "object"
      ? state.settings.fiberSizes
      : {};
    const names = new Set(Object.keys(configured));
    (state.makeReadyReferences || []).forEach(ref => {
      const name = normalizedFiberName(`${ref.attachmentFiber || ""} ${ref.attachmentSizeRaw || ""}`);
      if (name) names.add(name);
    });
    return {
      messengerSize: state.settings?.attachmentMessengerSize || "",
      fibers: Array.from(names)
        .sort((a, b) => Number((a.match(/\d+/) || [0])[0]) - Number((b.match(/\d+/) || [0])[0]))
        .map(fiber => ({ fiber, size: configured[fiber] || "" }))
    };
  }

  function normalizedOwner(value) {
    return String(value || "")
      .replace(/^communication\s*>\s*/i, "")
      .replace(/century\s*link/g, "centurylink")
      .trim()
      .toLowerCase();
  }

  function multipleWireOwnerGroups(state) {
    const groups = new Map();
    Object.values(state.spanComms || {}).forEach(row => {
      const key = [row.spanId, row.poleId, normalizedOwner(row.ownerBase || row.owner)].join(" | ");
      if (!groups.has(key)) groups.set(key, []);
      groups.get(key).push({ owner: row.owner, wireId: row.wireId || "", existingHOA: row.existingHOA || "", midspan: row.midspan || "" });
    });
    return Array.from(groups.entries())
      .filter(([, rows]) => rows.length > 1)
      .map(([identity, rows]) => ({ identity, count: rows.length, rows }));
  }

  // This mirrors the movement formula used by calculations.js and records each
  // input next to its matched remote comm. The debug file therefore shows when
  // a duplicated/stale row or an unexpected remote match changes the result.
  function midspanDebugRow(row) {
    const span = S().getSpan(row.spanId);
    const remote = global.Calculations.findRemoteComm(
      row.spanId,
      row.poleId,
      row.ownerBase || row.owner,
      row.wireId || ""
    );
    const localExisting = H().parseHeight(row.existingHOA || "");
    const localEffectiveText = global.Calculations.getEffectiveCommHOA(row);
    const localEffective = H().parseHeight(localEffectiveText);
    const remoteExisting = H().parseHeight(remote?.existingHOA || "");
    const remoteEffectiveText = remote ? global.Calculations.getEffectiveCommHOA(remote) : "";
    const remoteEffective = H().parseHeight(remoteEffectiveText);
    const importedText = row.midspan || row.ocalcMS || "";
    const imported = H().parseHeight(importedText);
    const localAdjustment = localExisting !== null && localEffective !== null ? (localEffective - localExisting) / 2 : 0;
    const remoteAdjustment = remoteExisting !== null && remoteEffective !== null ? (remoteEffective - remoteExisting) / 2 : 0;
    const expected = imported !== null ? Math.round(imported + localAdjustment + remoteAdjustment) : null;

    return {
      identity: {
        spanId: row.spanId,
        poleId: row.poleId,
        owner: row.owner,
        ownerBase: row.ownerBase || "",
        wireId: row.wireId || "",
        wireIndex: row.wireIndex || ""
      },
      span: span ? { fromPole: span.fromPole, toPole: span.toPole, type: span.type || "" } : null,
      local: {
        existingHOA: row.existingHOA || "",
        hoaChange: row.existingHOAChange || "",
        effectiveHOA: localEffectiveText || "",
        halfMovement: H().formatHeight(localAdjustment),
        halfMovementInches: localAdjustment
      },
      remoteMatch: remote ? {
        spanId: remote.spanId,
        poleId: remote.poleId,
        owner: remote.owner,
        wireId: remote.wireId || "",
        existingHOA: remote.existingHOA || "",
        hoaChange: remote.existingHOAChange || "",
        effectiveHOA: remoteEffectiveText || "",
        halfMovement: H().formatHeight(remoteAdjustment),
        halfMovementInches: remoteAdjustment
      } : null,
      midspan: {
        imported: importedText,
        expectedFromFormula: expected === null ? "" : H().formatHeight(expected),
        storedCalculated: row.calculatedMidspan || "",
        displayed: global.Calculations.displayMidspanForComm(row) || "",
        formula: "imported + (local change / 2) + (remote change / 2)"
      },
      flagging: {
        status: row.flaggingStatus || "",
        message: row.flaggingMessage || ""
      }
    };
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

  /** Recalculates and downloads the compact downstream O-Calc proposal package. */
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

    // A terminal pole can carry the attachment height used by the preceding
    // pole's End Drop without owning an outgoing span or midspan.
    Object.values(state.poles || {})
      .filter(pole => pole.standaloneProposedHOA)
      .forEach(pole => {
        const poleItem = ensurePoleExport(polesById, pole.poleId, state);
        addOutgoingSpansForPole(poleItem, pole.poleId);
        addProposedForPole(poleItem, {
          terminal: true,
          proposed: pole.standaloneProposedHOA
        });
      });

    Object.values(state.spanComms || {}).forEach(sc => {
      const mrLine = global.MRLogic?.generateMRForComm(sc) || "";
      if (!mrLine) return;
      const poleItem = ensurePoleExport(polesById, sc.poleId, state);
      addOutgoingSpansForPole(poleItem, sc.poleId);
      poleItem.commMakeReady.push(...String(mrLine).split(/\n+/).filter(Boolean));
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

    const jobName = safeJobFilePart(state.jobName || state.importedFileName || `pole_job_${date}`);
    const payload = {
      app: "pole-calculator",
      exportType: "proposed-for-ocalc",
      exportedAt: new Date().toISOString(),
      version: state.version || S().CURRENT_VERSION,
      sourceFile: state.importedFileName || "",
      attachmentSizes: attachmentSizeSettings(state),
      settings: {
        proposedOwner: state.settings?.proposedOwner || "",
        position: state.settings?.position || "",
        poleCommCommClearance: state.settings?.commClearance || "",
        poleBoltBoltClearance: state.settings?.boltClearance || "",
        midspanPowerCommClearance: state.settings?.midspanPowerCommClearance || "",
        midspanCommCommClearance: state.settings?.midspanCommCommClearance || ""
      },
      poles
    };
    downloadJson(`${jobName}_AutoProposed.json`, payload);
  }

  /** Recalculates and downloads full state plus midspan diagnostic traces. */
  function exportDebugJson() {
    global.Calculations.recalculateAll();
    const state = S().getState();
    const jobName = safeJobFilePart(state.jobName || state.importedFileName || "pole_job");
    const spanComms = Object.values(state.spanComms || {});
    downloadJson(`${jobName}_Debug.json`, {
      app: "pole-calculator",
      exportType: "calculation-debug",
      exportedAt: new Date().toISOString(),
      sourceFile: state.importedFileName || "",
      summary: {
        poles: Object.keys(state.poles || {}).length,
        spans: Object.keys(state.spans || {}).length,
        spanComms: spanComms.length,
        spanPower: Object.keys(state.spanPower || {}).length,
        multipleWiresPerOwner: multipleWireOwnerGroups(state),
        lastExcelUpdate: state.updateDiagnostics || null
      },
      midspanCalculations: spanComms.map(midspanDebugRow),
      state
    });
  }

  /** @namespace ProjectExport */
  global.ProjectExport = {
    exportJson,
    exportProposedJson,
    exportDebugJson,
    downloadJson
  };
})(window);
