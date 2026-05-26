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

  function heightToDecimalFeet(value) {
    const inches = H().parseHeight(value || "");
    return inches === null ? "" : Number((inches / 12).toFixed(3));
  }

  function spanLengthFeet(span) {
    if (!span) return "";
    const displayInches = H().parseHeight(span.lengthDisplay || "");
    if (displayInches !== null) return Number((displayInches / 12).toFixed(3));
    const raw = Number(span.length);
    return Number.isFinite(raw) ? Number((raw * 3.280839895).toFixed(3)) : "";
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

  function exportProposedJson() {
    global.Calculations.recalculateAll();
    const state = S().getState();
    const date = new Date().toISOString().slice(0, 10);

    const proposals = Object.values(state.spanSides || {})
      .filter(side => side.proposedHOA)
      .map(side => {
        const span = S().getSpan(side.spanId);
        const otherPoleId = span ? S().getOtherPoleId(span, side.poleId) : "";
        const attachmentReferences = makeReadyRefsForProposal(state, side.poleId, span);
        const primaryAttachmentReference = attachmentReferences[0] || null;
        return {
          poleId: side.poleId,
          spanId: side.spanId,
          spanLabel: span ? `${span.fromPole} -> ${span.toPole}` : side.spanId,
          otherPoleId,
          spanType: span?.type || "",
          spanDirection: directionFromPole(span, side.poleId),
          spanBearingDegrees: span?.bearingDegrees ?? "",
          spanLength: span?.length || "",
          spanLengthDisplay: span?.lengthDisplay || "",
          spanLengthFeet: spanLengthFeet(span),
          proposed: side.proposedHOA || "",
          proposedFeet: heightToDecimalFeet(side.proposedHOA),
          endDrop: side.endDrop || "",
          nextPoleProposed: side.proposedHOAChange || "",
          ocalcMS: side.ocalcMS || "",
          msProposed: side.msProposed || "",
          adjustedFinalMS: side.finalMidspan || "",
          maxHeightAtMS: span?.midspanMaxCommHeight || "",
          lowPowerAtMS: span?.midspanLowPower || "",
          proposedOwner: state.settings?.proposedOwner || "",
          attachmentSize: primaryAttachmentReference?.attachmentSizeRaw || "",
          messenger: primaryAttachmentReference?.attachmentMessenger || "",
          fiber: primaryAttachmentReference?.attachmentFiber || "",
          attachmentDirection: primaryAttachmentReference?.attachmentDirection || "",
          environment: span?.environment || "",
          environmentClearance: span?.environmentClearance || "",
          notes: side.notes || "",
          primaryAttachmentReference,
          attachmentReferences
        };
      });

    const commMovements = Object.values(state.spanComms || {})
      .map(sc => {
        const mrLine = global.MRLogic?.generateMRForComm(sc) || "";
        if (!sc.existingHOAChange || !mrLine) return null;
        const span = S().getSpan(sc.spanId);
        return {
          poleId: sc.poleId,
          spanId: sc.spanId,
          spanLabel: span ? `${span.fromPole} -> ${span.toPole}` : sc.spanId,
          otherPoleId: span ? S().getOtherPoleId(span, sc.poleId) : "",
          owner: global.Calculations.commOwnerLabel(sc),
          rawOwner: sc.rawOwner || "",
          ownerBase: sc.ownerBase || sc.owner || "",
          serviceDrop: Boolean(sc.serviceDrop),
          existingHOA: sc.existingHOA || "",
          hoaChange: sc.existingHOAChange || "",
          importedMidspan: sc.midspan || sc.ocalcMS || "",
          calculatedMidspan: sc.calculatedMidspan || "",
          mrLine,
          wireId: sc.wireId || "",
          size: sc.size || "",
          construction: sc.construction || "",
          insulator: sc.insulator || ""
        };
      })
      .filter(Boolean);

    const commMakeReadyByPole = commMovements.reduce((acc, item) => {
      if (!acc[item.poleId]) acc[item.poleId] = [];
      acc[item.poleId].push(item.mrLine);
      acc[item.poleId] = Array.from(new Set(acc[item.poleId]));
      return acc;
    }, {});

    const spans = Object.values(state.spans || {})
      .filter(span => proposals.some(proposal => proposal.spanId === span.spanId) || commMovements.some(move => move.spanId === span.spanId))
      .map(span => ({
        spanId: span.spanId,
        fromPole: span.fromPole,
        toPole: span.toPole,
        type: span.type || "",
        rawType: span.rawType || "",
        direction: span.direction || "",
        bearingDegrees: span.bearingDegrees ?? "",
        length: span.length || "",
        lengthDisplay: span.lengthDisplay || "",
        lengthFeet: spanLengthFeet(span),
        environment: span.environment || "",
        environmentClearance: span.environmentClearance || "",
        lowPowerAtMS: span.midspanLowPower || "",
        maxHeightAtMS: span.midspanMaxCommHeight || ""
      }));

    downloadJson(`pole-proposed-ocalc-${date}.json`, {
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
      spans,
      proposals,
      commMovements,
      commMakeReadyByPole,
      makeReadyReferences: state.makeReadyReferences || []
    });
  }

  global.ProjectExport = {
    exportJson,
    exportProposedJson,
    downloadJson
  };
})(window);
