(function (global) {
  "use strict";

  function calculateCommMidspanDetails(spanComm, dependencies) {
    const {
      store,
      heightUtils,
      format,
      getEffectiveCommHOA,
      findRemoteComm,
      ownerCanonicalOf,
      findMidspanSourceComm,
      getImportedMidspanInchesForComm
    } = dependencies;
    if (!spanComm) return { calculated: "", remote: null, remoteHOA: "", localAdjustment: 0, remoteAdjustment: 0 };
    const span = store.getSpan(spanComm.spanId);
    if (!span) return { calculated: spanComm.midspan || "", remote: null, remoteHOA: "", localAdjustment: 0, remoteAdjustment: 0 };

    const localExisting = heightUtils.parseHeight(spanComm.existingHOA);
    const localCurrent = heightUtils.parseHeight(getEffectiveCommHOA(spanComm));
    const remote = findRemoteComm(spanComm.spanId, spanComm.poleId, ownerCanonicalOf(spanComm), spanComm.wireId || "");
    const remoteHOA = remote ? getEffectiveCommHOA(remote) : "";
    const remoteExisting = heightUtils.parseHeight(remote?.existingHOA || "");
    const remoteCurrent = heightUtils.parseHeight(remoteHOA);
    const midspanSource = findMidspanSourceComm(spanComm) || spanComm;
    const importedMidspan = getImportedMidspanInchesForComm(midspanSource);

    let calculated = "";
    let localAdjustment = 0;
    let remoteAdjustment = 0;
    if (importedMidspan !== null) {
      // Each endpoint contributes half of its HOA change to the imported
      // Midspan baseline; the caller owns identity matching and state writes.
      localAdjustment = localExisting !== null && localCurrent !== null ? (localCurrent - localExisting) / 2 : 0;
      remoteAdjustment = remoteExisting !== null && remoteCurrent !== null ? (remoteCurrent - remoteExisting) / 2 : 0;
      calculated = format(Math.round(importedMidspan + localAdjustment + remoteAdjustment));
    }

    return { calculated, remote, remoteHOA, localAdjustment, remoteAdjustment };
  }

  const S = () => global.AppStore;

  function calculateMidspanForSpan(spanId, owner) {
    const rows = S().getSpanCommsForSpan(spanId).filter(sc => !owner || S().getOwnerDisplay?.(sc) === owner || S().getOwnerCanonical?.(sc) === owner || sc.owner === owner);
    rows.forEach(sc => global.Calculations.calculateMidspanForComm(sc));
    return S().getSpanCommsForSpan(spanId).filter(sc => !owner || S().getOwnerDisplay?.(sc) === owner || S().getOwnerCanonical?.(sc) === owner || sc.owner === owner);
  }

  function calculateAllMidspans() {
    const state = S().getState();
    Object.values(state.spans).forEach(span => calculateMidspanForSpan(span.spanId));
  }

  function getMidspanRowsForPole(poleId) {
    return S().getSpanCommsForPole(poleId).map(sc => {
      const span = S().getSpan(sc.spanId);
      return {
        ...sc,
        spanLabel: span ? `${span.fromPole} → ${span.toPole}` : sc.spanId,
        otherPoleId: span ? S().getOtherPoleId(span, poleId) : "",
        midspanDisplay: sc.calculatedMidspan || sc.midspan || sc.ocalcMS || ""
      };
    });
  }

  function estimateMidspanFromEndDrop(spanSide) {
    const proposed = global.HeightUtils.parseHeight(spanSide.proposedHOA);
    const endDrop = global.HeightUtils.parseHeight(spanSide.endDrop);
    if (proposed === null || endDrop === null) return "";
    return global.HeightUtils.formatHeight(proposed + endDrop);
  }

  const api = {
    calculateCommMidspanDetails,
    calculateMidspanForSpan,
    calculateAllMidspans,
    getMidspanRowsForPole,
    estimateMidspanFromEndDrop
  };
  global.MidspanEngine = Object.freeze(api);
  // Keep the old global name for modules and saved browser integrations that
  // still use the span-oriented facade.
  global.MidspanLogic = global.MidspanEngine;
})(window);
