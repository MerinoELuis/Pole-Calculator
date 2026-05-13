(function (global) {
  "use strict";

  const S = () => global.AppStore;

  function calculateMidspanForSpan(spanId, owner) {
    const rows = S().getSpanCommsForSpan(spanId).filter(sc => !owner || sc.owner === owner || sc.ownerBase === owner);
    rows.forEach(sc => global.Calculations.calculateMidspanForComm(sc));
    return S().getSpanCommsForSpan(spanId).filter(sc => !owner || sc.owner === owner || sc.ownerBase === owner);
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

  global.MidspanLogic = {
    calculateMidspanForSpan,
    calculateAllMidspans,
    getMidspanRowsForPole,
    estimateMidspanFromEndDrop
  };
})(window);
