(function (global) {
  "use strict";

  const H = () => global.HeightUtils;
  const S = () => global.AppStore;

  function getHeightInches(value) {
    return H().parseHeight(value);
  }

  function format(inches) {
    return H().formatHeight(inches);
  }

  function getEffectiveCommHOA(sc) {
    return sc ? (sc.existingHOAChange || sc.existingHOA || "") : "";
  }

  function getSpanLengthFeet(span) {
    if (!span) return null;
    const displayInches = H().parseHeight(span.lengthDisplay);
    if (displayInches !== null) return displayInches / 12;
    const raw = Number(span.length);
    if (Number.isFinite(raw)) return raw * 3.280839895;
    return null;
  }

  function getEstimatedSagInches(span) {
    const state = S().getState();
    const lengthFeet = getSpanLengthFeet(span);
    const sagPer100 = H().parseHeight(state.settings.sagPer100Ft || "1'") ?? 12;
    if (!Number.isFinite(lengthFeet)) return 0;
    return Math.round((lengthFeet / 100) * sagPer100);
  }

  function findRemoteComm(spanId, poleId, ownerBase) {
    const span = S().getSpan(spanId);
    if (!span) return null;
    const otherPole = S().getOtherPoleId(span, poleId);
    if (!otherPole) return null;
    return S().getSpanCommsForSpan(spanId).find(row => row.poleId === otherPole && (row.ownerBase || row.owner) === ownerBase) || null;
  }

  function calculatePoleDerived(poleId) {
    const pole = S().getPole(poleId);
    if (!pole) return null;

    const commRows = S().getSpanCommsForPole(poleId);
    const heights = commRows
      .map(sc => H().parseHeight(getEffectiveCommHOA(sc)))
      .filter(value => value !== null);

    const topComm = heights.length ? format(Math.max(...heights)) : "";
    const lowComm = heights.length ? format(Math.min(...heights)) : "";
    const lowPower = H().parseHeight(pole.lowPower);
    const clearance = H().parseHeight(S().getState().settings.clearanceToPower || "4'");
    const maxCommHeight = lowPower !== null && clearance !== null ? format(lowPower - clearance) : "";

    S().updatePoleField(poleId, "topComm", topComm);
    S().updatePoleField(poleId, "lowComm", lowComm);
    S().updatePoleField(poleId, "maxCommHeight", maxCommHeight);

    S().getSpanSidesForPole(poleId).forEach(side => {
      S().upsertSpanSide({
        ...side,
        topComm,
        lowComm,
        maxCommHeight
      });
    });

    return S().getPole(poleId);
  }

  function calculateEndDropForSpanSide(spanId, poleId) {
    const side = S().getSpanSide(spanId, poleId);
    if (!side) return "";
    const proposed = H().parseHeight(side.proposedHOA);
    const midspan = H().parseHeight(side.proposedMidspan);
    if (proposed === null || midspan === null) {
      if (!side.lockedEndDrop) S().upsertSpanSide({ ...side, endDrop: "" });
      return side.endDrop || "";
    }
    const calculated = format(midspan - proposed);
    S().upsertSpanSide({ ...side, endDrop: calculated, lockedEndDrop: false });
    return calculated;
  }

  function calculateMidspanForComm(spanComm) {
    if (!spanComm) return "";
    const span = S().getSpan(spanComm.spanId);
    if (!span) return spanComm.midspan || spanComm.ocalcMS || "";

    const local = H().parseHeight(getEffectiveCommHOA(spanComm));
    const remote = findRemoteComm(spanComm.spanId, spanComm.poleId, spanComm.ownerBase || spanComm.owner);
    const remoteHOA = remote ? getEffectiveCommHOA(remote) : "";
    const remoteInches = H().parseHeight(remoteHOA);

    let calculated = "";
    if (local !== null && remoteInches !== null) {
      const sag = getEstimatedSagInches(span);
      calculated = format(Math.round((local + remoteInches) / 2) - sag);
    } else {
      calculated = spanComm.ocalcMS || spanComm.midspan || "";
    }

    const difference = H().diffLabel(spanComm.existingHOA, spanComm.existingHOAChange || spanComm.existingHOA);
    S().upsertSpanComm({
      ...spanComm,
      remotePoleId: remote ? remote.poleId : S().getOtherPoleId(span, spanComm.poleId),
      remoteHOA,
      calculatedMidspan: calculated,
      difference
    });
    return calculated;
  }

  function updatePoleCommNewHeight(poleId, owner, spanId, newHeight) {
    return updateExistingHOAChange(poleId, owner, spanId, newHeight);
  }

  function updateExistingHOAChange(poleId, owner, spanId, newHeight, wireId = "") {
    const sc = S().getSpanComm(spanId, poleId, owner, wireId) || S().upsertSpanComm({ spanId, poleId, owner, wireId });
    S().upsertSpanComm({ ...sc, existingHOAChange: newHeight || "" });
    recalculateSpan(spanId);
    recalculateSpansForPole(poleId);
    S().getState().selectedPoleId = poleId;
    return S().getSpanComm(spanId, poleId, owner, wireId);
  }

  function updateProposedForSpan(poleId, ownerIgnored, spanId, proposedHeight) {
    return updateSpanSideField(spanId, poleId, "proposedHOA", proposedHeight);
  }

  function updateCambioProposed(poleId, owner, spanId, value) {
    return updateExistingHOAChange(poleId, owner, spanId, value);
  }

  function updateSpanSideField(spanId, poleId, field, value) {
    const allowed = ["proposedHOA", "proposedMidspan", "endDrop", "clearanceReference", "notes"];
    if (!allowed.includes(field)) return null;
    const side = S().getSpanSide(spanId, poleId) || S().upsertSpanSide({ spanId, poleId });
    const data = { ...side, [field]: value || "" };
    if (field === "endDrop") data.lockedEndDrop = Boolean(value);
    S().upsertSpanSide(data);
    if (["proposedHOA", "proposedMidspan"].includes(field)) calculateEndDropForSpanSide(spanId, poleId);
    recalculateSpansForPole(poleId);
    return S().getSpanSide(spanId, poleId);
  }

  function updateSpanCommField(spanId, poleId, owner, wireId, field, value) {
    const allowed = ["existingHOA", "existingHOAChange", "ocalcMS", "midspan", "notes", "mr"];
    if (!allowed.includes(field)) return null;
    const sc = S().getSpanComm(spanId, poleId, owner, wireId) || S().upsertSpanComm({ spanId, poleId, owner, wireId });
    S().upsertSpanComm({ ...sc, [field]: value || "" });
    recalculateSpan(spanId);
    recalculateSpansForPole(poleId);
    return S().getSpanComm(spanId, poleId, owner, wireId);
  }

  function updateOcalcValue(poleId, owner, spanId, field, value) {
    if (field === "proposedHOA" || field === "proposedMidspan" || field === "endDrop") return updateSpanSideField(spanId, poleId, field, value);
    return updateSpanCommField(spanId, poleId, owner, "", field, value);
  }

  function calculateEndDrop(spanId, ownerIgnored, poleId) {
    return calculateEndDropForSpanSide(spanId, poleId);
  }

  function getConnectedSpans(poleId) {
    return S().getConnectedSpans(poleId);
  }

  function getBackspanForPole(poleId) {
    return null;
  }

  function recalculateSpan(spanId) {
    S().getSpanCommsForSpan(spanId).forEach(sc => calculateMidspanForComm(sc));
    S().getSpanSidesForSpan(spanId).forEach(side => calculateEndDropForSpanSide(side.spanId, side.poleId));
  }

  function recalculateSpansForPole(poleId) {
    const spans = S().getConnectedSpans(poleId);
    spans.forEach(span => recalculateSpan(span.spanId));
    calculatePoleDerived(poleId);
    spans.forEach(span => {
      const other = S().getOtherPoleId(span, poleId);
      if (other) calculatePoleDerived(other);
    });
    global.MRLogic.generateMRForPole(poleId);
    global.Validations.validatePole(poleId);
  }

  function calculateOcalcValues() {
    const state = S().getState();
    Object.values(state.spanSides).forEach(side => calculateEndDropForSpanSide(side.spanId, side.poleId));
    Object.values(state.spanComms).forEach(sc => calculateMidspanForComm(sc));
  }

  function recalculateAll() {
    const state = S().getState();
    Object.keys(state.spans).forEach(recalculateSpan);
    Object.keys(state.poles).forEach(calculatePoleDerived);
    global.MRLogic.generateAllMR();
    global.Validations.validateAll();
  }

  global.Calculations = {
    updatePoleCommNewHeight,
    updateExistingHOAChange,
    updateProposedForSpan,
    updateCambioProposed,
    updateSpanSideField,
    updateSpanCommField,
    updateOcalcValue,
    calculateEndDrop,
    calculateEndDropForSpanSide,
    calculateMidspanForComm,
    getConnectedSpans,
    getBackspanForPole,
    recalculateSpan,
    recalculateSpansForPole,
    calculateOcalcValues,
    recalculateAll,
    getEffectiveCommHOA,
    getEstimatedSagInches,
    findRemoteComm
  };
})(window);
