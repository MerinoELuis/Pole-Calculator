(function (global) {
  "use strict";

  // Validations stores broad project warnings. Table-level flagging lives in
  // calculations.js; these warnings remain useful for import/data integrity.
  const H = () => global.HeightUtils;
  const S = () => global.AppStore;

  function addWarning(list, poleId, spanId, owner, code, message, level = "warning") {
    list.push({ poleId, spanId, owner, code, message, level });
  }

  function exceedsMax(height, maxHeight) {
    const h = H().parseHeight(height);
    const max = H().parseHeight(maxHeight);
    if (h === null || max === null) return false;
    return h > max;
  }

  function validateSpan(spanId) {
    const state = S().getState();
    const span = S().getSpan(spanId);
    const warnings = [];
    if (!span) return warnings;

    if (!span.toPole || /^Unknown-/i.test(span.toPole)) {
      addWarning(warnings, span.fromPole, span.spanId, "", "EDITABLE_OTHER_POLE", "The other pole was generated because Span did not link it. You can edit its data.");
    }

    const spanRows = S().getSpanCommsForSpan(spanId);
    if (!spanRows.length) addWarning(warnings, span.fromPole, span.spanId, "", "EMPTY_SPAN", "The span has no comms imported from Span.Wire.");
    if (spanRows.length && !span.midspanLowPower) {
      addWarning(warnings, span.fromPole, span.spanId, "", "MISSING_MIDSPAN_POWER", "Missing power midspan to calculate max height at midspan.");
    }
    if (span.environment && span.environment !== "NONE") {
      if (!span.environmentClearance) addWarning(warnings, span.fromPole, span.spanId, "", "MISSING_ENV_CLEARANCE", "El span tiene environment pero falta clearance editable.");
      if (span.environmentClearance && span.environmentClearance !== "Variable" && !H().isValidHeight(span.environmentClearance)) {
        addWarning(warnings, span.fromPole, span.spanId, "", "INVALID_ENV_CLEARANCE", "Invalid environment clearance.", "danger");
      }
    }

    S().getSpanSidesForSpan(spanId).forEach(side => {
      if (side.proposedHOA && !H().isValidHeight(side.proposedHOA)) addWarning(warnings, side.poleId, spanId, "", "INVALID_PROPOSED", "Invalid Proposed.", "danger");
      if (side.proposedHOAChange && !H().isValidHeight(side.proposedHOAChange)) addWarning(warnings, side.poleId, spanId, "", "INVALID_NEXT_POLE_PROPOSED", "Invalid Next Pole Proposed.", "danger");
      if (side.ocalcMS && !H().isValidHeight(side.ocalcMS)) addWarning(warnings, side.poleId, spanId, "", "INVALID_OCALC_MS", "Invalid O-CALC MS.", "danger");
      if (side.msProposed && !H().isValidHeight(side.msProposed)) addWarning(warnings, side.poleId, spanId, "", "INVALID_MS_PROPOSED", "Invalid MS Proposed.", "danger");
      if (side.finalMidspan && !H().isValidHeight(side.finalMidspan)) addWarning(warnings, side.poleId, spanId, "", "INVALID_FINAL_MIDSPAN", "Invalid Adjusted Final MS.", "danger");
      if (side.endDrop && !H().isValidHeight(side.endDrop)) addWarning(warnings, side.poleId, spanId, "", "INVALID_END_DROP", "Invalid End Drop.", "danger");
      if (side.proposedHOA && exceedsMax(side.proposedHOA, side.maxCommHeight)) addWarning(warnings, side.poleId, spanId, "", "PROPOSED_ABOVE_MAX", `Proposed ${side.proposedHOA} exceeds max height ${side.maxCommHeight}.`, "danger");
      const boltClearance = global.Calculations.evaluateProposedPoleClearance(side);
      if (!boltClearance.ok) addWarning(warnings, side.poleId, spanId, "", "PROPOSED_BOLT_CLEARANCE", boltClearance.message, "danger");
      if (side.clearanceMSIssue) {
        addWarning(warnings, side.poleId, spanId, "", "MS_PROPOSED_CLEARANCE", side.clearanceMSMessage || "MS Proposed has a clearance issue.", side.clearanceMSStatus === "PROBLEM" ? "danger" : "warning");
      }
      if (side.proposedFlaggingStatus === "PROBLEM") {
        addWarning(warnings, side.poleId, spanId, "", "PROPOSED_POSITION", side.proposedFlaggingMessage || "Proposed does not respect the configured position.", "danger");
      }
    });

    spanRows.forEach(sc => {
      if (sc.calculatedMidspan && span.midspanMaxCommHeight && exceedsMax(sc.calculatedMidspan, span.midspanMaxCommHeight)) {
        addWarning(warnings, sc.poleId, spanId, sc.owner, "MIDSPAN_ABOVE_POWER_CLEARANCE", `Midspan ${sc.calculatedMidspan} exceeds max height at midspan ${span.midspanMaxCommHeight}.`, "danger");
      }
      if (sc.clearanceMSIssue) {
        addWarning(warnings, sc.poleId, spanId, sc.owner, "COMM_MIDSPAN_CLEARANCE", sc.clearanceMSMessage || "The comm does not have enough midspan clearance.", "danger");
      }
      const effectiveHOA = sc.existingHOAChange || sc.existingHOA;
      if (effectiveHOA && exceedsMax(effectiveHOA, S().getPole(sc.poleId)?.maxCommHeight)) {
        const label = sc.existingHOAChange ? "HOA Change" : "Existing HOA";
        addWarning(warnings, sc.poleId, spanId, sc.owner, "COMM_ABOVE_MAX", `${label} ${effectiveHOA} exceeds max height on pole.`, "danger");
      }
      if (sc.existingHOAChange && !global.MRLogic.generateMRForComm(sc)) {
        addWarning(warnings, sc.poleId, spanId, sc.owner, "CHANGE_WITHOUT_MR", "HOA Change exists but no MR was generated.");
      }
    });

    state.warnings = state.warnings.filter(w => w.spanId !== spanId).concat(warnings);
    return warnings;
  }

  function validatePole(poleId) {
    const state = S().getState();
    const pole = S().getPole(poleId);
    const warnings = [];
    if (!pole) return warnings;

    if (pole.lowPower && !H().isValidHeight(pole.lowPower)) addWarning(warnings, poleId, "", "", "INVALID_LOW_POWER", "Invalid Low Power.", "danger");
    if (!pole.lowPower && !pole.isGenerated) addWarning(warnings, poleId, "", "", "MISSING_LOW_POWER", "Missing Low Power on pole.");
    if (pole.standaloneProposedHOA && !H().isValidHeight(pole.standaloneProposedHOA)) {
      addWarning(warnings, poleId, "", "", "INVALID_STANDALONE_PROPOSED", "Invalid terminal pole Proposed.", "danger");
    } else if (pole.standaloneProposedHOA && exceedsMax(pole.standaloneProposedHOA, pole.maxCommHeight)) {
      addWarning(warnings, poleId, "", "", "STANDALONE_PROPOSED_ABOVE_MAX", `Proposed ${pole.standaloneProposedHOA} exceeds max height ${pole.maxCommHeight}.`, "danger");
    }

    pole.comms.forEach(comm => {
      if (!comm.owner) addWarning(warnings, poleId, "", "", "UNKNOWN_OWNER", "Hay un comm sin owner.");
      if (comm.unknownOwner) addWarning(warnings, poleId, "", comm.owner, "UNKNOWN_OWNER", `Unnormalized owner: ${comm.rawOwner || comm.owner}.`);
      if (comm.existingHOA && !H().isValidHeight(comm.existingHOA)) addWarning(warnings, poleId, "", comm.owner, "INVALID_EXISTING_HOA", `Invalid Existing HOA for ${comm.owner}.`, "danger");
    });

    const connected = S().getConnectedSpans(poleId);
    if (!connected.length) addWarning(warnings, poleId, "", "", "DISCONNECTED_POLE", "The pole has no connected spans.");

    S().getSpanCommsForPole(poleId).forEach(sc => {
      if (sc.unknownOwner) addWarning(warnings, poleId, sc.spanId, sc.owner, "UNKNOWN_OWNER", `Unnormalized owner in Span.Wire: ${sc.rawOwner || sc.owner}.`);
      ["existingHOA", "existingHOAChange", "ocalcMS", "midspan", "calculatedMidspan", "msProposed", "finalMidspan"].forEach(field => {
        if (sc[field] && !H().isValidHeight(sc[field])) addWarning(warnings, poleId, sc.spanId, sc.owner, `INVALID_${field.toUpperCase()}`, `Invalid ${field} for ${sc.owner}.`, "danger");
      });
      const effectiveHOA = sc.existingHOAChange || sc.existingHOA;
      if (effectiveHOA && exceedsMax(effectiveHOA, pole.maxCommHeight)) {
        const label = sc.existingHOAChange ? "HOA Change" : "Existing HOA";
        addWarning(warnings, poleId, sc.spanId, sc.owner, "COMM_ABOVE_MAX", `${label} ${effectiveHOA} exceeds max height on pole.`, "danger");
      }
    });

    state.warnings = state.warnings.filter(w => w.poleId !== poleId).concat(warnings);
    return warnings;
  }

  function validateAll() {
    const state = S().getState();
    state.warnings = [];
    Object.keys(state.spans).forEach(validateSpan);
    Object.keys(state.poles).forEach(validatePole);
    return state.warnings;
  }

  global.Validations = {
    validatePole,
    validateSpan,
    validateAll
  };
})(window);
