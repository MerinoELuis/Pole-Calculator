(function (global) {
  "use strict";

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
      addWarning(warnings, span.fromPole, span.spanId, "", "EDITABLE_OTHER_POLE", "El other pole fue generado porque no venía ligado en Span. Puedes editar sus datos.");
    }

    const spanRows = S().getSpanCommsForSpan(spanId);
    if (!spanRows.length) addWarning(warnings, span.fromPole, span.spanId, "", "EMPTY_SPAN", "El span no tiene comms importados desde Span.Wire.");
    if (spanRows.length && !span.midspanLowPower) {
      addWarning(warnings, span.fromPole, span.spanId, "", "MISSING_MIDSPAN_POWER", "Falta midspan de power para calcular altura máxima en midspan.");
    }
    if (span.environment && span.environment !== "NONE") {
      if (!span.environmentClearance) addWarning(warnings, span.fromPole, span.spanId, "", "MISSING_ENV_CLEARANCE", "El span tiene environment pero falta clearance editable.");
      if (span.environmentClearance && span.environmentClearance !== "Variable" && !H().isValidHeight(span.environmentClearance)) {
        addWarning(warnings, span.fromPole, span.spanId, "", "INVALID_ENV_CLEARANCE", "Environment clearance inválido.", "danger");
      }
    }

    S().getSpanSidesForSpan(spanId).forEach(side => {
      if (side.proposedHOA && !H().isValidHeight(side.proposedHOA)) addWarning(warnings, side.poleId, spanId, "", "INVALID_PROPOSED", "Proposed inválido.", "danger");
      if (side.proposedHOAChange && !H().isValidHeight(side.proposedHOAChange)) addWarning(warnings, side.poleId, spanId, "", "INVALID_PROPOSED_CHANGE", "Cambio Proposed inválido.", "danger");
      if (side.ocalcMS && !H().isValidHeight(side.ocalcMS)) addWarning(warnings, side.poleId, spanId, "", "INVALID_OCALC_MS", "O-CALC MS inválido.", "danger");
      if (side.msProposed && !H().isValidHeight(side.msProposed)) addWarning(warnings, side.poleId, spanId, "", "INVALID_MS_PROPOSED", "MS Proposed inválido.", "danger");
      if (side.finalMidspan && !H().isValidHeight(side.finalMidspan)) addWarning(warnings, side.poleId, spanId, "", "INVALID_FINAL_MIDSPAN", "Midspan final ajustado inválido.", "danger");
      if (side.endDrop && !H().isValidHeight(side.endDrop)) addWarning(warnings, side.poleId, spanId, "", "INVALID_END_DROP", "End Drop inválido.", "danger");
      const otherPoleId = S().getOtherPoleId(span, side.poleId);
      const otherSide = otherPoleId ? S().getSpanSide(spanId, otherPoleId) : null;
      if (side.proposedHOA && !side.proposedHOAChange && !(otherSide?.proposedHOAChange || otherSide?.proposedHOA)) {
        addWarning(warnings, side.poleId, spanId, "", "MISSING_PROPOSED_TARGET", "Hay Proposed pero falta Cambio Proposed o Proposed del otro poste para calcular End Drop.");
      }
      if (side.proposedHOA && exceedsMax(side.proposedHOA, side.maxCommHeight)) addWarning(warnings, side.poleId, spanId, "", "PROPOSED_ABOVE_MAX", `Proposed ${side.proposedHOA} supera la altura máxima ${side.maxCommHeight}.`, "danger");
      if (side.clearanceMSIssue) {
        addWarning(warnings, side.poleId, spanId, "", "MS_PROPOSED_CLEARANCE", side.clearanceMSMessage || "MS Proposed tiene problema de clearance.", side.clearanceMSStatus === "PROBLEM" ? "danger" : "warning");
      }
    });

    spanRows.forEach(sc => {
      if (sc.calculatedMidspan && span.midspanMaxCommHeight && exceedsMax(sc.calculatedMidspan, span.midspanMaxCommHeight)) {
        addWarning(warnings, sc.poleId, spanId, sc.owner, "MIDSPAN_ABOVE_POWER_CLEARANCE", `Midspan ${sc.calculatedMidspan} supera la altura máxima en midspan ${span.midspanMaxCommHeight}.`, "danger");
      }
      if (sc.clearanceMSIssue) {
        addWarning(warnings, sc.poleId, spanId, sc.owner, "COMM_MIDSPAN_CLEARANCE", sc.clearanceMSMessage || "El comm no tiene suficiente espacio en midspan.", "danger");
      }
      if (sc.existingHOAChange && exceedsMax(sc.existingHOAChange, S().getPole(sc.poleId)?.maxCommHeight)) {
        addWarning(warnings, sc.poleId, spanId, sc.owner, "COMM_CHANGE_ABOVE_MAX", `Cambio de HOA ${sc.existingHOAChange} supera la altura máxima del poste.`, "danger");
      }
      if (sc.existingHOAChange && !global.MRLogic.generateMRForComm(sc)) {
        addWarning(warnings, sc.poleId, spanId, sc.owner, "CHANGE_WITHOUT_MR", "Hay Cambio de HOA pero no se generó MR.");
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

    if (pole.lowPower && !H().isValidHeight(pole.lowPower)) addWarning(warnings, poleId, "", "", "INVALID_LOW_POWER", "Low Power inválido.", "danger");
    if (!pole.lowPower && !pole.isGenerated) addWarning(warnings, poleId, "", "", "MISSING_LOW_POWER", "Falta Low Power del poste.");

    pole.comms.forEach(comm => {
      if (!comm.owner) addWarning(warnings, poleId, "", "", "UNKNOWN_OWNER", "Hay un comm sin owner.");
      if (comm.unknownOwner) addWarning(warnings, poleId, "", comm.owner, "UNKNOWN_OWNER", `Owner no normalizado: ${comm.rawOwner || comm.owner}.`);
      if (comm.existingHOA && !H().isValidHeight(comm.existingHOA)) addWarning(warnings, poleId, "", comm.owner, "INVALID_EXISTING_HOA", `Existing HOA inválido para ${comm.owner}.`, "danger");
    });

    const connected = S().getConnectedSpans(poleId);
    if (!connected.length) addWarning(warnings, poleId, "", "", "DISCONNECTED_POLE", "El poste no tiene spans conectados.");

    S().getSpanCommsForPole(poleId).forEach(sc => {
      if (sc.unknownOwner) addWarning(warnings, poleId, sc.spanId, sc.owner, "UNKNOWN_OWNER", `Owner no normalizado en Span.Wire: ${sc.rawOwner || sc.owner}.`);
      ["existingHOA", "existingHOAChange", "ocalcMS", "midspan", "calculatedMidspan", "msProposed", "finalMidspan"].forEach(field => {
        if (sc[field] && !H().isValidHeight(sc[field])) addWarning(warnings, poleId, sc.spanId, sc.owner, `INVALID_${field.toUpperCase()}`, `${field} inválido para ${sc.owner}.`, "danger");
      });
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
