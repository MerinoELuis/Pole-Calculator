(function (global) {
  "use strict";

  const H = () => global.HeightUtils;
  const S = () => global.AppStore;
  const CLEARANCE_FIX_DELAY_MS = 700;

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

  function getSettingPosition() {
    const value = String(S().getState().settings?.position || "TOP_COMM").toUpperCase();
    return value === "LOW_COMM" ? "LOW_COMM" : "TOP_COMM";
  }

  function getMidspanCommCommClearance() {
    const settings = S().getState().settings || {};
    return H().parseHeight(settings.midspanCommCommClearance || "4\"") ?? 4;
  }

  function getPoleCommCommClearance() {
    const settings = S().getState().settings || {};
    return H().parseHeight(settings.commClearance || "12\"") ?? 12;
  }

  function getPoleBoltBoltClearance() {
    const settings = S().getState().settings || {};
    return H().parseHeight(settings.boltClearance || "4\"") ?? 4;
  }

  function commOwnerLabel(sc) {
    return String(sc?.rawOwner || sc?.owner || "")
      .replace(/^COMMUNICATION\s*>\s*/i, "")
      .trim();
  }

  function normalizedHeightLabelForCalc(value) {
    const parsed = H().parseHeight(value || "");
    return parsed === null ? String(value || "").trim() : format(parsed);
  }

  function parseMidspanValue(value) {
    if (value === null || value === undefined || value === "") return null;
    const text = String(value).trim();
    if (/^[-−]?\d+(?:\.\d+)?$/.test(text)) {
      return Math.round(Number(text.replace("−", "-")) * 12);
    }
    return H().parseHeight(value);
  }

  function displayMidspanForComm(sc) {
    if (!sc) return "";
    const value = sc.finalMidspan || sc.msProposed || sc.calculatedMidspan || sc.midspan || sc.ocalcMS || "";
    const parsed = parseMidspanValue(value);
    return parsed === null ? value : format(parsed);
  }

  function getMidspanInchesForComm(sc) {
    if (!sc) return null;
    return parseMidspanValue(sc.finalMidspan || sc.msProposed || sc.calculatedMidspan || sc.midspan || sc.ocalcMS || "");
  }

  function getImportedMidspanInchesForComm(sc) {
    return parseMidspanValue(sc?.ocalcMS || sc?.midspan || "");
  }

  function spanCommKey(sc) {
    return S().keyForSpanComm(sc.spanId, sc.poleId, sc.owner, sc.wireId || "");
  }

  function getReferenceMidspansForSpan(spanId, excludeKey = "") {
    return S().getSpanCommsForSpan(spanId)
      .filter(sc => spanCommKey(sc) !== excludeKey)
      .map(getMidspanInchesForComm)
      .filter(value => value !== null);
  }

  function getReferenceMidspansForSpanSide(spanId, poleId) {
    const rows = poleId ? S().getSpanCommsForPole(poleId) : S().getSpanCommsForSpan(spanId);
    const values = rows
      .map(getMidspanInchesForComm)
      .filter(value => value !== null);
    return Array.from(new Set(values));
  }

  function getEnvironmentMinimum(span) {
    if (!span || !span.environmentClearance || span.environmentClearance === "Variable") return null;
    return H().parseHeight(span.environmentClearance);
  }

  function evaluateSpanSideMidspan(baseInches, span, poleId = "") {
    if (baseInches === null || !span) {
      return {
        baseFormatted: "",
        finalFormatted: "",
        status: "MISSING",
        issue: false,
        impossible: false,
        needsAdjustment: false,
        message: "Falta O-CALC MS."
      };
    }

    const settings = S().getState().settings || {};
    const position = getSettingPosition();
    const commClearance = getMidspanCommCommClearance();
    const references = getReferenceMidspansForSpanSide(span.spanId, poleId);
    const maxMS = H().parseHeight(span.midspanMaxCommHeight);
    const envMin = getEnvironmentMinimum(span);
    const messages = [];
    let target = baseInches;
    let issue = false;
    let impossible = false;
    let needsAdjustment = false;
    let clearanceMSReason = "";

    if (references.length) {
      messages.push(`Referencia tomada contra ${references.length} comm(s)/span(s) conectados al poste.`);
      if (position === "TOP_COMM") {
        const reference = Math.max(...references);
        const required = reference + commClearance;
        if (target < required) {
          issue = true;
          needsAdjustment = true;
          target = required;
          messages.push(`Requiere ${format(required)} para quedar ${format(commClearance)} arriba de ${format(reference)}.`);
        }
      } else {
        const reference = Math.min(...references);
        const required = reference - commClearance;
        if (target > required) {
          issue = true;
          needsAdjustment = true;
          target = required;
          messages.push(`Requiere ${format(required)} para quedar ${format(commClearance)} abajo de ${format(reference)}.`);
        }
      }
    }

    if (envMin !== null && target < envMin) {
      issue = true;
      needsAdjustment = true;
      target = envMin;
      messages.push(`Se ajusta al clearance de environment ${format(envMin)}.`);
    }

    if (maxMS !== null && target > maxMS) {
      issue = true;
      needsAdjustment = true;
      target = maxMS;
      clearanceMSReason = "LOW_POWER";
      messages.push(`Se ajusta a la altura máxima en midspan ${format(maxMS)}.`);
    }

    if (references.length) {
      if (position === "TOP_COMM") {
        const reference = Math.max(...references);
        const required = reference + commClearance;
        if (target < required) {
          impossible = true;
          issue = true;
          messages.push(`No hay espacio suficiente: para Top Comm necesita ${format(required)} y el máximo disponible es ${maxMS !== null ? format(maxMS) : "sin dato"}.`);
        }
      } else {
        const reference = Math.min(...references);
        const required = reference - commClearance;
        if (target > required) {
          impossible = true;
          issue = true;
          messages.push(`No hay espacio suficiente: para Low Comm necesita ${format(required)} o menos.`);
        }
      }
    }

    if (maxMS !== null && target > maxMS) {
      impossible = true;
      issue = true;
      messages.push(`El midspan final supera Max MS ${format(maxMS)}.`);
    }

    return {
      baseFormatted: format(baseInches),
      finalFormatted: format(target),
      status: issue ? (impossible ? "PROBLEM" : "ADJUSTMENT_NEEDED") : "OK",
      issue,
      impossible,
      needsAdjustment,
      clearanceMSReason,
      message: messages.join(" ") || `OK con Midspan Comm-Comm ${format(commClearance)}.`,
      position
    };
  }

  function applyDelayedMidspanResult(existing, evaluation) {
    const now = Date.now();
    let msProposed = evaluation.baseFormatted;
    let finalMidspan = evaluation.finalFormatted;
    let pendingMidspanFinal = "";
    let clearanceFixReadyAt = 0;
    let status = evaluation.status;

    if (evaluation.needsAdjustment) {
      const previousReadyAt = Number(existing.clearanceFixReadyAt || 0);
      if (previousReadyAt && now >= previousReadyAt) {
        msProposed = evaluation.finalFormatted;
        finalMidspan = evaluation.finalFormatted;
        status = evaluation.impossible ? "PROBLEM" : "ADJUSTED";
      } else {
        msProposed = evaluation.baseFormatted;
        finalMidspan = evaluation.baseFormatted;
        pendingMidspanFinal = evaluation.finalFormatted;
        clearanceFixReadyAt = previousReadyAt || now + CLEARANCE_FIX_DELAY_MS;
        status = "PENDING";
      }
    }

    return {
      msProposed,
      finalMidspan,
      pendingMidspanFinal,
      clearanceFixReadyAt,
      clearanceMSStatus: status,
      clearanceMSMessage: evaluation.message,
      clearanceMSReason: evaluation.clearanceMSReason || "",
      clearanceMSIssue: Boolean(evaluation.issue)
    };
  }

  function calculateSpanSideMidspan(spanId, poleId) {
    const side = S().getSpanSide(spanId, poleId);
    const span = S().getSpan(spanId);
    if (!side || !span) return null;

    const baseInches = calculateProposedMidspanBase(side, span);
    const evaluation = evaluateSpanSideMidspan(baseInches, span, poleId);
    const delayed = applyDelayedMidspanResult(side, evaluation);

    S().upsertSpanSide({
      ...side,
      ...delayed
    });
    return S().getSpanSide(spanId, poleId);
  }

  function evaluateCommMidspanClearance(sc, calculatedMidspan) {
    const span = S().getSpan(sc.spanId);
    const midspan = H().parseHeight(calculatedMidspan || displayMidspanForComm(sc));
    const maxMS = H().parseHeight(span?.midspanMaxCommHeight || "");
    const commClearance = getMidspanCommCommClearance();

    if (midspan === null) {
      return {
        msProposed: "",
        finalMidspan: "",
        clearanceMSStatus: "MISSING",
        clearanceMSMessage: "Falta midspan del comm.",
        clearanceMSIssue: false
      };
    }

    if (maxMS === null) {
      return {
        msProposed: format(midspan),
        finalMidspan: format(midspan),
        clearanceMSStatus: "MISSING_POWER",
        clearanceMSMessage: "Falta Max MS Comm / Low Power en midspan.",
        clearanceMSIssue: false
      };
    }

    const requiredWithClearance = midspan + commClearance;
    const ok = requiredWithClearance <= maxMS;
    return {
      msProposed: format(midspan),
      finalMidspan: format(midspan),
      clearanceMSStatus: ok ? "OK" : "PROBLEM",
      clearanceMSMessage: ok
        ? `${format(midspan)} + ${format(commClearance)} <= ${format(maxMS)}.`
        : `${format(midspan)} + ${format(commClearance)} supera el límite ${format(maxMS)}.`,
      clearanceMSIssue: !ok
    };
  }

  function evaluateCommFlagging(sc, calculatedMidspan) {
    const span = S().getSpan(sc.spanId);
    const pole = S().getPole(sc.poleId);
    const midspan = H().parseHeight(calculatedMidspan || displayMidspanForComm(sc));
    const poleHeight = H().parseHeight(getEffectiveCommHOA(sc));
    const owner = commOwnerLabel(sc);
    const issues = [];

    const envMin = getEnvironmentMinimum(span);
    if (midspan !== null && envMin !== null && midspan < envMin) {
      issues.push(`Environment: ${format(midspan)} < ${format(envMin)}.`);
    }

    const maxMS = H().parseHeight(span?.midspanMaxCommHeight || "");
    if (midspan !== null && maxMS !== null && midspan > maxMS) {
      issues.push(`Power MS: ${format(midspan)} > max ${format(maxMS)}.`);
    }

    const midspanClearance = getMidspanCommCommClearance();
    if (midspan !== null) {
      S().getSpanCommsForSpan(sc.spanId).forEach(other => {
        if (spanCommKey(other) === spanCommKey(sc)) return;
        if (commOwnerLabel(other) && commOwnerLabel(other) === owner) return;
        const otherMidspan = getMidspanInchesForComm(other);
        if (otherMidspan === null) return;
        const diff = Math.abs(midspan - otherMidspan);
        if (diff < midspanClearance) {
          issues.push(`Comm-comm MS: ${format(diff)} con ${commOwnerLabel(other) || "sin owner"}; mínimo ${format(midspanClearance)}.`);
        }
      });
    }

    if (poleHeight !== null) {
      const poleClearance = getPoleCommCommClearance();
      S().getSpanCommsForPole(sc.poleId).forEach(other => {
        if (spanCommKey(other) === spanCommKey(sc)) return;
        const otherOwner = commOwnerLabel(other);
        const sameOwner = otherOwner && otherOwner === owner;
        const otherEffective = getEffectiveCommHOA(other);
        if (sameOwner && normalizedHeightLabelForCalc(otherEffective) === normalizedHeightLabelForCalc(getEffectiveCommHOA(sc))) return;
        const otherHeight = H().parseHeight(getEffectiveCommHOA(other));
        if (otherHeight === null) return;
        const diff = Math.abs(poleHeight - otherHeight);
        const required = sameOwner ? getPoleBoltBoltClearance() : poleClearance;
        const label = sameOwner ? "Bolt-bolt poste" : "Comm-comm poste";
        if (diff < required) {
          issues.push(`${label}: ${format(diff)} con ${otherOwner || "sin owner"}; mínimo ${format(required)}.`);
        }
      });
    }

    const maxPole = H().parseHeight(pole?.maxCommHeight || "");
    const changedPoleHeight = H().parseHeight(sc.existingHOAChange || "");
    if (changedPoleHeight !== null && maxPole !== null && changedPoleHeight > maxPole) {
      issues.push(`Poste: Cambio de HOA ${format(changedPoleHeight)} supera max ${format(maxPole)}.`);
    }

    return {
      flaggingStatus: issues.length ? "PROBLEM" : "OK",
      flaggingMessage: issues.length ? Array.from(new Set(issues)).join(" ") : "OK"
    };
  }

  function findRemoteComm(spanId, poleId, ownerBase) {
    const span = S().getSpan(spanId);
    if (!span) return null;
    const otherPole = S().getOtherPoleId(span, poleId);
    if (!otherPole) return null;
    const target = String(ownerBase || "").toLowerCase();
    const candidates = S().getSpanCommsForSpan(spanId).filter(row => row.poleId === otherPole);
    return candidates.find(row => {
      const values = [row.ownerBase, row.owner, row.rawOwner].map(value => String(value || "").toLowerCase());
      return values.includes(target) || values.some(value => target && value.includes(target));
    }) || candidates[0] || null;
  }

  function findMidspanSourceComm(spanComm) {
    // Some exports only place the imported midspan on one side of the span.
    // The opposite pole still needs to reference that same physical cable so
    // movements on either end update the single shared midspan calculation.
    const ownMidspan = getImportedMidspanInchesForComm(spanComm);
    if (ownMidspan !== null) return spanComm;
    const target = String(spanComm.ownerBase || spanComm.owner || "").toLowerCase();
    return S().getSpanCommsForSpan(spanComm.spanId).find(row => {
      if (spanCommKey(row) === spanCommKey(spanComm)) return false;
      if (getImportedMidspanInchesForComm(row) === null) return false;
      const values = [row.ownerBase, row.owner, row.rawOwner].map(value => String(value || "").toLowerCase());
      return values.includes(target) || values.some(value => target && value.includes(target));
    }) || null;
  }

  function evaluateProposedBoltClearance(spanSide) {
    const proposed = H().parseHeight(spanSide?.proposedHOA || "");
    if (proposed === null) return { ok: true, message: "" };
    const required = getPoleBoltBoltClearance();
    const conflicts = S().getSpanCommsForPole(spanSide.poleId)
      .map(sc => ({ owner: commOwnerLabel(sc) || "sin owner", height: H().parseHeight(getEffectiveCommHOA(sc)) }))
      .filter(item => item.height !== null)
      .filter(item => {
        const diff = Math.abs(proposed - item.height);
        return diff < required;
      });
    if (!conflicts.length) return { ok: true, message: "" };
    const detail = conflicts
      .map(item => `${item.owner} ${format(item.height)}`)
      .join(", ");
    return {
      ok: false,
      message: `Proposed ${format(proposed)} no respeta Pole · Bolt-bolt ${format(required)} contra ${detail}.`
    };
  }

  function calculateSpanPowerDerived(spanId) {
    const span = S().getSpan(spanId);
    if (!span) return null;
    const settings = S().getState().settings || {};
    const clearance = H().parseHeight(settings.midspanPowerCommClearance || "30\"");
    const powerHeights = S().getSpanPowerForSpan(spanId)
      .map(row => H().parseHeight(row.midspan))
      .filter(value => value !== null);
    const midspanLowPower = powerHeights.length ? Math.min(...powerHeights) : null;
    const midspanMaxCommHeight = midspanLowPower !== null && clearance !== null ? midspanLowPower - clearance : null;
    S().updateSpanField(spanId, "midspanLowPower", midspanLowPower !== null ? format(midspanLowPower) : "");
    S().updateSpanField(spanId, "midspanMaxCommHeight", midspanMaxCommHeight !== null ? format(midspanMaxCommHeight) : "");
    return S().getSpan(spanId);
  }

  function nextPoleProposedValue(side) {
    if (!side) return "";
    return side.proposedHOA || "";
  }

  function clearAutoNextPoleProposed(side) {
    if (!side || !side.nextPoleProposedAuto) return side;
    return S().upsertSpanSide({ ...side, proposedHOAChange: "", nextPoleProposedAuto: false });
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
    const settings = S().getState().settings || {};
    const clearance = H().parseHeight(settings.polePowerCommsClearance || settings.clearanceToPower || "40\"");
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
    let side = S().getSpanSide(spanId, poleId);
    if (!side) return "";
    const span = S().getSpan(spanId);
    const proposed = H().parseHeight(side.proposedHOA);
    const userValue = side.nextPoleProposedAuto ? "" : side.proposedHOAChange;
    let target = H().parseHeight(userValue);
    let targetText = userValue;

    // Next Pole Proposed is a helper value, not local data. Only copy it from
    // the connected side when that connected pole has a real Proposed value.
    if (target === null && span) {
      const otherPoleId = S().getOtherPoleId(span, poleId);
      const otherSide = otherPoleId ? S().getSpanSide(spanId, otherPoleId) : null;
      const otherProposed = nextPoleProposedValue(otherSide);
      target = H().parseHeight(otherProposed);
      targetText = otherProposed;
      if (target !== null && targetText) {
        side = S().upsertSpanSide({ ...side, proposedHOAChange: H().formatHeight(target), nextPoleProposedAuto: true });
      } else {
        side = clearAutoNextPoleProposed(side);
      }
    }
    if (proposed === null || target === null) {
      if (!side.lockedEndDrop) S().upsertSpanSide({ ...side, endDrop: "", ...(target === null ? { proposedHOAChange: side.nextPoleProposedAuto ? "" : side.proposedHOAChange, nextPoleProposedAuto: false } : {}) });
      return side.endDrop || "";
    }
    const calculated = format(target - proposed);
    S().upsertSpanSide({ ...side, endDrop: calculated, lockedEndDrop: false });
    return calculated;
  }

  function calculateProposedMidspanBase(side, span) {
    if (!side) return null;
    const importedMidspan = parseMidspanValue(side.ocalcMS || side.proposedMidspan || "");
    if (importedMidspan === null) return null;

    const localBase = H().parseHeight(side.proposedHOA);
    const localNew = H().parseHeight(side.proposedHOA || "");
    let localAdjustment = 0;
    if (localBase !== null && localNew !== null) {
      localAdjustment = localNew - localBase;
    }

    let remoteAdjustment = 0;
    if (span) {
      const otherPoleId = S().getOtherPoleId(span, side.poleId);
      const otherSide = otherPoleId ? S().getSpanSide(span.spanId, otherPoleId) : null;
      const remoteBase = H().parseHeight(otherSide?.proposedHOA || "");
      const remoteNew = H().parseHeight((side.proposedHOAChange && !side.nextPoleProposedAuto ? side.proposedHOAChange : "") || otherSide?.proposedHOA || "");
      if (remoteBase !== null && remoteNew !== null) {
        remoteAdjustment = remoteNew - remoteBase;
      }
    }

    return Math.round(importedMidspan + ((localAdjustment + remoteAdjustment) / 2));
  }

  function calculateMidspanForComm(spanComm) {
    if (!spanComm) return "";
    const span = S().getSpan(spanComm.spanId);
    if (!span) return spanComm.midspan || spanComm.ocalcMS || "";

    const local = H().parseHeight(getEffectiveCommHOA(spanComm));
    const localExisting = H().parseHeight(spanComm.existingHOA);
    const remote = findRemoteComm(spanComm.spanId, spanComm.poleId, spanComm.ownerBase || spanComm.owner);
    const remoteHOA = remote ? getEffectiveCommHOA(remote) : "";
    const remoteInches = H().parseHeight(remoteHOA);
    const remoteExisting = H().parseHeight(remote?.existingHOA || "");
    const midspanSource = findMidspanSourceComm(spanComm) || spanComm;
    const importedMidspan = getImportedMidspanInchesForComm(midspanSource);

    let calculated = "";
    if (importedMidspan !== null) {
      // A comm midspan is affected by both ends of the same span. Each pole
      // movement contributes half of its height delta to the midspan; this is
      // why recalculating one pole also refreshes its connected/reference pole.
      const sourceIsRemote = spanCommKey(midspanSource) !== spanCommKey(spanComm);
      const sourceExisting = sourceIsRemote ? remoteExisting : localExisting;
      const sourceCurrent = sourceIsRemote ? remoteInches : local;
      const otherExisting = sourceIsRemote ? localExisting : remoteExisting;
      const otherCurrent = sourceIsRemote ? local : remoteInches;
      const sourceAdjustment = sourceExisting !== null && sourceCurrent !== null ? (sourceExisting - sourceCurrent) / 2 : 0;
      const otherAdjustment = otherExisting !== null && otherCurrent !== null ? (otherExisting - otherCurrent) / 2 : 0;
      calculated = format(Math.round(importedMidspan - sourceAdjustment - otherAdjustment));
    } else {
      calculated = "";
    }

    const difference = H().diffLabel(spanComm.existingHOA, spanComm.existingHOAChange || spanComm.existingHOA);
    const clearance = evaluateCommMidspanClearance(spanComm, calculated);
    const flagging = evaluateCommFlagging(spanComm, calculated);
    S().upsertSpanComm({
      ...spanComm,
      remotePoleId: remote ? remote.poleId : S().getOtherPoleId(span, spanComm.poleId),
      remoteHOA,
      calculatedMidspan: calculated,
      difference,
      ...clearance,
      ...flagging
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
    const span = S().getSpan(spanId);
    const affectedPoles = span ? [span.fromPole, span.toPole].filter(Boolean) : [poleId];
    affectedPoles.forEach(recalculateSpansForPole);
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
    const allowed = ["proposedHOA", "proposedHOAChange", "proposedMidspan", "ocalcMS", "endDrop", "clearanceReference", "notes"];
    if (!allowed.includes(field)) return null;
    const side = S().getSpanSide(spanId, poleId) || S().upsertSpanSide({ spanId, poleId });
    const data = { ...side, [field]: value || "" };
    if (field === "proposedHOAChange") data.nextPoleProposedAuto = false;
    if (field === "proposedHOA") {
      const span = S().getSpan(spanId);
      const otherPoleId = span ? S().getOtherPoleId(span, poleId) : "";
      const otherSide = otherPoleId ? S().getSpanSide(spanId, otherPoleId) : null;
      if (otherSide?.nextPoleProposedAuto) S().upsertSpanSide({ ...otherSide, proposedHOAChange: value || "", nextPoleProposedAuto: Boolean(value) });
    }
    if (field === "endDrop") data.lockedEndDrop = Boolean(value);
    if (field === "ocalcMS" || field === "proposedMidspan") {
      data.clearanceFixReadyAt = 0;
      data.pendingMidspanFinal = "";
    }
    S().upsertSpanSide(data);
    if (["proposedHOA", "proposedHOAChange", "proposedMidspan", "ocalcMS"].includes(field)) calculateEndDropForSpanSide(spanId, poleId);
    recalculateSpansForPole(poleId);
    return S().getSpanSide(spanId, poleId);
  }

  function updateSpanField(spanId, field, value) {
    const span = S().updateSpanField(spanId, field, value);
    if (!span) return null;
    S().getSpanSidesForSpan(spanId).forEach(side => {
      S().upsertSpanSide({ ...side, clearanceFixReadyAt: 0, pendingMidspanFinal: "" });
    });
    recalculateSpan(spanId);
    return span;
  }

  function updateSpanCommField(spanId, poleId, owner, wireId, field, value) {
    const allowed = ["existingHOA", "existingHOAChange", "serviceDrop", "ocalcMS", "midspan", "notes", "mr"];
    if (!allowed.includes(field)) return null;
    const sc = S().getSpanComm(spanId, poleId, owner, wireId) || S().upsertSpanComm({ spanId, poleId, owner, wireId });
    S().upsertSpanComm({ ...sc, [field]: value || "" });
    recalculateSpan(spanId);
    const span = S().getSpan(spanId);
    const affectedPoles = span ? [span.fromPole, span.toPole].filter(Boolean) : [poleId];
    affectedPoles.forEach(recalculateSpansForPole);
    return S().getSpanComm(spanId, poleId, owner, wireId);
  }

  function updateOcalcValue(poleId, owner, spanId, field, value) {
    if (field === "proposedHOA" || field === "proposedMidspan" || field === "ocalcMS" || field === "endDrop") return updateSpanSideField(spanId, poleId, field, value);
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
    // Recalculate the span as a graph edge: power limits first, then both pole
    // derivations, then every comm/proposed calculation tied to the two ends.
    const span = S().getSpan(spanId);
    calculateSpanPowerDerived(spanId);
    if (span) {
      if (span.fromPole) calculatePoleDerived(span.fromPole);
      if (span.toPole) calculatePoleDerived(span.toPole);
    }
    S().getSpanCommsForSpan(spanId).forEach(sc => calculateMidspanForComm(sc));
    S().getSpanSidesForSpan(spanId).forEach(side => {
      calculateEndDropForSpanSide(side.spanId, side.poleId);
      calculateSpanSideMidspan(side.spanId, side.poleId);
    });
  }

  function recalculateSpansForPole(poleId) {
    // A pole edit can affect forespans and backspans. This walks every
    // connected span so reference rows on the previous/next pole are refreshed.
    const spans = S().getConnectedSpans(poleId);
    spans.forEach(span => calculateSpanPowerDerived(span.spanId));

    calculatePoleDerived(poleId);
    spans.forEach(span => {
      const other = S().getOtherPoleId(span, poleId);
      if (other) calculatePoleDerived(other);
    });

    spans.forEach(span => {
      S().getSpanCommsForSpan(span.spanId).forEach(sc => calculateMidspanForComm(sc));
    });

    spans.forEach(span => {
      S().getSpanSidesForSpan(span.spanId).forEach(side => {
        calculateEndDropForSpanSide(side.spanId, side.poleId);
        calculateSpanSideMidspan(side.spanId, side.poleId);
      });
    });

    global.MRLogic.generateMRForPole(poleId);
    global.Validations.validatePole(poleId);
    spans.forEach(span => global.Validations.validateSpan(span.spanId));
  }

  function calculateOcalcValues() {
    const state = S().getState();
    Object.keys(state.spans).forEach(calculateSpanPowerDerived);
    Object.keys(state.poles).forEach(calculatePoleDerived);
    Object.values(state.spanComms).forEach(sc => calculateMidspanForComm(sc));
    Object.values(state.spanSides).forEach(side => {
      calculateEndDropForSpanSide(side.spanId, side.poleId);
      calculateSpanSideMidspan(side.spanId, side.poleId);
    });
  }

  function recalculateAll() {
    const state = S().getState();
    Object.keys(state.spans).forEach(calculateSpanPowerDerived);
    Object.keys(state.poles).forEach(calculatePoleDerived);
    Object.values(state.spanComms).forEach(sc => calculateMidspanForComm(sc));
    Object.values(state.spanSides).forEach(side => {
      calculateEndDropForSpanSide(side.spanId, side.poleId);
      calculateSpanSideMidspan(side.spanId, side.poleId);
    });
    global.MRLogic.generateAllMR();
    global.Validations.validateAll();
  }

  global.Calculations = {
    CLEARANCE_FIX_DELAY_MS,
    updatePoleCommNewHeight,
    updateExistingHOAChange,
    updateProposedForSpan,
    updateCambioProposed,
    updateSpanSideField,
    updateSpanField,
    updateSpanCommField,
    updateOcalcValue,
    calculateEndDrop,
    calculateEndDropForSpanSide,
    calculateMidspanForComm,
    calculateProposedMidspanBase,
    calculateSpanSideMidspan,
    calculateSpanPowerDerived,
    evaluateSpanSideMidspan,
    evaluateCommMidspanClearance,
    evaluateCommFlagging,
    evaluateProposedBoltClearance,
    commOwnerLabel,
    displayMidspanForComm,
    getConnectedSpans,
    getBackspanForPole,
    recalculateSpan,
    recalculateSpansForPole,
    calculateOcalcValues,
    recalculateAll,
    getEffectiveCommHOA,
    getEstimatedSagInches,
    findRemoteComm,
    getReferenceMidspansForSpanSide
  };
})(window);
