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

  function normalizeOwnerForMatch(value) {
    const text = String(value || "")
      .replace(/^COMMUNICATION\s*>\s*/i, "")
      .replace(/,\s*.*$/, "")
      .toLowerCase()
      .normalize("NFD")
      .replace(/[\u0300-\u036f]/g, "")
      .replace(/[^a-z0-9]+/g, " ")
      .trim();

    if (/century\s*link|centurylink|\bctl\b|telco/.test(text)) return "ctl";
    if (/cable\s*one|cox|catv/.test(text)) return "catv";
    if (/vexus/.test(text)) return "vexus";
    if (/wecom/.test(text)) return "wecom";
    if (/fiber/.test(text)) return "fiber";
    return text;
  }

  function ownerTokens(scOrValue) {
    if (typeof scOrValue === "string") return new Set([normalizeOwnerForMatch(scOrValue)].filter(Boolean));
    return new Set([
      normalizeOwnerForMatch(scOrValue?.ownerBase),
      normalizeOwnerForMatch(scOrValue?.owner),
      normalizeOwnerForMatch(scOrValue?.rawOwner),
      normalizeOwnerForMatch(commOwnerLabel(scOrValue))
    ].filter(Boolean));
  }

  function ownersMatch(a, b) {
    const left = ownerTokens(a);
    const right = ownerTokens(b);
    return Array.from(left).some(token => right.has(token));
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
    // El midspan base viene de la columna Midspan de Span.Wire. O-CALC MS queda
    // solo como respaldo para archivos guardados viejos que ya traían ese campo.
    return parseMidspanValue(sc?.midspan || sc?.ocalcMS || "");
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

  function samePhysicalSpan(a, b) {
    if (!a || !b) return false;
    const aPoles = [a.fromPole, a.toPole].filter(Boolean).sort().join("|");
    const bPoles = [b.fromPole, b.toPole].filter(Boolean).sort().join("|");
    return Boolean(aPoles && bPoles && aPoles === bPoles);
  }

  function getLocalCommCandidate(spanId, poleId, ownerBase = "", preferredWireId = "") {
    const target = normalizeOwnerForMatch(ownerBase);
    const candidates = S().getSpanCommsForSpan(spanId).filter(row => row.poleId === poleId);
    const exactWire = candidates.find(row => preferredWireId && row.wireId && row.wireId === preferredWireId);
    if (exactWire) return exactWire;
    return candidates.find(row => target && ownerTokens(row).has(target)) || candidates[0] || null;
  }

  function candidateRank(candidate, context) {
    const candidateSpan = S().getSpan(candidate.spanId);
    const exactWire = Boolean(context.local?.wireId && candidate.wireId && candidate.wireId === context.local.wireId);
    const sameOwner = ownersMatch(candidate, context.local || context.ownerBase);
    if (!exactWire && !sameOwner) return 0;
    let score = 0;

    if (exactWire) score += 100;
    if (candidate.spanId === context.spanId) score += 45;
    if (samePhysicalSpan(candidateSpan, context.span)) score += 35;
    if (sameOwner) score += 20;
    if (getImportedMidspanInchesForComm(candidate) !== null) score += 5;

    return score;
  }

  function findRemoteComm(spanId, poleId, ownerBase = "", preferredWireId = "") {
    const span = S().getSpan(spanId);
    if (!span) return null;

    const otherPole = S().getOtherPoleId(span, poleId);
    if (!otherPole) return null;

    // Primero se busca por wireId porque un mismo owner puede tener varios
    // cables en el mismo poste. El owner queda como respaldo para datos viejos.
    const local = getLocalCommCandidate(spanId, poleId, ownerBase, preferredWireId) || {
      spanId,
      poleId,
      owner: ownerBase,
      ownerBase,
      wireId: preferredWireId
    };
    if (preferredWireId && !local.wireId) local.wireId = preferredWireId;

    const candidates = S().getSpanCommsForPole(otherPole);
    if (!candidates.length) return null;

    const context = { spanId, poleId, ownerBase, span, local };
    const ranked = candidates
      .map(row => ({ row, score: candidateRank(row, context) }))
      .filter(item => item.score > 0)
      .sort((a, b) => b.score - a.score);

    return ranked[0]?.row || null;
  }

  function findMidspanSourceComm(spanComm) {
    if (!spanComm) return null;
    const ownMidspan = getImportedMidspanInchesForComm(spanComm);
    if (ownMidspan !== null) return spanComm;

    const span = S().getSpan(spanComm.spanId);
    const rows = Object.values(S().getState().spanComms || {})
      .filter(row => spanCommKey(row) !== spanCommKey(spanComm))
      .filter(row => getImportedMidspanInchesForComm(row) !== null);

    const exactWire = rows.find(row => spanComm.wireId && row.wireId === spanComm.wireId);
    if (exactWire) return exactWire;

    const physicalOwnerMatch = rows.find(row => {
      const rowSpan = S().getSpan(row.spanId);
      return samePhysicalSpan(span, rowSpan) && ownersMatch(row, spanComm);
    });
    if (physicalOwnerMatch) return physicalOwnerMatch;

    const ownerMatch = rows.find(row => ownersMatch(row, spanComm));
    return ownerMatch || null;
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
        clearanceMSReason: "",
        message: "Missing O-CALC MS.",
        position: getSettingPosition()
      };
    }
    const commClearance = getMidspanCommCommClearance();
    const references = getReferenceMidspansForSpanSide(span.spanId, poleId);
    const maxMS = H().parseHeight(span.midspanMaxCommHeight || "");
    const envMin = getEnvironmentMinimum(span);
    const position = getSettingPosition();
    let target = baseInches;
    let issue = false;
    let needsAdjustment = false;
    let impossible = false;
    let clearanceMSReason = "";
    const messages = [];

    if (references.length) {
      const reference = position === "TOP_COMM" ? Math.max(...references) : Math.min(...references);
      const required = position === "TOP_COMM" ? reference + commClearance : reference - commClearance;
      if ((position === "TOP_COMM" && target < required) || (position === "LOW_COMM" && target > required)) {
        target = required;
        issue = true;
        needsAdjustment = true;
        messages.push(`Adjusted to keep ${format(commClearance)} from comm midspan ${format(reference)}.`);
      }
    }
    if (envMin !== null && target < envMin) {
      target = envMin;
      issue = true;
      needsAdjustment = true;
      messages.push(`Adjusted to environment minimum ${format(envMin)}.`);
    }
    if (maxMS !== null && target > maxMS) {
      target = maxMS;
      issue = true;
      needsAdjustment = true;
      clearanceMSReason = "LOW_POWER";
      messages.push(`Adjusted to max height at MS ${format(maxMS)}.`);
    }
    if (maxMS !== null && target > maxMS) impossible = true;
    return {
      baseFormatted: format(baseInches),
      finalFormatted: format(target),
      status: issue ? (impossible ? "PROBLEM" : "ADJUSTED") : "OK",
      issue,
      impossible,
      needsAdjustment,
      clearanceMSReason,
      message: messages.join(" ") || "OK",
      position
    };
  }

  function getOrderedReferenceMidspans(spanId) {
    return S().getSpanCommsForSpan(spanId)
      .map(sc => ({
        owner: commOwnerLabel(sc),
        poleHeight: H().parseHeight(getEffectiveCommHOA(sc)),
        midspan: getMidspanInchesForComm(sc)
      }))
      .filter(item => item.poleHeight !== null && item.midspan !== null);
  }

  function evaluateSpanSideFlagging(spanSide) {
    const pole = S().getPole(spanSide?.poleId);
    const proposed = H().parseHeight(spanSide?.proposedHOA || "");
    if (!spanSide || proposed === null) return { status: "OK", message: "" };
    const position = getSettingPosition();
    const topComm = H().parseHeight(pole?.topComm || "");
    const lowComm = H().parseHeight(pole?.lowComm || "");
    if (position === "TOP_COMM" && topComm !== null && proposed < topComm) {
      return { status: "PROBLEM", message: "Proposed below top comm." };
    }
    if (position === "LOW_COMM" && lowComm !== null && proposed > lowComm) {
      return { status: "PROBLEM", message: "Proposed above low comm." };
    }
    return { status: "OK", message: "OK" };
  }

  function applyDelayedMidspanResult(existing, evaluation) {
    return {
      msProposed: evaluation.baseFormatted,
      finalMidspan: evaluation.finalFormatted,
      pendingMidspanFinal: "",
      clearanceFixReadyAt: 0,
      clearanceMSStatus: evaluation.status,
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
    const proposedFlagging = evaluateSpanSideFlagging(side);
    S().upsertSpanSide({
      ...side,
      ...applyDelayedMidspanResult(side, evaluation),
      proposedFlaggingStatus: proposedFlagging.status,
      proposedFlaggingMessage: proposedFlagging.message
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

    if (midspan !== null && poleHeight !== null) {
      const ordered = getOrderedReferenceMidspans(sc.spanId);
      ordered.forEach(other => {
        if (!other.owner || other.owner === owner) return;
        const samePoleRows = S().getSpanCommsForSpan(sc.spanId).filter(row => commOwnerLabel(row) === other.owner);
        const otherAtPole = samePoleRows.find(row => row.poleId === sc.poleId);
        const otherPoleHeight = H().parseHeight(getEffectiveCommHOA(otherAtPole));
        if (otherPoleHeight === null) return;
        const poleRelation = Math.sign(poleHeight - otherPoleHeight);
        const midspanRelation = Math.sign(midspan - other.midspan);
        if (poleRelation && midspanRelation && poleRelation !== midspanRelation) {
          issues.push(`Midspan order crosses ${other.owner}.`);
        }
      });
    }

    if (midspan !== null && maxMS !== null && getSettingPosition() === "TOP_COMM") {
      const spanMidspans = S().getSpanCommsForSpan(sc.spanId)
        .map(getMidspanInchesForComm)
        .filter(value => value !== null);
      const isTopCommAtMidspan = spanMidspans.length && midspan === Math.max(...spanMidspans);
      if (isTopCommAtMidspan && midspan + midspanClearance > maxMS) {
        issues.push("No room for proposed above top comm.");
      }
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
    const span = S().getSpan(spanId);
    if (!side || !span) return "";
    const proposed = H().parseHeight(side.proposedHOA || "");
    const manualNext = side.nextPoleProposedAuto ? "" : side.proposedHOAChange;
    let nextPoleProposed = H().parseHeight(manualNext || "");
    let debugOtherPoleId = "";
    let debugCandidates = [];
    let debugSelected = null;
    if (nextPoleProposed === null) {
      const otherPoleId = S().getOtherPoleId(span, poleId);
      // El proposed del poste siguiente no siempre vive en el mismo spanId.
      // Para este flujo se toma el primer proposed que el poste siguiente
      // realmente propone hacia adelante; si no existe, se deja vacío.
      const otherSides = otherPoleId ? S().getSpanSidesForPole(otherPoleId) : [];
      debugOtherPoleId = otherPoleId || "";
      debugCandidates = otherSides.map(otherSide => {
        const candidateSpan = S().getSpan(otherSide.spanId);
        return {
          spanId: otherSide.spanId,
          fromPole: candidateSpan?.fromPole || "",
          proposedHOA: otherSide.proposedHOA || ""
        };
      });
      const forwardProposedSide = otherSides.find(otherSide => {
        const candidateSpan = S().getSpan(otherSide.spanId);
        return Boolean(otherSide.proposedHOA && candidateSpan?.fromPole === otherPoleId);
      });
      const sourceSide = forwardProposedSide || null;
      debugSelected = sourceSide ? { spanId: sourceSide.spanId, proposedHOA: sourceSide.proposedHOA } : null;
      nextPoleProposed = H().parseHeight(sourceSide?.proposedHOA || "");
      side = S().upsertSpanSide({
        ...side,
        proposedHOAChange: nextPoleProposed !== null ? format(nextPoleProposed) : "",
        nextPoleProposedAuto: nextPoleProposed !== null
      });
    }

    const endDrop = proposed !== null && nextPoleProposed !== null ? format(nextPoleProposed - proposed) : "";
    S().upsertSpanSide({ ...side, endDrop });
    if (side.proposedHOA) {
      console.debug("[PoleCalc next proposed]", {
        spanId,
        poleId,
        proposed: side.proposedHOA || "",
        manualNext: manualNext || "",
        otherPoleId: debugOtherPoleId,
        candidates: debugCandidates,
        selected: debugSelected,
        nextPoleProposed: nextPoleProposed !== null ? format(nextPoleProposed) : "",
        endDrop
      });
    }
    return endDrop;
  }

  function calculateProposedMidspanBase(side, span) {
    return parseMidspanValue(side?.ocalcMS || side?.proposedMidspan || "");
  }

  function calculateMidspanForComm(spanComm) {
    if (!spanComm) return "";
    const span = S().getSpan(spanComm.spanId);
    if (!span) return spanComm.midspan || "";

    const localExisting = H().parseHeight(spanComm.existingHOA);
    const localCurrent = H().parseHeight(getEffectiveCommHOA(spanComm));
    const remote = findRemoteComm(spanComm.spanId, spanComm.poleId, spanComm.ownerBase || spanComm.owner, spanComm.wireId || "");
    const remoteHOA = remote ? getEffectiveCommHOA(remote) : "";
    const remoteExisting = H().parseHeight(remote?.existingHOA || "");
    const remoteCurrent = H().parseHeight(remoteHOA);
    const midspanSource = findMidspanSourceComm(spanComm) || spanComm;
    const importedMidspan = getImportedMidspanInchesForComm(midspanSource);

    let calculated = "";
    let localAdjustment = 0;
    let remoteAdjustment = 0;
    if (importedMidspan !== null) {
      // Fórmula documentada para comms:
      // Midspan nuevo = Midspan importado + (movimiento local / 2) + (movimiento remoto / 2).
      // Movimiento = HOA efectivo - Existing HOA.
      // El HOA efectivo es Cambio de HOA cuando existe; si no, Existing HOA.
      // Ejemplo: 20' -> 19' aporta -6" al midspan. Si el otro poste pasa
      // de 20' -> 21', aporta +6", y ambos cambios se compensan.
      localAdjustment = localExisting !== null && localCurrent !== null ? (localCurrent - localExisting) / 2 : 0;
      remoteAdjustment = remoteExisting !== null && remoteCurrent !== null ? (remoteCurrent - remoteExisting) / 2 : 0;
      calculated = format(Math.round(importedMidspan + localAdjustment + remoteAdjustment));
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

    if (field === "endDrop") data.lockedEndDrop = Boolean(value);
    if (["ocalcMS", "proposedMidspan", "proposedHOA", "proposedHOAChange"].includes(field)) {
      data.clearanceFixReadyAt = 0;
      data.pendingMidspanFinal = "";
      data.msProposed = "";
      data.finalMidspan = "";
      data.clearanceMSStatus = "";
      data.clearanceMSMessage = "";
      data.clearanceMSReason = "";
      data.clearanceMSIssue = false;
    }

    if (field === "proposedHOAChange") data.nextPoleProposedAuto = false;
    S().upsertSpanSide(data);
    if (field === "proposedHOA") {
      console.debug("[PoleCalc proposed]", { spanId, poleId, proposedHOA: value || "" });
    }
    if (["proposedHOA", "proposedHOAChange"].includes(field)) calculateEndDropForSpanSide(spanId, poleId);
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

    // Recalcula también comms de spans recíprocos por Wire Id. Esto cubre casos
    // Fore/Back importados con distinto Span Id, pero pertenecientes al mismo cable.
    const affectedWireIds = new Set(
      S().getSpanCommsForPole(poleId)
        .map(sc => sc.wireId)
        .filter(Boolean)
    );
    Object.values(S().getState().spanComms || {}).forEach(sc => {
      if (affectedWireIds.has(sc.wireId)) calculateMidspanForComm(sc);
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
    evaluateSpanSideFlagging,
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
