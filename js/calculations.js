(function (global) {
  "use strict";

  // calculations.js contains derived business logic. It recalculates spans,
  // comm movements, proposed values, clearances, end drops, and flagging.
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
    // A changed HOA replaces the imported HOA for downstream calculations.
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
    // only as a fallback for older saved files that already had that field.
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
    // wires on the same pole. Owner stays as the fallback for older data.
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
    // Proposed flagging is intentionally compact: one field summarizes every
    // pole-side rule that can make a proposed attachment invalid.
    const pole = S().getPole(spanSide?.poleId);
    const proposed = H().parseHeight(spanSide?.proposedHOA || "");
    if (!spanSide || proposed === null) return { status: "OK", message: "" };
    const position = getSettingPosition();
    const topComm = H().parseHeight(pole?.topComm || "");
    const lowComm = H().parseHeight(pole?.lowComm || "");
    const maxPole = H().parseHeight(pole?.maxCommHeight || spanSide.maxCommHeight || "");
    const issues = [];
    if (maxPole !== null && proposed > maxPole) {
      issues.push(`Proposed ${format(proposed)} exceeds Max Height on Pole ${format(maxPole)}.`);
    }
    if (position === "TOP_COMM" && topComm !== null && proposed < topComm) {
      issues.push("Proposed below top comm.");
    }
    if (position === "LOW_COMM" && lowComm !== null && proposed > lowComm) {
      issues.push("Proposed above low comm.");
    }
    // Proposed attachments have pole rules of their own. Keep those in the same
    // flagging field so the operator does not have to discover them in notes.
    const poleClearance = evaluateProposedPoleClearance(spanSide);
    if (!poleClearance.ok && poleClearance.message) issues.push(poleClearance.message);
    return {
      status: issues.length ? "PROBLEM" : "OK",
      message: issues.length ? Array.from(new Set(issues)).join(" ") : "OK"
    };
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

    if (midspan === null) {
      return {
        msProposed: "",
        finalMidspan: "",
        clearanceMSStatus: "MISSING",
        clearanceMSMessage: "Missing comm midspan.",
        clearanceMSIssue: false
      };
    }

    if (maxMS === null) {
      return {
        msProposed: format(midspan),
        finalMidspan: format(midspan),
        clearanceMSStatus: "MISSING_POWER",
        clearanceMSMessage: "Missing Max MS Comm / Low Power at midspan.",
        clearanceMSIssue: false
      };
    }

    // span.midspanMaxCommHeight already equals Low Power MS minus the required
    // power-to-comm clearance. The comm only has to be at or below this ceiling;
    // comm-to-comm spacing is validated separately in evaluateCommFlagging().
    const ok = midspan <= maxMS;
    return {
      msProposed: format(midspan),
      finalMidspan: format(midspan),
      clearanceMSStatus: ok ? "OK" : "PROBLEM",
      clearanceMSMessage: ok
        ? `${format(midspan)} <= ${format(maxMS)}.`
        : `${format(midspan)} exceeds limit ${format(maxMS)}.`,
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
          issues.push(`Comm-comm MS: ${format(diff)} with ${commOwnerLabel(other) || "no owner"}; minimum ${format(midspanClearance)}.`);
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

    if (midspan !== null && maxMS !== null && getSettingPosition() === "TOP_COMM" && span?.fromPole === sc.poleId) {
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
      const boltClearance = getPoleBoltBoltClearance();
      const ownExistingHeight = H().parseHeight(sc.existingHOA || "");
      const thisExisting = normalizedHeightLabelForCalc(sc.existingHOA);
      const thisEffective = normalizedHeightLabelForCalc(getEffectiveCommHOA(sc));
      if (sc.existingHOAChange && ownExistingHeight !== null) {
        const ownBoltDiff = Math.abs(poleHeight - ownExistingHeight);
        if (ownBoltDiff > 0 && ownBoltDiff < boltClearance) {
          issues.push(`Pole bolt-bolt: ${format(ownBoltDiff)} against Existing HOA ${format(ownExistingHeight)}; minimum ${format(boltClearance)}.`);
        }
      }
      S().getSpanCommsForPole(sc.poleId).forEach(other => {
        if (spanCommKey(other) === spanCommKey(sc)) return;
        const otherOwner = commOwnerLabel(other);
        const sameOwner = normalizeOwnerForMatch(otherOwner) && normalizeOwnerForMatch(otherOwner) === normalizeOwnerForMatch(owner);
        const otherEffective = getEffectiveCommHOA(other);
        const isSamePhysicalComm = sameOwner
          && normalizedHeightLabelForCalc(other.existingHOA) === thisExisting
          && normalizedHeightLabelForCalc(otherEffective) === thisEffective;
        if (isSamePhysicalComm) return;
        const otherHeight = H().parseHeight(getEffectiveCommHOA(other));
        if (otherHeight === null) return;
        const diff = Math.abs(poleHeight - otherHeight);
        const required = sameOwner ? boltClearance : poleClearance;
        const label = sameOwner ? "Pole bolt-bolt" : "Pole comm-comm";
        if (diff < required) {
          issues.push(`${label}: ${format(diff)} with ${otherOwner || "no owner"}; minimum ${format(required)}.`);
        }

        // Bolt-bolt also applies against previous attachment points. Example:
        // if Cable is moved to 22' and another comm used to sit at 21'10",
        // the new bolt is only 2" from that existing point, even if the other
        // comm was moved down to 21'. That must still flag.
        const otherExistingHeight = H().parseHeight(other.existingHOA || "");
        if (otherExistingHeight !== null) {
          const existingPointDiff = Math.abs(poleHeight - otherExistingHeight);
          if (existingPointDiff > 0 && existingPointDiff < boltClearance) {
            issues.push(`Pole bolt-bolt: ${format(existingPointDiff)} against Existing HOA ${format(otherExistingHeight)} from ${otherOwner || "no owner"}; minimum ${format(boltClearance)}.`);
          }
        }
      });
    }

    const maxPole = H().parseHeight(pole?.maxCommHeight || "");
    const changedPoleHeight = H().parseHeight(sc.existingHOAChange || "");
    if (changedPoleHeight !== null && maxPole !== null && changedPoleHeight > maxPole) {
      issues.push(`Pole: HOA Change ${format(changedPoleHeight)} exceeds max ${format(maxPole)}.`);
    }

    return {
      flaggingStatus: issues.length ? "PROBLEM" : "OK",
      flaggingMessage: issues.length ? Array.from(new Set(issues)).join(" ") : "OK"
    };
  }

  function evaluateProposedPoleClearance(spanSide) {
    const proposed = H().parseHeight(spanSide?.proposedHOA || "");
    if (proposed === null) return { ok: true, message: "" };
    const boltRequired = getPoleBoltBoltClearance();
    const commRequired = getPoleCommCommClearance();
    const issues = [];

    S().getSpanCommsForPole(spanSide.poleId)
      .map(sc => ({
        owner: commOwnerLabel(sc) || "sin owner",
        existing: H().parseHeight(sc.existingHOA || ""),
        effective: H().parseHeight(getEffectiveCommHOA(sc)),
        moved: Boolean(sc.existingHOAChange)
      }))
      .filter(item => item.effective !== null || item.existing !== null)
      .forEach(item => {
        if (item.effective !== null) {
          const diff = Math.abs(proposed - item.effective);
          if (diff === 0 && !item.moved) {
            issues.push(`Proposed ${format(proposed)} occupies the same HOA as ${item.owner}.`);
          }
          if (diff > 0 && diff < commRequired) {
            issues.push(`Proposed ${format(proposed)} does not respect Pole · Comm-comm ${format(commRequired)} against ${item.owner} ${format(item.effective)}.`);
          }
        }
        if (item.existing !== null) {
          const boltDiff = Math.abs(proposed - item.existing);
          if (boltDiff > 0 && boltDiff < boltRequired) {
            issues.push(`Proposed ${format(proposed)} does not respect Pole · Bolt-bolt ${format(boltRequired)} against Existing HOA ${format(item.existing)}.`);
          }
        }
      });

    S().getSpanSidesForPole(spanSide.poleId)
      .filter(otherSide => otherSide.spanId !== spanSide.spanId || otherSide.poleId !== spanSide.poleId)
      .map(otherSide => ({ spanId: otherSide.spanId, height: H().parseHeight(otherSide.proposedHOA || "") }))
      .filter(item => item.height !== null)
      .forEach(item => {
        const diff = Math.abs(proposed - item.height);
        if (diff > 0 && diff < boltRequired) {
          issues.push(`Proposed ${format(proposed)} does not respect Pole · Bolt-bolt ${format(boltRequired)} against another proposed ${format(item.height)}.`);
        }
      });

    if (!issues.length) return { ok: true, message: "" };
    return {
      ok: false,
      message: Array.from(new Set(issues)).join(" ")
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
      // The next pole proposed value does not always live on the same spanId.
      // This flow uses the first forward proposed value from the next pole; if
      // none exists, the field stays empty.
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
      // Documented comm formula:
      // New midspan = imported midspan + (local move / 2) + (remote move / 2).
      // Move = effective HOA - Existing HOA.
      // Effective HOA is HOA Change when present; otherwise Existing HOA.
      // Example: 20' -> 19' contributes -6" to midspan. If the other pole moves
      // from 20' -> 21', it contributes +6", and both changes offset each other.
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

  function cloneStateForAutoCalc() {
    return JSON.parse(JSON.stringify(S().getState()));
  }

  function autoCalcGroupKey(sc) {
    return [
      normalizeOwnerForMatch(commOwnerLabel(sc) || sc.owner),
      normalizedHeightLabelForCalc(sc.existingHOA)
    ].join("|");
  }

  function autoCalcGroupsForPole(poleId) {
    const groups = new Map();
    S().getSpanCommsForPole(poleId).forEach(sc => {
      const key = autoCalcGroupKey(sc);
      if (!groups.has(key)) {
        groups.set(key, {
          key,
          owner: commOwnerLabel(sc) || sc.owner,
          ownerToken: normalizeOwnerForMatch(commOwnerLabel(sc) || sc.owner),
          existingInches: H().parseHeight(sc.existingHOA),
          rows: []
        });
      }
      groups.get(key).rows.push(sc);
    });
    return Array.from(groups.values())
      .filter(group => group.existingInches !== null)
      .sort((a, b) => b.existingInches - a.existingInches);
  }

  function autoCalcMaxPole(poleId) {
    const pole = S().getPole(poleId);
    return H().parseHeight(pole?.maxCommHeight || "");
  }

  function autoCalcGapBetweenGroups(upper, lower) {
    if (upper.ownerToken && upper.ownerToken === lower.ownerToken) return getPoleBoltBoltClearance();
    return getPoleCommCommClearance();
  }

  function autoCalcPoleHasCommProblems(groups) {
    return groups.some(group => group.rows.some(row =>
      row.flaggingStatus === "PROBLEM" || row.clearanceMSStatus === "PROBLEM"
    ));
  }

  function autoCalcHasRealMidspan(spanId) {
    return S().getSpanCommsForSpan(spanId).some(row => getImportedMidspanInchesForComm(row) !== null)
      || S().getSpanSidesForSpan(spanId).some(side => parseMidspanValue(side.ocalcMS || side.proposedMidspan || side.msProposed || "") !== null);
  }

  function autoCalcProposedSpansForPole(poleId) {
    const seen = new Set();
    return S().getConnectedSpans(poleId)
      .filter(span => span.fromPole === poleId)
      .filter(span => autoCalcHasRealMidspan(span.spanId) || S().getSpanSide(span.spanId, poleId)?.isManualProposed)
      .filter(span => {
        const key = `${span.fromPole || ""}->${span.toPole || ""}`;
        if (seen.has(key)) return false;
        seen.add(key);
        return true;
      });
  }

  function autoCalcSetProposed(poleId, spans, proposedInches) {
    const proposed = format(proposedInches);
    spans.forEach(span => {
      const side = S().getSpanSide(span.spanId, poleId) || S().upsertSpanSide({ spanId: span.spanId, poleId });
      S().upsertSpanSide({
        ...side,
        proposedHOA: proposed,
        clearanceFixReadyAt: 0,
        pendingMidspanFinal: "",
        msProposed: "",
        finalMidspan: "",
        clearanceMSStatus: "",
        clearanceMSMessage: "",
        clearanceMSReason: "",
        clearanceMSIssue: false
      });
    });
  }

  function autoCalcBuildStackPlan(groups, topCommTarget) {
    if (topCommTarget === null || !groups.length) return [];
    const plan = [];
    groups.forEach((group, index) => {
      const previous = plan[index - 1];
      const stackTarget = previous ? previous.targetInches - autoCalcGapBetweenGroups(previous.group, group) : topCommTarget;
      const midspanTarget = autoCalcGroupMidspanTarget(group);
      const targetInches = Math.round(Math.min(group.existingInches, stackTarget, midspanTarget ?? Number.POSITIVE_INFINITY));
      plan.push({ group, targetInches });
    });
    return plan;
  }

  function autoCalcPlanMovesComm(plan) {
    return plan.some(item => item.group.rows.some(row => {
      const existing = H().parseHeight(row.existingHOA);
      return existing !== null && existing !== item.targetInches;
    }));
  }

  function autoCalcGroupMidspanTarget(group) {
    // A comm movement changes the related midspan by half of the pole movement.
    // This target is calculated from imported midspan + remote movement only,
    // ignoring the current local auto movement. That lets the second auto pass
    // reduce or remove a local move when the other pole already solved part of
    // the clearance.
    let target = null;
    group.rows.forEach(row => {
      const span = S().getSpan(row.spanId);
      const localExisting = H().parseHeight(row.existingHOA);
      const midspanSource = findMidspanSourceComm(row) || row;
      const importedMidspan = getImportedMidspanInchesForComm(midspanSource);
      const remote = findRemoteComm(row.spanId, row.poleId, row.ownerBase || row.owner, row.wireId || "");
      const remoteExisting = H().parseHeight(remote?.existingHOA || "");
      const remoteCurrent = H().parseHeight(remote ? getEffectiveCommHOA(remote) : "");
      const maxMS = H().parseHeight(span?.midspanMaxCommHeight || "");
      if (localExisting === null || importedMidspan === null || maxMS === null) return;

      let requiredMS = maxMS;
      if (getSettingPosition() === "TOP_COMM" && span?.fromPole === row.poleId) {
        const currentMS = getMidspanInchesForComm(row);
        const spanMidspans = S().getSpanCommsForSpan(row.spanId)
          .map(getMidspanInchesForComm)
          .filter(value => value !== null);
        const isTopCommAtMidspan = currentMS !== null && spanMidspans.length && currentMS === Math.max(...spanMidspans);
        if (isTopCommAtMidspan) requiredMS = Math.min(requiredMS, maxMS - getMidspanCommCommClearance());
      }

      const remoteAdjustment = remoteExisting !== null && remoteCurrent !== null ? (remoteCurrent - remoteExisting) / 2 : 0;
      const midspanWithoutLocalMove = importedMidspan + remoteAdjustment;
      if (midspanWithoutLocalMove <= requiredMS) return;
      const localTarget = localExisting + ((requiredMS - midspanWithoutLocalMove) * 2);
      target = target === null ? localTarget : Math.min(target, localTarget);
    });
    return target;
  }

  function autoCalcFinalProposedForPlan(proposedInches, plan, maxPole) {
    if (!plan.length) return proposedInches;
    const topCommAfterMove = plan[0].targetInches;
    return Math.min(proposedInches, maxPole, topCommAfterMove + getPoleCommCommClearance());
  }

  function autoCalcProposedCandidates(groups, maxPole, hasCommProblems, hasSpaceAbove, aboveProposed) {
    // Candidate order matters:
    // 1. If the existing stack is healthy and there is space above, try the
    //    cleanest attachment first: top comm + Pole · Comm-comm.
    // 2. Try the pole maximum so violation cases can use all available room.
    // 3. Try existing HOA heights exactly. Reusing the same bolt elevation is
    //    allowed only if the comm at that elevation is moved away.
    // 4. Try one bolt clearance below the maximum as a last simple fallback.
    const values = [];
    const seen = new Set();
    const add = value => {
      if (value === null || value === undefined || !Number.isFinite(value) || value < 0 || value > maxPole) return;
      const rounded = Math.round(value);
      if (seen.has(rounded)) return;
      seen.add(rounded);
      values.push(rounded);
    };

    if (hasSpaceAbove) add(aboveProposed);
    add(maxPole);
    groups
      .map(group => group.existingInches)
      .filter(value => value !== null && value <= maxPole)
      .sort((a, b) => b - a)
      .forEach(add);
    add(maxPole - getPoleBoltBoltClearance());
    return values;
  }

  function autoCalcAffectedPoleIds(poleId) {
    const ids = new Set([poleId]);
    S().getConnectedSpans(poleId).forEach(span => {
      if (span.fromPole) ids.add(span.fromPole);
      if (span.toPole) ids.add(span.toPole);
    });
    const wireIds = new Set(S().getSpanCommsForPole(poleId).map(row => row.wireId).filter(Boolean));
    Object.values(S().getState().spanComms || {}).forEach(row => {
      if (wireIds.has(row.wireId)) ids.add(row.poleId);
    });
    return Array.from(ids).filter(Boolean);
  }

  function autoCalcClearExistingMarks() {
    Object.values(S().getState().spanComms || {}).forEach(row => {
      if (row.autoCalcStatus === "AUTO" && !row.autoCalcMessage) return;
      if (!row.autoCalcStatus && !row.autoCalcMessage) return;
      S().upsertSpanComm({
        ...row,
        autoCalcStatus: row.autoCalcStatus === "AUTO" ? "AUTO" : "",
        autoCalcMessage: ""
      });
    });
  }

  function autoCalcBadCommKeys(state) {
    return new Set(Object.values(state?.spanComms || {})
      .filter(row => row.flaggingStatus === "PROBLEM" || row.clearanceMSStatus === "PROBLEM")
      .map(row => S().keyForSpanComm(row.spanId, row.poleId, row.owner, row.wireId || "")));
  }

  function autoCalcBadSpanSideKeys(state) {
    return new Set(Object.values(state?.spanSides || {})
      .filter(side => side.proposedHOA && (side.proposedFlaggingStatus === "PROBLEM" || side.clearanceMSStatus === "PROBLEM"))
      .map(side => S().keyForSpanSide(side.spanId, side.poleId)));
  }

  function autoCalcResultIsSafe(affectedPoleIds, beforeState = null, primaryPoleId = "") {
    const beforeBadComms = autoCalcBadCommKeys(beforeState);
    const beforeBadSides = autoCalcBadSpanSideKeys(beforeState);
    const affected = new Set(affectedPoleIds);
    const rows = Object.values(S().getState().spanComms || {}).filter(row => affected.has(row.poleId));
    const badComm = rows.find(row => {
      if (row.flaggingStatus !== "PROBLEM" && row.clearanceMSStatus !== "PROBLEM") return false;
      const key = S().keyForSpanComm(row.spanId, row.poleId, row.owner, row.wireId || "");
      return row.poleId === primaryPoleId || !beforeBadComms.has(key);
    });
    if (badComm) return { ok: false, reason: badComm.flaggingMessage || badComm.clearanceMSMessage || "Comm clearance issue." };

    const badProposed = Object.values(S().getState().spanSides || {}).find(side =>
      affected.has(side.poleId)
      && side.proposedHOA
      && (side.proposedFlaggingStatus === "PROBLEM" || side.clearanceMSStatus === "PROBLEM")
      && (side.poleId === primaryPoleId || !beforeBadSides.has(S().keyForSpanSide(side.spanId, side.poleId)))
    );
    if (badProposed) return { ok: false, reason: badProposed.proposedFlaggingMessage || badProposed.clearanceMSMessage || "Proposed clearance issue." };
    return { ok: true, reason: "" };
  }

  function autoCalcApplyPlan(plan) {
    plan.forEach(item => {
      const target = format(item.targetInches);
      item.group.rows.forEach(row => {
        const existing = H().parseHeight(row.existingHOA);
        S().upsertSpanComm({
          ...row,
          existingHOAChange: existing === item.targetInches ? "" : target,
          autoCalcStatus: existing === item.targetInches ? "" : "AUTO",
          autoCalcMessage: ""
        });
      });
    });
  }

  function autoCalcTryCandidate(beforePole, poleId, proposedInches, requireMoveBecauseOfProblems = false, hasManualChanges = false) {
    // Each candidate is tested against a clean copy of the pole state. That
    // lets the solver try a proposed height, calculate any required comm stack,
    // run the normal validations, and throw that attempt away if it fails.
    S().setState(beforePole);
    const groups = autoCalcGroupsForPole(poleId);
    const proposedSpans = autoCalcProposedSpansForPole(poleId);
    if (!proposedSpans.length) return { ok: false, reason: "No proposed spans." };

    autoCalcSetProposed(poleId, proposedSpans, proposedInches);
    const topExisting = groups.length ? Math.max(...groups.map(group => group.existingInches)) : null;
    const topCommTarget = groups.length ? proposedInches - getPoleCommCommClearance() : null;
    const hasAutoChanges = groups.some(group => group.rows.some(row => row.autoCalcStatus === "AUTO"));
    const mustMoveComms = groups.length && (requireMoveBecauseOfProblems || hasAutoChanges || (topExisting !== null && topExisting > topCommTarget));
    const plan = mustMoveComms ? autoCalcBuildStackPlan(groups, topCommTarget) : [];

    if (!plan || plan.some(item => item.targetInches < 0)) return { ok: false, reason: "No safe pole stack found." };
    const maxPole = autoCalcMaxPole(poleId);
    const finalProposed = maxPole === null ? proposedInches : autoCalcFinalProposedForPlan(proposedInches, plan, maxPole);
    // Do not overwrite user-entered HOA changes. A candidate can still pass if
    // it only fills Proposed and the existing manual comm moves already work.
    if (hasManualChanges && autoCalcPlanMovesComm(plan)) return { ok: false, reason: "Existing HOA Change already set." };
    autoCalcSetProposed(poleId, proposedSpans, finalProposed);
    autoCalcApplyPlan(plan);
    const affected = autoCalcAffectedPoleIds(poleId);
    affected.forEach(recalculateSpansForPole);
    const safe = autoCalcResultIsSafe(affected, beforePole, poleId);
    return safe.ok ? { ok: true, reason: "" } : safe;
  }

  function autoCalcMovementSignature() {
    const state = S().getState();
    const commMoves = Object.values(state.spanComms || {})
      .map(row => ({
        key: S().keyForSpanComm(row.spanId, row.poleId, row.owner, row.wireId || ""),
        existingHOAChange: row.existingHOAChange || "",
        autoCalcStatus: row.autoCalcStatus || ""
      }))
      .sort((a, b) => a.key.localeCompare(b.key));
    const proposedMoves = Object.values(state.spanSides || {})
      .map(side => ({
        key: S().keyForSpanSide(side.spanId, side.poleId),
        proposedHOA: side.proposedHOA || "",
        nextPoleProposed: side.proposedHOAChange || "",
        nextPoleProposedAuto: Boolean(side.nextPoleProposedAuto)
      }))
      .sort((a, b) => a.key.localeCompare(b.key));

    return JSON.stringify({ commMoves, proposedMoves });
  }

  function autoCalculateMovements() {
    const summary = { applied: 0, manual: 0, skipped: 0, passes: 0, converged: false, stoppedByRepeat: false, maxPassesReached: false };
    if (getSettingPosition() !== "TOP_COMM") {
      return { ...summary, disabled: true };
    }
    // Old auto/manual labels were only helper metadata. Clear them every run so
    // the UI only shows real clearance flagging, not solver bookkeeping.
    autoCalcClearExistingMarks();
    const poleIds = Object.keys(S().getState().poles || {});
    const appliedPoles = new Set();
    const manualPoles = new Set();
    const skippedPoles = new Set();
    const seenSignatures = new Set();
    const maxPasses = Math.max(8, poleIds.length * 2 + 4);

    let previousSignature = autoCalcMovementSignature();
    seenSignatures.add(previousSignature);

    for (let pass = 0; pass < maxPasses; pass += 1) {
      recalculateAll();
      poleIds.forEach(poleId => {
        const beforePole = cloneStateForAutoCalc();
        const groups = autoCalcGroupsForPole(poleId);
        const proposedSpans = autoCalcProposedSpansForPole(poleId);
        const maxPole = autoCalcMaxPole(poleId);
        if (!proposedSpans.length) {
          skippedPoles.add(poleId);
          return;
        }
        if (maxPole === null) {
          manualPoles.add(poleId);
          return;
        }

        const hasManualChanges = groups.some(group => group.rows.some(row => row.existingHOAChange && row.autoCalcStatus !== "AUTO"));
        const hasCommProblems = autoCalcPoleHasCommProblems(groups);
        const topExisting = groups.length ? Math.max(...groups.map(group => group.existingInches)) : null;
        const aboveProposed = topExisting === null ? maxPole : topExisting + getPoleCommCommClearance();
        const hasSpaceAbove = aboveProposed <= maxPole;

        const candidates = autoCalcProposedCandidates(groups, maxPole, hasCommProblems, hasSpaceAbove, aboveProposed);
        const solved = candidates.some(candidate => autoCalcTryCandidate(beforePole, poleId, candidate, hasCommProblems, hasManualChanges).ok);
        if (solved) {
          appliedPoles.add(poleId);
          manualPoles.delete(poleId);
          return;
        }

        S().setState(beforePole);
        recalculateSpansForPole(poleId);
        manualPoles.add(poleId);
      });

      recalculateAll();
      summary.passes = pass + 1;
      const nextSignature = autoCalcMovementSignature();
      if (nextSignature === previousSignature) {
        summary.converged = true;
        break;
      }
      if (seenSignatures.has(nextSignature)) {
        summary.stoppedByRepeat = true;
        break;
      }
      seenSignatures.add(nextSignature);
      previousSignature = nextSignature;
    }

    recalculateAll();
    if (!summary.converged && !summary.stoppedByRepeat && summary.passes >= maxPasses) {
      summary.maxPassesReached = true;
    }
    summary.applied = appliedPoles.size;
    summary.manual = manualPoles.size;
    summary.skipped = skippedPoles.size;
    return summary;
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
    const next = { ...sc, [field]: value || "" };
    if (field === "existingHOAChange") {
      next.autoCalcStatus = "";
      next.autoCalcMessage = "";
    }
    S().upsertSpanComm(next);
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

    // Also recalculate reciprocal span comms by Wire Id. This covers imported
    // Fore/Back cases with different Span Ids that belong to the same cable.
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
    evaluateProposedPoleClearance,
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
    getReferenceMidspansForSpanSide,
    autoCalculateMovements
  };
})(window);
