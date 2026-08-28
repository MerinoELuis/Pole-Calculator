(function (global) {
  "use strict";

  const STATUS = Object.freeze({
    SAFE: "SAFE",
    BEST_AVAILABLE: "BEST_AVAILABLE",
    CRITICAL: "CRITICAL",
    MANUAL: "MANUAL",
    SKIPPED: "SKIPPED"
  });
  const AUTO = "AUTO";
  const MAX_PASSES = 3;
  const MAX_CANDIDATES = 72;
  const CANDIDATE_YIELD_INTERVAL = 6;
  let lastDebugTrace = null;

  const S = () => global.AppStore;
  const C = () => global.Calculations;
  const H = () => global.HeightUtils;
  const clone = value => JSON.parse(JSON.stringify(value));
  const parse = value => H()?.parseHeight?.(value) ?? null;
  const format = value => H()?.formatHeight?.(Math.round(value)) || "";
  const text = value => String(value ?? "").trim();

  function planSnapshot(poleId) {
    return {
      proposed: (S()?.getSpanSidesForPole?.(poleId) || [])
        .filter(side => text(side.proposedHOA))
        .map(side => ({
          spanId: side.spanId,
          proposedHOA: side.proposedHOA,
          source: side.autoCalcProposedStatus || "MANUAL"
        })),
      comms: (S()?.getSpanCommsForPole?.(poleId) || [])
        .filter(row => text(row.existingHOAChange))
        .map(row => ({
          spanId: row.spanId,
          wireId: row.wireId || "",
          owner: owner(row),
          existingHOA: row.existingHOA || "",
          hoaChange: row.existingHOAChange,
          source: row.autoCalcStatus || "MANUAL"
        }))
    };
  }

  function analysisSnapshot(analysis) {
    if (!analysis) return null;
    return {
      status: statusForAnalysis(analysis),
      poleViolations: {
        count: analysis.poleViolationCount,
        totalShortfall: format(analysis.poleViolationInches),
        messages: [...(analysis.poleMessages || [])]
      },
      midspanViolations: {
        count: analysis.midspanViolationCount,
        totalShortfall: format(analysis.midspanViolationInches),
        messages: [...(analysis.midspanMessages || [])]
      },
      movement: {
        movedCommGroups: analysis.movedCommCount,
        totalMovement: format(analysis.totalMovementInches),
        idealProposed: Number.isFinite(analysis.idealProposedInches) ? format(analysis.idealProposedInches) : "",
        selectedProposed: Number.isFinite(analysis.selectedProposedInches) ? format(analysis.selectedProposedInches) : ""
      }
    };
  }

  function groupSnapshot(group) {
    return {
      groupKey: group.key,
      owner: group.owner,
      existingHOA: format(group.existingInches),
      effectiveBeforeCandidate: format(group.effectiveInches),
      minimumHOA: Number.isFinite(group.minimumInches) ? format(group.minimumInches) : "",
      maximumHOA: Number.isFinite(group.maximumInches) ? format(group.maximumInches) : "",
      locked: Boolean(group.locked),
      lockedHOA: Number.isFinite(group.lockedInches) ? format(group.lockedInches) : ""
    };
  }

  function beginPoleTrace(poleId, mode, pass) {
    if (!lastDebugTrace) return null;
    const trace = {
      poleId,
      pass,
      mode,
      startedAt: new Date().toISOString(),
      automaticPlanBeforeRetry: planSnapshot(poleId),
      candidates: []
    };
    lastDebugTrace.poleAttempts.push(trace);
    return trace;
  }

  function finishPoleTrace(trace, result, poleId) {
    if (!trace) return;
    trace.completedAt = new Date().toISOString();
    trace.result = result ? clone(result) : null;
    trace.selectedPlan = planSnapshot(poleId);
  }

  // Candidate evaluation is CPU-heavy. Yielding between small batches lets the
  // browser paint progress and process its own events instead of reporting that
  // the page has stopped responding on large jobs.
  function yieldToBrowser() {
    return new Promise(resolve => {
      if (typeof global.setTimeout === "function") global.setTimeout(resolve, 0);
      else Promise.resolve().then(resolve);
    });
  }

  function reportProgress(callback, detail) {
    if (typeof callback !== "function") return;
    try {
      callback(detail);
    } catch (error) {
      global.console?.warn?.("Auto Calculate progress callback failed.", error);
    }
  }

  function modeFromState(state = S()?.getState?.()) {
    return String(state?.settings?.position || "TOP_COMM").toUpperCase() === "LOW_COMM"
      ? "LOW_COMM"
      : "TOP_COMM";
  }

  function commClearance(state = S()?.getState?.()) {
    return parse(state?.settings?.commClearance || "12\"") ?? 12;
  }

  function boltClearance(state = S()?.getState?.()) {
    return parse(state?.settings?.boltClearance || "4\"") ?? 4;
  }

  function normalizeOwner(value) {
    const normalized = text(value)
      .replace(/^COMMUNICATION\s*>\s*/i, "")
      .replace(/,\s*.*$/, "")
      .toLowerCase()
      .normalize("NFD")
      .replace(/[\u0300-\u036f]/g, "")
      .replace(/[^a-z0-9]+/g, " ")
      .trim();
    if (/century\s*link|centurylink|\bctl\b|telco/.test(normalized)) return "ctl";
    if (/cable\s*one|cox|catv/.test(normalized)) return "catv";
    if (/vexus/.test(normalized)) return "vexus";
    if (/wecom/.test(normalized)) return "wecom";
    if (/fiber/.test(normalized)) return "fiber";
    return normalized;
  }

  function owner(row) {
    return C()?.commOwnerLabel?.(row) || row?.ownerBase || row?.owner || "COMM";
  }

  function commKey(row) {
    return S()?.keyForSpanComm?.(row.spanId, row.poleId, row.owner, row.wireId || "")
      || [row.spanId, row.poleId, row.owner, row.wireId || ""].join("__");
  }

  function effective(row) {
    return parse(C()?.getEffectiveCommHOA?.(row) || row?.existingHOA || "");
  }

  function manualComm(row) {
    const active = C()?.isCommMovementsActive ? C().isCommMovementsActive(row?.poleId) : true;
    return Boolean(active && row?.existingHOAChange && row?.autoCalcStatus !== AUTO);
  }

  // A local HOA movement changes its span midspan by half that movement.
  // Convert current Environment/Power MS shortfalls into pole-height bounds so
  // TOP COMM can raise a comm when raising is the safe correction.
  function midspanBoundsForGroup(group) {
    const rows = group.rowKeys
      .map(key => S()?.getState?.()?.spanComms?.[key])
      .filter(Boolean);
    let minimumInches = null;
    let maximumInches = null;
    rows.forEach(row => {
      const span = S()?.getSpan?.(row.spanId);
      const midspan = parse(C()?.displayMidspanForComm?.(row) || row.calculatedMidspan || row.midspan || "");
      const current = effective(row);
      const flaggingMessage = text(row.flaggingMessage);
      if (!span || midspan === null || current === null) return;
      const environmentMinimum = parse(span.environmentClearance || "");
      if (/Environment:/i.test(flaggingMessage) && environmentMinimum !== null && midspan < environmentMinimum) {
        const required = Math.round(current + (environmentMinimum - midspan) * 2);
        minimumInches = minimumInches === null ? required : Math.max(minimumInches, required);
      }
      const powerMaximum = parse(span.midspanMaxCommHeight || "");
      if (/Power MS:/i.test(flaggingMessage) && powerMaximum !== null && midspan > powerMaximum) {
        const required = Math.round(current - (midspan - powerMaximum) * 2);
        maximumInches = maximumInches === null ? required : Math.min(maximumInches, required);
      }
    });
    return { minimumInches, maximumInches };
  }

  function manualProposed(side) {
    return Boolean(side?.proposedHOA && side?.autoCalcProposedStatus !== AUTO);
  }

  function groupsForPole(poleId) {
    const spanComms = S()?.getState?.()?.spanComms || {};
    const groups = new Map();
    (S()?.getSpanCommsForPole?.(poleId) || [])
      .filter(row => !C()?.isPofComm?.(row))
      .forEach(row => {
        const existing = parse(row.existingHOA || "");
        if (existing === null) return;
        const key = `${normalizeOwner(owner(row))}|${existing}`;
        if (!groups.has(key)) {
          groups.set(key, {
            key,
            owner: owner(row),
            ownerToken: normalizeOwner(owner(row)),
            existingInches: existing,
            effectiveInches: effective(row) ?? existing,
            locked: false,
            lockedInches: null,
            rowKeys: []
          });
        }
        const group = groups.get(key);
        group.rowKeys.push(commKey(row));
        group.effectiveInches = effective(row) ?? group.effectiveInches;
        if (manualComm(row)) {
          group.locked = true;
          group.lockedInches = effective(row);
        }
      });
    return Array.from(groups.values())
      .map(group => ({
        ...group,
        serviceDrop: group.rowKeys.length > 0 && group.rowKeys.every(key => Boolean(spanComms[key]?.serviceDrop)),
        ...midspanBoundsForGroup(group)
      }))
      .sort((a, b) => b.existingInches - a.existingInches);
  }

  function gap(upper, lower, state) {
    if (upper?.ownerToken && upper.ownerToken === lower?.ownerToken && (upper.serviceDrop || lower.serviceDrop)) return 0;
    return upper?.ownerToken && upper.ownerToken === lower?.ownerToken
      ? boltClearance(state)
      : commClearance(state);
  }

  function eligibleProposedSpans(poleId) {
    const state = S()?.getState?.() || {};
    const allowNoMidspan = state.settings?.proposeForeSpanWithoutMidspan === true;
    const seen = new Set();
    return (S()?.getConnectedSpans?.(poleId) || [])
      .filter(span => {
        const side = S()?.getSpanSide?.(span.spanId, poleId);
        const type = String(span?.type || span?.rawType || "").toLowerCase();
        const forward = C()?.isSpanEligibleForProposed
          ? C().isSpanEligibleForProposed(span, poleId)
          : /fore\s*span|forespan/.test(type) && span.fromPole === poleId;
        return forward || side?.isManualProposed;
      })
      .filter(span => !S()?.getSpanSide?.(span.spanId, poleId)?.isProposedExcluded)
      .filter(span => {
        const side = S()?.getSpanSide?.(span.spanId, poleId);
        return allowNoMidspan || C()?.spanHasRealMidspan?.(span.spanId) || side?.isManualProposed;
      })
      .filter(span => !S()?.getSpanSide?.(span.spanId, poleId)?.isAdditionalProposed)
      .filter(span => {
        const key = `${span.fromPole || ""}->${span.toPole || ""}`;
        if (seen.has(key)) return false;
        seen.add(key);
        return true;
      });
  }

  function idealProposedHeight(groups, mode, state = S()?.getState?.()) {
    const values = groups
      .map(group => group.locked && group.lockedInches !== null
        ? group.lockedInches
        : group.existingInches)
      .filter(Number.isFinite);
    if (!values.length) return null;
    return mode === "LOW_COMM"
      ? Math.min(...values) - commClearance(state)
      : Math.max(...values) + commClearance(state);
  }

  // A lower comm's Midspan correction can force every comm above it upward.
  // Propagate those minimums before selecting Proposed so the candidate list
  // reserves enough vertical room for the complete comm stack.
  function minimumTopCommHeight(groups, state = S()?.getState?.()) {
    const ordered = [...groups].sort((a, b) => b.existingInches - a.existingInches);
    if (!ordered.length) return null;
    const hasRequiredFloor = ordered.some(group =>
      (group.locked && group.lockedInches !== null)
      || Number.isFinite(group.minimumInches)
    );
    if (!hasRequiredFloor) return null;
    const floors = ordered.map(group => {
      if (group.locked && group.lockedInches !== null) return group.lockedInches;
      return Number.isFinite(group.minimumInches) ? group.minimumInches : 0;
    });
    for (let index = floors.length - 2; index >= 0; index -= 1) {
      floors[index] = Math.max(
        floors[index],
        floors[index + 1] + gap(ordered[index], ordered[index + 1], state)
      );
    }
    return floors[0];
  }

  function addCandidate(set, value, maxPole) {
    if (!Number.isFinite(value)) return;
    const rounded = Math.round(value);
    if (rounded >= 0 && rounded <= maxPole) set.add(rounded);
  }

  function candidateHeights({ groups, maxPole, mode, currentProposed = [], manualReference = null, preferMaximum = false, state = S()?.getState?.() }) {
    if (!Number.isFinite(maxPole) || maxPole < 0) return [];
    if (manualReference !== null) return [Math.round(manualReference)];
    const values = new Set();
    const requiredValues = new Set();
    const progressive = [];
    const progressiveSeen = new Set();
    const ideal = idealProposedHeight(groups, mode, state);
    const comm = commClearance(state);
    const bolt = boltClearance(state);
    const anchors = [ideal, maxPole, maxPole - bolt, ...currentProposed];
    const addProgressive = value => {
      if (!Number.isFinite(value)) return;
      const rounded = Math.round(value);
      if (rounded < 0 || rounded > maxPole || progressiveSeen.has(rounded)) return;
      progressiveSeen.add(rounded);
      progressive.push(rounded);
    };
    const minimumTop = mode === "TOP_COMM" ? minimumTopCommHeight(groups, state) : null;
    const highestLegalProposed = mode === "TOP_COMM"
      ? highestLegalTarget(
        maxPole,
        0,
        groups
          .filter(group => Number.isFinite(group.existingInches))
          .map(group => ({ height: group.existingInches, serviceDrop: Boolean(group.serviceDrop) })),
        state
      )
      : null;
    if (minimumTop !== null) {
      const requiredProposed = Math.round(minimumTop + comm);
      anchors.push(requiredProposed);
      if (requiredProposed >= 0 && requiredProposed <= maxPole) requiredValues.add(requiredProposed);
    }
    groups.forEach(group => {
      const existing = group.existingInches;
      const current = group.effectiveInches ?? existing;
      anchors.push(
        existing,
        current,
        existing + comm,
        existing - comm,
        existing + bolt,
        existing - bolt,
        group.minimumInches,
        group.maximumInches,
        group.minimumInches === null ? null : group.minimumInches + comm,
        group.maximumInches === null ? null : group.maximumInches + comm
      );
    });
    anchors
      .filter(Number.isFinite)
      .forEach(value => [-1, 0, 1].forEach(offset => addCandidate(values, value + offset, maxPole)));
    if (ideal !== null) {
      for (let offset = -18; offset <= 18; offset += 1) addCandidate(values, ideal + offset, maxPole);
      const start = mode === "LOW_COMM" ? ideal - 48 : Math.min(ideal, maxPole) - 60;
      const end = mode === "LOW_COMM" ? ideal + 48 : maxPole;
      for (let value = start; value <= end; value += 6) addCandidate(values, value, maxPole);
    }
    for (let offset = 0; offset <= 12; offset += 1) addCandidate(values, maxPole - offset, maxPole);
    const preferred = ideal ?? maxPole;
    if (mode === "TOP_COMM") {
      if (preferMaximum) {
        addProgressive(maxPole);
        addProgressive(highestLegalProposed);
      }
      addProgressive(ideal);
      if (!preferMaximum && (ideal === null || ideal > maxPole)) {
        addProgressive(maxPole);
        addProgressive(highestLegalProposed);
      }
      Array.from(requiredValues).forEach(addProgressive);
      addProgressive(maxPole);
      addProgressive(highestLegalProposed);
    }
    const orderedCandidates = [
      ...progressive,
      ...Array.from(values)
        .filter(value => !progressiveSeen.has(value))
        .sort((a, b) => Math.abs(a - preferred) - Math.abs(b - preferred) || (mode === "LOW_COMM" ? a - b : b - a))
    ]
      .slice(0, MAX_CANDIDATES);
    Object.defineProperty(orderedCandidates, "progressiveCount", {
      value: Math.min(progressive.length, orderedCandidates.length),
      enumerable: false
    });
    return orderedCandidates;
  }

  function isLegalExistingBoltTarget(target, existingBoltPoints, state, movingServiceDrop = false) {
    if (movingServiceDrop) return true;
    const requiredClearance = boltClearance(state);
    return existingBoltPoints.every(point => {
      const existing = typeof point === "number" ? point : point?.height;
      if (!Number.isFinite(existing) || point?.serviceDrop) return true;
      const difference = Math.abs(target - existing);
      return difference === 0 || difference >= requiredClearance;
    });
  }

  function highestLegalTarget(ceiling, floor, existingBoltPoints, state, movingServiceDrop = false) {
    const roundedCeiling = Math.max(0, Math.floor(ceiling));
    const roundedFloor = Math.max(0, Math.ceil(floor));
    for (let target = roundedCeiling; target >= roundedFloor; target -= 1) {
      if (isLegalExistingBoltTarget(target, existingBoltPoints, state, movingServiceDrop)) return target;
    }
    for (let target = Math.min(roundedCeiling, roundedFloor - 1); target >= 0; target -= 1) {
      if (isLegalExistingBoltTarget(target, existingBoltPoints, state, movingServiceDrop)) return target;
    }
    return roundedCeiling;
  }

  function buildStackPlan(groups, proposedInches, mode, maxPole, state = S()?.getState?.(), options = {}) {
    const ordered = mode === "LOW_COMM"
      ? [...groups].sort((a, b) => a.existingInches - b.existingInches)
      : [...groups].sort((a, b) => b.existingInches - a.existingInches);
    const existingBoltPoints = ordered
      .filter(group => Number.isFinite(group.existingInches))
      .map(group => ({ height: group.existingInches, serviceDrop: Boolean(group.serviceDrop) }));
    const topCommFloors = mode === "TOP_COMM"
      ? ordered.map(group => {
        if (group.locked && group.lockedInches !== null) return group.lockedInches;
        return Number.isFinite(group.minimumInches) ? group.minimumInches : 0;
      })
      : [];
    for (let index = topCommFloors.length - 2; index >= 0; index -= 1) {
      topCommFloors[index] = Math.max(
        topCommFloors[index],
        topCommFloors[index + 1] + gap(ordered[index], ordered[index + 1], state)
      );
    }
    const plan = [];
    ordered.forEach((group, index) => {
      const previous = plan[index - 1];
      let target;
      if (group.locked && group.lockedInches !== null) {
        target = group.lockedInches;
      } else if (mode === "LOW_COMM") {
        const minimum = previous
          ? previous.targetInches + gap(previous.group, group, state)
          : proposedInches + commClearance(state);
        let preferred = group.existingInches;
        if (Number.isFinite(group.minimumInches)) preferred = Math.max(preferred, group.minimumInches);
        if (Number.isFinite(group.maximumInches)) preferred = Math.min(preferred, group.maximumInches);
        target = Math.min(maxPole, Math.max(preferred, minimum));
      } else {
        const maximum = previous
          ? previous.targetInches - gap(previous.group, group, state)
          : proposedInches - commClearance(state);
        const ceiling = Number.isFinite(group.maximumInches)
          ? Math.min(maximum, group.maximumInches, maxPole)
          : Math.min(maximum, maxPole);
        const floor = topCommFloors[index] || 0;
        let preferred = group.existingInches;
        if (options.preferHighest) {
          target = highestLegalTarget(ceiling, floor, existingBoltPoints, state, group.serviceDrop);
        } else if (ceiling >= floor) {
          target = Math.max(floor, Math.min(preferred, ceiling));
        } else {
          // This Proposed candidate cannot fit the required stack. Keep the
          // candidate inside its upper envelope and let validation reject it.
          target = Math.max(0, ceiling);
        }
      }
      plan.push({ group, targetInches: Math.round(target) });
    });
    return plan;
  }

  function manualProposedReference(spans, poleId, mode) {
    const values = spans
      .map(span => S()?.getSpanSide?.(span.spanId, poleId))
      .filter(manualProposed)
      .map(side => parse(side.proposedHOA))
      .filter(value => value !== null);
    if (!values.length) return null;
    return mode === "LOW_COMM" ? Math.max(...values) : Math.min(...values);
  }

  function currentProposed(spans, poleId) {
    return spans.map(span => parse(S()?.getSpanSide?.(span.spanId, poleId)?.proposedHOA || "")).filter(value => value !== null);
  }

  function affectedPoles(poleId) {
    const ids = new Set([poleId]);
    (S()?.getConnectedSpans?.(poleId) || []).forEach(span => {
      if (span.fromPole) ids.add(span.fromPole);
      if (span.toPole) ids.add(span.toPole);
    });
    const wireIds = new Set((S()?.getSpanCommsForPole?.(poleId) || []).map(row => row.wireId).filter(Boolean));
    Object.values(S()?.getState?.()?.spanComms || {}).forEach(row => {
      if (wireIds.has(row.wireId)) ids.add(row.poleId);
    });
    return Array.from(ids).filter(Boolean);
  }

  function recalculateAffected(poleId) {
    const ids = affectedPoles(poleId);
    if (typeof C()?.recalculateSpansForPole === "function") ids.forEach(id => C().recalculateSpansForPole(id));
    else C()?.recalculateAll?.();
    return ids;
  }

  function applyPlan(plan) {
    const state = S()?.getState?.();
    plan.forEach(item => item.group.rowKeys.forEach(key => {
      const row = state?.spanComms?.[key];
      if (!row || manualComm(row)) return;
      const moved = parse(row.existingHOA || "") !== item.targetInches;
      S()?.upsertSpanComm?.({
        ...row,
        existingHOAChange: moved ? format(item.targetInches) : "",
        autoCalcStatus: moved ? AUTO : "",
        autoCalcMessage: ""
      });
    }));
  }

  function applyProposed(spans, poleId, proposedInches, mode) {
    const proposed = format(proposedInches);
    spans.forEach(span => {
      const side = S()?.getSpanSide?.(span.spanId, poleId) || S()?.upsertSpanSide?.({ spanId: span.spanId, poleId });
      if (!side || manualProposed(side)) return;
      S()?.upsertSpanSide?.({
        ...side,
        proposedHOA: proposed,
        autoCalcProposedStatus: AUTO,
        autoCalcProposedMode: mode,
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

  function splitIssueMessage(message) {
    const source = text(message);
    if (!source || source === "OK") return [];
    return source.split(/\.\s+/).map(value => value.trim()).filter(Boolean).map(value => value.endsWith(".") ? value : `${value}.`);
  }

  function isMidspanIssue(message) {
    return /^(Environment:|Power MS:|Comm-comm MS|Midspan order|No room for proposed above top comm)/i.test(text(message));
  }

  function issueSeverity(message) {
    const values = (text(message).match(/[-−]?\d+\s*'(?:\s*\d+\s*")?|[-−]?\d+\s*"/g) || [])
      .map(value => parse(value.replace("−", "-")))
      .filter(value => value !== null);
    if (!values.length) return 1;
    if (/does not respect/i.test(message) && values.length >= 3) return Math.max(1, values[1] - Math.abs(values[0] - values[2]));
    if (/separation/i.test(message) && /minimum/i.test(message) && values.length >= 2) return Math.max(1, values.at(-1) - values.at(-2));
    if (/(exceeds|<|>)/.test(message) && values.length >= 2) return Math.max(1, Math.abs(values[0] - values[1]));
    return 1;
  }

  function physicalSpanPair(span) {
    return [span?.fromPole, span?.toPole].filter(Boolean).sort().join("|");
  }

  function collectIssues(poleId) {
    const state = S()?.getState?.() || {};
    const poleIssues = new Map();
    const midspanIssues = new Map();
    const add = (target, message) => {
      const clean = text(message);
      if (clean && clean !== "OK") target.set(clean.toLowerCase().replace(/\s+/g, " "), clean);
    };
    const pairs = new Set((S()?.getConnectedSpans?.(poleId) || []).map(physicalSpanPair).filter(Boolean));
    const spanIds = new Set(Object.values(state.spans || {}).filter(span => pairs.has(physicalSpanPair(span))).map(span => span.spanId));
    const rows = Object.values(state.spanComms || {}).filter(row => row.poleId === poleId || spanIds.has(row.spanId));
    rows.forEach(row => {
      if (row.flaggingStatus === "PROBLEM") {
        splitIssueMessage(row.flaggingMessage).forEach(message => {
          if (isMidspanIssue(message)) add(midspanIssues, message);
          else if (row.poleId === poleId) add(poleIssues, message);
        });
      }
      if (row.clearanceMSStatus === "PROBLEM" && !/Power MS:/i.test(row.flaggingMessage || "")) {
        add(midspanIssues, row.clearanceMSMessage || "Midspan clearance issue.");
      }
    });
    (S()?.getSpanSidesForPole?.(poleId) || []).forEach(side => {
      if (!side.proposedHOA) return;
      if (side.proposedFlaggingStatus === "PROBLEM") splitIssueMessage(side.proposedFlaggingMessage).forEach(message => add(poleIssues, message));
      if (side.clearanceMSStatus === "PROBLEM") add(midspanIssues, side.clearanceMSMessage || "Proposed midspan clearance issue.");
    });
    const poleMessages = Array.from(poleIssues.values());
    const midspanMessages = Array.from(midspanIssues.values());
    return {
      poleMessages,
      midspanMessages,
      poleViolationCount: poleMessages.length,
      poleViolationInches: poleMessages.reduce((sum, message) => sum + issueSeverity(message), 0),
      midspanViolationCount: midspanMessages.length,
      midspanViolationInches: midspanMessages.reduce((sum, message) => sum + issueSeverity(message), 0)
    };
  }

  function movementMetrics(poleId, mode) {
    const groups = groupsForPole(poleId);
    let movedCommCount = 0;
    let totalMovementInches = 0;
    groups.forEach(group => {
      const movement = Math.abs((group.effectiveInches ?? group.existingInches) - group.existingInches);
      if (movement) {
        movedCommCount += 1;
        totalMovementInches += movement;
      }
    });
    const ideal = idealProposedHeight(groups, mode);
    const proposals = (S()?.getSpanSidesForPole?.(poleId) || []).map(side => parse(side.proposedHOA || "")).filter(value => value !== null);
    return {
      movedCommCount,
      totalMovementInches,
      idealProposedInches: ideal,
      selectedProposedInches: proposals.length
        ? (mode === "LOW_COMM" ? Math.min(...proposals) : Math.max(...proposals))
        : null,
      proposedDistanceFromIdeal: ideal === null || !proposals.length
        ? Number.MAX_SAFE_INTEGER
        : proposals.reduce((sum, value) => sum + Math.abs(value - ideal), 0)
    };
  }

  function categoryForAnalysis(value) {
    if (!value.poleViolationCount && !value.midspanViolationCount) return 0;
    if (!value.poleViolationCount) return 1;
    return 2;
  }

  function rankAnalysis(value) {
    return [
      categoryForAnalysis(value),
      value.poleViolationCount,
      value.poleViolationInches,
      value.midspanViolationCount,
      value.midspanViolationInches,
      value.movedCommCount,
      value.totalMovementInches,
      value.proposedDistanceFromIdeal
    ];
  }

  function compareAnalyses(left, right) {
    if (!left) return 1;
    if (!right) return -1;
    const a = rankAnalysis(left);
    const b = rankAnalysis(right);
    for (let index = 0; index < a.length; index += 1) if (a[index] !== b[index]) return a[index] - b[index];
    return 0;
  }

  function compareTopRecoveryAnalyses(left, right) {
    if (!left) return 1;
    if (!right) return -1;
    const a = rankAnalysis(left);
    const b = rankAnalysis(right);
    for (let index = 0; index <= 4; index += 1) {
      if (a[index] !== b[index]) return a[index] - b[index];
    }
    const leftProposed = Number.isFinite(left.selectedProposedInches) ? left.selectedProposedInches : -1;
    const rightProposed = Number.isFinite(right.selectedProposedInches) ? right.selectedProposedInches : -1;
    if (leftProposed !== rightProposed) return rightProposed - leftProposed;
    for (let index = 5; index < a.length; index += 1) {
      if (a[index] !== b[index]) return a[index] - b[index];
    }
    return 0;
  }

  function statusForAnalysis(value) {
    return [STATUS.SAFE, STATUS.BEST_AVAILABLE, STATUS.CRITICAL][categoryForAnalysis(value)];
  }

  function resultMessage(value, status) {
    if (status === STATUS.SAFE) return "Pole and Midspan clearances satisfied.";
    if (status === STATUS.BEST_AVAILABLE) {
      return `Pole clearances satisfied. Midspan keeps ${value.midspanViolationCount} issue${value.midspanViolationCount === 1 ? "" : "s"} (${format(value.midspanViolationInches)} total shortfall).`;
    }
    if (status === STATUS.CRITICAL) {
      return `No pole-compliant aerial arrangement was found. Best candidate keeps ${value.poleViolationCount} pole issue${value.poleViolationCount === 1 ? "" : "s"} (${format(value.poleViolationInches)} total) and ${value.midspanViolationCount} Midspan issue${value.midspanViolationCount === 1 ? "" : "s"}.`;
    }
    return "Manual review required.";
  }

  function analyzeCurrentState(poleId, mode) {
    return { poleId, mode, ...collectIssues(poleId), ...movementMetrics(poleId, mode) };
  }

  function makeResult(mode, analysis, candidateCount, statusOverride = "") {
    const status = statusOverride || statusForAnalysis(analysis);
    const isIntec = String(S()?.getState?.()?.settings?.projectProfile || "INTEC").toUpperCase() === "INTEC";
    return {
      status,
      mode,
      message: resultMessage(analysis, status),
      poleViolationCount: analysis?.poleViolationCount || 0,
      poleViolationInches: analysis?.poleViolationInches || 0,
      midspanViolationCount: analysis?.midspanViolationCount || 0,
      midspanViolationInches: analysis?.midspanViolationInches || 0,
      movedCommCount: analysis?.movedCommCount || 0,
      totalMovementInches: analysis?.totalMovementInches || 0,
      proposedDistanceFromIdeal: Number.isFinite(analysis?.proposedDistanceFromIdeal) ? analysis.proposedDistanceFromIdeal : null,
      idealProposed: analysis?.idealProposedInches === null || analysis?.idealProposedInches === undefined ? "" : format(analysis.idealProposedInches),
      candidateCount,
      recommendation: status === STATUS.SAFE
        ? ""
        : isIntec
          ? "Review the best aerial result, then decide whether the pole should be UG or PCO. Nothing is selected automatically."
          : "Review the best available aerial result manually.",
      updatedAt: new Date().toISOString()
    };
  }

  function setResult(poleId, result) {
    const pole = S()?.getPole?.(poleId);
    if (!pole) return;
    pole.metadata = pole.metadata && typeof pole.metadata === "object" ? pole.metadata : {};
    pole.metadata.autoCalculateResult = result;
  }

  function clearAutomaticPolePlan(poleId) {
    (S()?.getSpanCommsForPole?.(poleId) || []).forEach(row => {
      if (row.autoCalcStatus !== AUTO) return;
      S().upsertSpanComm({
        ...row,
        existingHOAChange: "",
        autoCalcStatus: "",
        autoCalcMessage: ""
      });
    });
    (S()?.getSpanSidesForPole?.(poleId) || []).forEach(side => {
      if (side.autoCalcProposedStatus !== AUTO) return;
      S().upsertSpanSide({
        ...side,
        proposedHOA: "",
        autoCalcProposedStatus: "",
        autoCalcProposedMode: ""
      });
    });
  }

  async function solvePole(poleId, mode, options = {}) {
    const pole = S()?.getPole?.(poleId);
    if (!pole) return { status: STATUS.SKIPPED, applied: false };
    const debugTrace = beginPoleTrace(poleId, mode, options.tracePass || 1);
    if (pole.ugActive || pole.pcoActive) {
      const result = { status: STATUS.SKIPPED, mode, message: pole.ugActive ? "Pole is already marked UG." : "Pole is already marked PCO.", recommendation: "", candidateCount: 0, updatedAt: new Date().toISOString() };
      setResult(poleId, result);
      if (debugTrace) debugTrace.skipReason = pole.ugActive ? "POLE_ALREADY_UG" : "POLE_ALREADY_PCO";
      finishPoleTrace(debugTrace, result, poleId);
      return { status: result.status, applied: false, result };
    }
    const spans = eligibleProposedSpans(poleId);
    if (!spans.length) {
      const result = { status: STATUS.SKIPPED, mode, message: "No eligible Proposed span was found.", recommendation: "", candidateCount: 0, updatedAt: new Date().toISOString() };
      setResult(poleId, result);
      if (debugTrace) debugTrace.skipReason = "NO_ELIGIBLE_PROPOSED_SPAN";
      finishPoleTrace(debugTrace, result, poleId);
      return { status: result.status, applied: false, result };
    }
    const maxPole = parse(pole.maxCommHeight || "");
    if (maxPole === null) {
      const result = { status: STATUS.MANUAL, mode, message: "Missing Max Height on Pole. Auto Calculate cannot compare aerial arrangements.", recommendation: "Complete the pole power/equipment data and run Auto Calculate again.", candidateCount: 0, updatedAt: new Date().toISOString() };
      setResult(poleId, result);
      if (debugTrace) debugTrace.skipReason = "MISSING_MAX_HEIGHT_ON_POLE";
      finishPoleTrace(debugTrace, result, poleId);
      return { status: result.status, applied: false, result };
    }

    clearAutomaticPolePlan(poleId);
    recalculateAffected(poleId);
    const baseline = clone(S().getState());
    const groups = groupsForPole(poleId);
    const baselineAnalysis = analyzeCurrentState(poleId, mode);
    const recoverMidspanUpward = mode === "TOP_COMM" && baselineAnalysis.midspanViolationCount > 0;
    const manualReference = manualProposedReference(spans, poleId, mode);
    const current = currentProposed(spans, poleId);
    const candidates = candidateHeights({
      groups,
      maxPole,
      mode,
      currentProposed: current,
      manualReference,
      preferMaximum: recoverMidspanUpward,
      state: baseline
    });
    if (debugTrace) {
      debugTrace.automaticPlanCleared = true;
      debugTrace.maxHeightOnPole = format(maxPole);
      debugTrace.baseline = {
        groups: groups.map(groupSnapshot),
        analysis: analysisSnapshot(baselineAnalysis)
      };
      debugTrace.search = {
        strategy: recoverMidspanUpward ? "TOP_COMM_MIDSPAN_RECOVERY" : "STANDARD_PROGRESSIVE",
        progressiveCandidateCount: candidates.progressiveCount || 0,
        candidateOrder: candidates.map((value, index) => ({
          order: index + 1,
          proposedHOA: format(value),
          phase: index < (candidates.progressiveCount || 0) ? "PROGRESSIVE" : "FALLBACK"
        }))
      };
    }
    let best = null;
    if (spans.every(span => parse(S()?.getSpanSide?.(span.spanId, poleId)?.proposedHOA || "") !== null)) {
      best = { analysis: baselineAnalysis, state: clone(S().getState()) };
    }
    let evaluatedCount = 0;
    for (let candidateIndex = 0; candidateIndex < candidates.length; candidateIndex += 1) {
      const proposedInches = candidates[candidateIndex];
      S().setState(clone(baseline));
      applyProposed(spans, poleId, proposedInches, mode);
      const candidatePlan = buildStackPlan(
        groupsForPole(poleId),
        proposedInches,
        mode,
        maxPole,
        S().getState(),
        { preferHighest: recoverMidspanUpward }
      );
      applyPlan(candidatePlan);
      recalculateAffected(poleId);
      const analysis = analyzeCurrentState(poleId, mode);
      const comparison = recoverMidspanUpward
        ? compareTopRecoveryAnalyses(analysis, best?.analysis)
        : compareAnalyses(analysis, best?.analysis);
      const becameBest = !best || comparison < 0;
      if (!best || comparison < 0) {
        best = { analysis, state: clone(S().getState()) };
      }
      const evaluated = candidateIndex + 1;
      evaluatedCount = evaluated;
      const safe = statusForAnalysis(analysis) === STATUS.SAFE;
      const progressiveBestAvailable = mode === "TOP_COMM"
        && Number.isInteger(candidates.progressiveCount)
        && candidates.progressiveCount > 0
        && evaluated >= candidates.progressiveCount
        && best?.analysis?.poleViolationCount === 0;
      if (debugTrace) {
        debugTrace.candidates.push({
          order: evaluated,
          phase: candidateIndex < (candidates.progressiveCount || 0) ? "PROGRESSIVE" : "FALLBACK",
          proposedHOA: format(proposedInches),
          commPlan: candidatePlan.map(item => ({
            ...groupSnapshot(item.group),
            targetHOA: format(item.targetInches)
          })),
          analysis: analysisSnapshot(analysis),
          becameBest,
          decision: safe
            ? "ACCEPT_SAFE"
            : progressiveBestAvailable
              ? "STOP_WITH_BEST_PROGRESSIVE"
              : becameBest
                ? "KEEP_AS_CURRENT_BEST"
                : "REJECT_LOWER_RANK",
          stopReason: safe
            ? "SAFE_DISTRIBUTION_FOUND"
            : progressiveBestAvailable
              ? "POLE_COMPLIANT_PROGRESSIVE_OPTIONS_EXHAUSTED"
              : ""
        });
      }
      if (evaluated % CANDIDATE_YIELD_INTERVAL === 0 || evaluated === candidates.length || safe || progressiveBestAvailable) {
        reportProgress(options.onCandidateProgress, {
          poleId,
          candidateIndex: evaluated,
          candidateCount: safe || progressiveBestAvailable ? evaluated : candidates.length
        });
        await yieldToBrowser();
      }
      if (safe || progressiveBestAvailable) break;
    }
    if (!best) {
      S().setState(baseline);
      const result = { status: STATUS.MANUAL, mode, message: "No candidate arrangement could be evaluated.", recommendation: "Review the imported heights and span data.", candidateCount: evaluatedCount, updatedAt: new Date().toISOString() };
      setResult(poleId, result);
      if (debugTrace) debugTrace.skipReason = "NO_CANDIDATE_EVALUATED";
      finishPoleTrace(debugTrace, result, poleId);
      return { status: result.status, applied: false, result };
    }
    S().setState(best.state);
    recalculateAffected(poleId);
    const analysis = analyzeCurrentState(poleId, mode);
    const result = makeResult(mode, analysis, evaluatedCount);
    setResult(poleId, result);
    finishPoleTrace(debugTrace, result, poleId);
    return { status: result.status, applied: true, result, analysis };
  }

  function signature() {
    const state = S()?.getState?.() || {};
    const comms = Object.values(state.spanComms || {}).map(row => [commKey(row), row.existingHOAChange || "", row.autoCalcStatus || "", C()?.isCommMovementsActive?.(row.poleId) !== false]).sort((a, b) => a[0].localeCompare(b[0]));
    const proposed = Object.values(state.spanSides || {}).map(side => [S()?.keyForSpanSide?.(side.spanId, side.poleId) || `${side.spanId}__${side.poleId}`, side.proposedHOA || "", side.autoCalcProposedStatus || ""]).sort((a, b) => a[0].localeCompare(b[0]));
    return JSON.stringify({ comms, proposed });
  }

  function poleAutomaticSignature(poleId) {
    const comms = (S()?.getSpanCommsForPole?.(poleId) || [])
      .map(row => [commKey(row), row.existingHOAChange || "", row.autoCalcStatus || "", C()?.isCommMovementsActive?.(row.poleId) !== false])
      .sort((a, b) => a[0].localeCompare(b[0]));
    const proposed = (S()?.getSpanSidesForPole?.(poleId) || [])
      .map(side => [
        S()?.keyForSpanSide?.(side.spanId, side.poleId) || `${side.spanId}__${side.poleId}`,
        side.proposedHOA || "",
        side.autoCalcProposedStatus || ""
      ])
      .sort((a, b) => a[0].localeCompare(b[0]));
    return JSON.stringify({ comms, proposed });
  }

  function hasAutomaticMovement(poleId) {
    return (S()?.getSpanCommsForPole?.(poleId) || [])
      .some(row => row.autoCalcStatus === AUTO && text(row.existingHOAChange)
        && C()?.isCommMovementsActive?.(row.poleId) !== false);
  }

  function needsSelectiveRetry(poleId, mode) {
    const pole = S()?.getPole?.(poleId);
    if (!pole || pole.ugActive || pole.pcoActive) return false;
    const previous = pole.metadata?.autoCalculateResult;
    if (!previous || [STATUS.SKIPPED, STATUS.MANUAL].includes(previous.status)) return false;
    return statusForAnalysis(analyzeCurrentState(poleId, mode)) !== STATUS.SAFE
      || hasAutomaticMovement(poleId);
  }

  function selectivePoleSignature(poleId, mode) {
    return JSON.stringify({
      automatic: poleAutomaticSignature(poleId),
      rank: rankAnalysis(analyzeCurrentState(poleId, mode))
    });
  }

  async function autoCalculateMovements(options = {}) {
    if (!S()?.getState || !C()?.recalculateAll || !H()?.parseHeight) return { applied: 0, manual: 0, skipped: 0, safe: 0, bestAvailable: 0, critical: 0, passes: 0, converged: false, stoppedByRepeat: false, maxPassesReached: false, disabled: true };
    Object.values(S().getState().poles || {}).forEach(pole => { if (pole?.metadata?.autoCalculateResult) delete pole.metadata.autoCalculateResult; });
    const mode = modeFromState();
    const poleIds = Object.keys(S().getState().poles || {});
    lastDebugTrace = {
      version: 1,
      startedAt: new Date().toISOString(),
      mode,
      maxAttemptsPerPole: MAX_PASSES,
      poleOrder: [...poleIds],
      poleAttempts: []
    };
    const totalPoleSteps = Math.max(1, poleIds.length * MAX_PASSES);
    let currentSignature = signature();
    const seen = new Set([currentSignature]);
    const queue = poleIds.map(poleId => ({ poleId, pass: 1 }));
    const queued = new Set(poleIds);
    const processed = new Set();
    const attempts = new Map();
    let completedSteps = 0;
    let passes = poleIds.length ? 1 : 0;
    let converged = true;
    let stoppedByRepeat = false;
    reportProgress(options.onProgress, {
      phase: "starting",
      progress: 0,
      pass: 0,
      maxPasses: MAX_PASSES,
      poleIndex: 0,
      poleCount: poleIds.length,
      poleId: ""
    });
    await yieldToBrowser();
    C().recalculateAll();
    while (queue.length && completedSteps < totalPoleSteps) {
      const item = queue.shift();
      const poleId = item.poleId;
      queued.delete(poleId);
      passes = Math.max(passes, item.pass);
      const affected = affectedPoles(poleId);
      const before = new Map(affected.map(id => [id, selectivePoleSignature(id, mode)]));
      const completedBeforePole = completedSteps;
      const poleIndex = poleIds.indexOf(poleId);
      reportProgress(options.onProgress, {
        phase: "pole",
        progress: (completedBeforePole / totalPoleSteps) * 100,
        pass: item.pass,
        maxPasses: MAX_PASSES,
        poleIndex: poleIndex + 1,
        poleCount: poleIds.length,
        poleId
      });
      await solvePole(poleId, mode, {
        tracePass: item.pass,
        onCandidateProgress(candidate) {
          const fraction = candidate.candidateCount
            ? candidate.candidateIndex / candidate.candidateCount
            : 1;
          reportProgress(options.onProgress, {
            phase: "candidate",
            progress: ((completedBeforePole + fraction) / totalPoleSteps) * 100,
            pass: item.pass,
            maxPasses: MAX_PASSES,
            poleIndex: poleIndex + 1,
            poleCount: poleIds.length,
            poleId,
            candidateIndex: candidate.candidateIndex,
            candidateCount: candidate.candidateCount
          });
        }
      });
      completedSteps += 1;
      processed.add(poleId);
      attempts.set(poleId, (attempts.get(poleId) || 0) + 1);

      const changedAffected = affected.filter(id => before.get(id) !== selectivePoleSignature(id, mode));
      changedAffected.forEach(affectedPoleId => {
        if (affectedPoleId === poleId || !processed.has(affectedPoleId) || queued.has(affectedPoleId)) return;
        const count = attempts.get(affectedPoleId) || 0;
        if (count >= MAX_PASSES || !needsSelectiveRetry(affectedPoleId, mode)) return;
        queue.push({ poleId: affectedPoleId, pass: count + 1 });
        queued.add(affectedPoleId);
      });

      const next = signature();
      if (next !== currentSignature) {
        if (seen.has(next)) {
          stoppedByRepeat = true;
          queue.length = 0;
          break;
        }
        seen.add(next);
        currentSignature = next;
      }
      await yieldToBrowser();
    }
    converged = !queue.length && !stoppedByRepeat;
    const maxPassesReached = !converged && !stoppedByRepeat && completedSteps >= totalPoleSteps;
    const results = Object.values(S().getState().poles || {}).map(pole => pole?.metadata?.autoCalculateResult).filter(Boolean);
    const safe = results.filter(result => result.status === STATUS.SAFE).length;
    const bestAvailable = results.filter(result => result.status === STATUS.BEST_AVAILABLE).length;
    const critical = results.filter(result => result.status === STATUS.CRITICAL).length;
    const manualOnly = results.filter(result => result.status === STATUS.MANUAL).length;
    const skipped = results.filter(result => result.status === STATUS.SKIPPED).length;
    reportProgress(options.onProgress, {
      phase: "complete",
      progress: 100,
      pass: passes,
      maxPasses: MAX_PASSES,
      poleIndex: poleIds.length,
      poleCount: poleIds.length,
      poleId: ""
    });
    const summary = { applied: safe + bestAvailable + critical, manual: bestAvailable + critical + manualOnly, skipped, safe, bestAvailable, critical, passes, converged, stoppedByRepeat, maxPassesReached, disabled: false };
    lastDebugTrace.completedAt = new Date().toISOString();
    lastDebugTrace.summary = clone(summary);
    return summary;
  }

  function installManualTracking() {
    const api = C();
    if (!api || api.__autoCalculateManualTrackingInstalled) return;
    if (typeof api.updateExistingHOAChange === "function") {
      const original = api.updateExistingHOAChange.bind(api);
      api.updateExistingHOAChange = function (...args) {
        const result = original(...args);
        if (result) S()?.upsertSpanComm?.({ ...result, autoCalcStatus: "", autoCalcMessage: "" });
        return result;
      };
    }
    if (typeof api.updateSpanCommField === "function") {
      const original = api.updateSpanCommField.bind(api);
      api.updateSpanCommField = function (spanId, poleId, commOwner, wireId, field, value) {
        const result = original(spanId, poleId, commOwner, wireId, field, value);
        if (result && field === "existingHOAChange") S()?.upsertSpanComm?.({ ...result, autoCalcStatus: "", autoCalcMessage: "" });
        return result;
      };
    }
    if (typeof api.updateSpanSideField === "function") {
      const original = api.updateSpanSideField.bind(api);
      api.updateSpanSideField = function (spanId, poleId, field, value) {
        const result = original(spanId, poleId, field, value);
        if (result && field === "proposedHOA") S()?.upsertSpanSide?.({ ...result, autoCalcProposedStatus: "", autoCalcProposedMode: "" });
        return result;
      };
    }
    api.__autoCalculateManualTrackingInstalled = true;
  }

  function install() {
    const api = C();
    if (!api || api.__bestAvailableAutoCalculateInstalled) return false;
    installManualTracking();
    api.autoCalculateMovements = autoCalculateMovements;
    api.__bestAvailableAutoCalculateInstalled = true;
    return true;
  }

  global.AutoCalculateSolver = {
    RESULT_STATUS: STATUS,
    modeFromState,
    normalizeOwner,
    groupsForPole,
    proposedSpansForPole: eligibleProposedSpans,
    idealProposedHeight,
    candidateHeights,
    buildStackPlan,
    splitIssueMessage,
    isMidspanIssue,
    physicalSpanPair,
    collectIssues,
    issueSeverity,
    categoryForAnalysis,
    rankAnalysis,
    compareAnalyses,
    compareTopRecoveryAnalyses,
    statusForAnalysis,
    resultMessage,
    analyzeCurrentState,
    solvePole,
    autoCalculateMovements,
    getDebugTrace: () => lastDebugTrace ? clone(lastDebugTrace) : null,
    install
  };

  install();
})(window);
