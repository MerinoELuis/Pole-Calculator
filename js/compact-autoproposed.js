(function (global) {
  "use strict";

  const DEFAULT_MESSENGER_144 = "0.242";
  const DEFAULT_FIBER_144 = "0.51";
  const DEFAULT_COX_MESSENGER_96 = "0.25";
  const DEFAULT_COX_FIBER_96 = "0.53";
  const SUPPORTED_PROFILES = new Set(["INTEC", "METRONET"]);
  const OPPOSITE_DIRECTION = {
    N: "S",
    NE: "SW",
    E: "W",
    SE: "NW",
    S: "N",
    SW: "NE",
    W: "E",
    NW: "SE"
  };
  const BEARING_DIRECTIONS = ["N", "NE", "E", "SE", "S", "SW", "W", "NW"];
  const DIRECTION_DEGREES = {
    N: 0,
    NE: 45,
    E: 90,
    SE: 135,
    S: 180,
    SW: 225,
    W: 270,
    NW: 315
  };
  const DIRECTION_MATCH_TOLERANCE_DEGREES = 45;

  const S = () => global.AppStore;
  const H = () => global.HeightUtils;

  function text(value) {
    return String(value ?? "").trim();
  }

  function escapeHtml(value) {
    return String(value ?? "")
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&#039;");
  }

  function fiberCount(value) {
    const match = text(value).match(/\b(\d+)\s*CT\b/i);
    return match ? String(Number(match[1])) : "";
  }

  function canonicalFiberName(count) {
    return count ? `${Number(count)}CT Fiber` : "";
  }

  function profileName(state) {
    return text(state?.settings?.projectProfile || "INTEC").toUpperCase();
  }

  function detectedReferenceFiberCounts(state) {
    const counts = new Set();
    (state?.makeReadyReferences || []).forEach(ref => {
      const count = fiberCount(`${ref?.attachmentFiber || ""} ${ref?.attachmentSizeRaw || ""}`);
      if (count) counts.add(count);
    });
    return counts;
  }

  function fiberEntries(state) {
    const entries = new Map();
    const configured = state?.settings?.fiberSizes && typeof state.settings.fiberSizes === "object"
      ? state.settings.fiberSizes
      : {};

    Object.keys(configured).forEach(name => {
      const count = fiberCount(name);
      if (count && !entries.has(count)) entries.set(count, name);
    });

    (state?.makeReadyReferences || []).forEach(ref => {
      const source = `${ref?.attachmentFiber || ""} ${ref?.attachmentSizeRaw || ""}`;
      const count = fiberCount(source);
      if (count && !entries.has(count)) entries.set(count, canonicalFiberName(count));
    });

    return Array.from(entries.entries())
      .map(([count, name]) => ({ count, name }))
      .sort((a, b) => Number(a.count) - Number(b.count));
  }

  function configuredFiberKey(state, count) {
    const configured = state?.settings?.fiberSizes && typeof state.settings.fiberSizes === "object"
      ? state.settings.fiberSizes
      : {};
    return Object.keys(configured).find(name => fiberCount(name) === String(Number(count))) || canonicalFiberName(count);
  }

  function applyAttachmentDefaults(state = S()?.getState?.()) {
    if (!state || !SUPPORTED_PROFILES.has(profileName(state))) return false;
    const detectedCounts = detectedReferenceFiberCounts(state);
    const has144 = detectedCounts.has("144");
    const hasCox96 = text(state?.settings?.proposedOwner).toUpperCase() === "COX" && detectedCounts.has("96");
    if (!has144 && !hasCox96) return false;

    state.settings = state.settings || {};
    state.settings.fiberSizes = state.settings.fiberSizes && typeof state.settings.fiberSizes === "object"
      ? state.settings.fiberSizes
      : {};

    let changed = false;
    if (!text(state.settings.attachmentMessengerSize)) {
      state.settings.attachmentMessengerSize = hasCox96 ? DEFAULT_COX_MESSENGER_96 : DEFAULT_MESSENGER_144;
      changed = true;
    }

    if (has144) {
      const key144 = configuredFiberKey(state, "144");
      if (!text(state.settings.fiberSizes[key144])) {
        state.settings.fiberSizes[key144] = DEFAULT_FIBER_144;
        changed = true;
      }
    }
    if (hasCox96) {
      const key96 = configuredFiberKey(state, "96");
      if (!text(state.settings.fiberSizes[key96])) {
        state.settings.fiberSizes[key96] = DEFAULT_COX_FIBER_96;
        changed = true;
      }
    }
    return changed;
  }

  function numericSize(value) {
    const parsed = Number(text(value));
    return Number.isFinite(parsed) && parsed > 0 ? parsed : null;
  }

  function inches(value) {
    if (typeof value === "number" && Number.isFinite(value)) return Math.round(value);
    const parsed = H()?.parseHeight?.(value);
    return parsed === null || parsed === undefined || !Number.isFinite(Number(parsed))
      ? null
      : Math.round(Number(parsed));
  }

  function spanLengthInches(span) {
    const display = inches(span?.lengthDisplay || "");
    if (display !== null) return display;
    const raw = Number(span?.length);
    if (!Number.isFinite(raw) || raw <= 0) return null;
    return Math.round(raw * 12);
  }

  function ownerForMove(row) {
    const raw = text(row?.rawOwner || row?.ownerBase || row?.owner || "COMM");
    if (/century\s*link|centurylink|\bctl\b/i.test(raw)) return "CTL";
    return raw.replace(/^COMMUNICATION\s*>\s*/i, "").replace(/,\s*.*$/, "").trim() || "COMM";
  }

  function bearingDegrees(span) {
    const raw = text(span?.bearingDegrees);
    const value = Number(raw);
    return raw && Number.isFinite(value) ? value : null;
  }

  function directionFromBearing(value) {
    const bearing = Number(value);
    if (!Number.isFinite(bearing)) return "";
    const normalized = ((bearing % 360) + 360) % 360;
    return BEARING_DIRECTIONS[Math.round(normalized / 45) % 8];
  }

  function normalizeBearing(value) {
    const bearing = Number(value);
    if (!Number.isFinite(bearing)) return null;
    return ((bearing % 360) + 360) % 360;
  }

  function directionBearing(value) {
    const direction = text(value).toUpperCase();
    return Object.prototype.hasOwnProperty.call(DIRECTION_DEGREES, direction)
      ? DIRECTION_DEGREES[direction]
      : null;
  }

  function angularDistance(first, second) {
    const a = normalizeBearing(first);
    const b = normalizeBearing(second);
    if (a === null || b === null) return null;
    const difference = Math.abs(a - b);
    return Math.min(difference, 360 - difference);
  }

  function directionFromPole(span, poleId) {
    if (!span) return "";
    const imported = text(span.direction).toUpperCase();
    if (span.fromPole === poleId) return imported || directionFromBearing(bearingDegrees(span));
    if (imported) return OPPOSITE_DIRECTION[imported] || "";
    const bearing = bearingDegrees(span);
    return bearing === null ? "" : directionFromBearing(bearing + 180);
  }

  function bearingFromPole(span, poleId) {
    if (!span) return null;
    const importedBearing = normalizeBearing(bearingDegrees(span));
    if (importedBearing !== null) {
      return span.fromPole === poleId ? importedBearing : normalizeBearing(importedBearing + 180);
    }
    return directionBearing(directionFromPole(span, poleId));
  }

  function directionTokensForReference(ref) {
    if (Array.isArray(ref?.attachmentDirectionTokens) && ref.attachmentDirectionTokens.length) {
      return Array.from(new Set(ref.attachmentDirectionTokens
        .map(value => text(value).toUpperCase())
        .filter(value => BEARING_DIRECTIONS.includes(value))));
    }

    const source = [
      ref?.attachmentDirection,
      ref?.attachmentFiber,
      ref?.attachmentSizeRaw,
      ref?.attachmentType
    ].map(text).filter(Boolean).join(" ").toUpperCase();
    return Array.from(new Set(source.match(/\b(?:NE|NW|SE|SW|N|E|S|W)\b/g) || []));
  }

  function fiberReferencesForPole(state, poleId) {
    return (state?.makeReadyReferences || []).filter(ref => {
      if (ref?.poleId !== poleId) return false;
      return Boolean(fiberCount(`${ref?.attachmentFiber || ""} ${ref?.attachmentSizeRaw || ""}`));
    });
  }

  function referenceForSpan(state, poleId, span) {
    const refs = fiberReferencesForPole(state, poleId);
    if (!refs.length) return null;

    const direction = directionFromPole(span, poleId);
    const exact = direction
      ? refs.find(ref => directionTokensForReference(ref).includes(direction))
      : null;
    if (exact) return exact;

    // Make Ready directions are often entered visually (for example, SE for
    // a span whose actual bearing is closer to E). Accept the nearest
    // directional reference within one adjacent compass sector instead of
    // dropping an otherwise valid proposed attachment.
    const spanBearing = bearingFromPole(span, poleId);
    if (spanBearing !== null) {
      let nearest = null;
      refs.forEach(ref => {
        const distance = Math.min(...directionTokensForReference(ref)
          .map(token => angularDistance(spanBearing, directionBearing(token)))
          .filter(value => value !== null));
        if (!Number.isFinite(distance) || distance > DIRECTION_MATCH_TOLERANCE_DEGREES) return;
        if (!nearest || distance < nearest.distance) nearest = { ref, distance };
      });
      if (nearest) return nearest.ref;
    }

    return refs.find(ref => directionTokensForReference(ref).length === 0) || null;
  }

  function fiberForSpan(state, poleId, span) {
    const ref = referenceForSpan(state, poleId, span);
    return ref ? fiberCount(`${ref.attachmentFiber || ""} ${ref.attachmentSizeRaw || ""}`) : "";
  }

  function spanKind(span) {
    const type = text(span?.type || span?.rawType).toLowerCase();
    if (/fore\s*span|forespan/.test(type)) return "F";
    if (/back\s*span|backspan/.test(type)) return "B";
    return "O";
  }

  function hasUgToken(value) {
    return /(?:^|\s)UG(?:\s|$)/i.test(text(value));
  }

  function hasPcoToken(value) {
    return /(?:^|\s)PCO(?:\s|$)/i.test(text(value));
  }

  function findPole(state, poleId) {
    if (!poleId) return null;
    if (state?.poles?.[poleId]) return state.poles[poleId];
    const canonical = S()?.canonicalPoleIdentity?.(poleId);
    if (canonical) {
      const match = Object.values(state?.poles || {}).find(pole =>
        S()?.canonicalPoleIdentity?.(pole?.poleId || "") === canonical);
      if (match) return match;
    }
    return S()?.getPole?.(poleId) || null;
  }

  function isPoleFullyUg(state, poleId) {
    const pole = findPole(state, poleId);
    return Boolean(pole?.ugActive || hasUgToken(poleId));
  }

  function isPolePco(state, poleId) {
    const pole = findPole(state, poleId);
    return Boolean(pole?.pcoActive || hasPcoToken(poleId));
  }

  function blocksLocalActions(state, poleId) {
    return isPoleFullyUg(state, poleId) || isPolePco(state, poleId);
  }

  function geometryOnlySpan(span, markUg) {
    const item = { to: span.to, kind: span.kind };
    if (Object.prototype.hasOwnProperty.call(span, "bearing")) item.bearing = span.bearing;
    if (Object.prototype.hasOwnProperty.call(span, "length")) item.length = span.length;
    if (markUg) item.ug = true;
    return item;
  }

  function sanitizeBlockedLocalPole(pole, state) {
    if (!pole) return pole;
    const fullyUg = isPoleFullyUg(state, pole.id);
    const pco = !fullyUg && isPolePco(state, pole.id);
    if (!fullyUg && !pco) return pole;

    const sanitized = { id: pole.id };
    if (Array.isArray(pole.spans) && pole.spans.length) {
      sanitized.spans = pole.spans.map(span => geometryOnlySpan(span, fullyUg));
    }
    return sanitized;
  }

  function hasExportablePoleData(pole) {
    return Boolean(
      pole && (
        (Array.isArray(pole.spans) && pole.spans.length) ||
        (Array.isArray(pole.moves) && pole.moves.length) ||
        Object.prototype.hasOwnProperty.call(pole, "terminalHoa")
      )
    );
  }

  function sanitizeBlockedLocalPayload(payload, state = S()?.getState?.()) {
    if (!payload || !Array.isArray(payload.poles)) return payload;
    return {
      ...payload,
      poles: payload.poles
        .map(pole => sanitizeBlockedLocalPole(pole, state))
        .filter(hasExportablePoleData)
    };
  }

  function isUgSpan(state, span, poleId = span?.fromPole) {
    if (!span) return false;
    if (isPoleFullyUg(state, poleId || span.fromPole)) return true;

    const targetId = text(span.toPole);
    const targetPole = findPole(state, targetId);
    const explicitSource = `${targetId} ${span.rawType || ""} ${span.notes || ""}`;
    return Boolean(targetPole?.ugActive || hasUgToken(explicitSource));
  }

  function sideForSpan(state, poleId, spanId) {
    return Object.values(state?.spanSides || {}).find(side => side.poleId === poleId && side.spanId === spanId) || null;
  }

  function primaryProposalForPole(state, poleId) {
    return Object.values(state?.spanSides || {})
      .filter(side => side.poleId === poleId && inches(side.proposedHOA) !== null)
      .sort((a, b) => Number(Boolean(a.isAdditionalProposed)) - Number(Boolean(b.isAdditionalProposed)))[0] || null;
  }

  function hasSpanGeometry(span) {
    return Boolean(text(span?.toPole) && (
      bearingDegrees(span) !== null || spanLengthInches(span) !== null || text(span?.direction)
    ));
  }

  function compactSpan(state, poleId, span, primaryProposal) {
    const item = {
      to: text(span.toPole),
      kind: spanKind(span)
    };
    const bearing = bearingDegrees(span);
    const length = spanLengthInches(span);
    if (bearing !== null) item.bearing = bearing;
    if (length !== null) item.length = length;

    if (isUgSpan(state, span, poleId)) {
      item.ug = true;
      return item;
    }

    const fiber = fiberForSpan(state, poleId, span);
    if (!fiber) return item;

    const exactSide = sideForSpan(state, poleId, span.spanId);
    const proposal = inches(exactSide?.proposedHOA) !== null ? exactSide : primaryProposal;
    const hoa = inches(proposal?.proposedHOA);
    if (hoa === null) return item;

    item.hoa = hoa;
    item.fiber = Number(fiber);

    if (proposal === exactSide) {
      const endDrop = inches(exactSide?.endDrop);
      const nextHoa = inches(exactSide?.proposedHOAChange);
      if (endDrop !== null) item.endDrop = endDrop;
      if (nextHoa !== null) item.nextHoa = nextHoa;
    }
    return item;
  }

  function normalizedPhysicalPoleId(value) {
    return text(S()?.canonicalPoleIdentity?.(value) || value).toUpperCase();
  }

  function physicalSpanKey(span) {
    const bearing = Number.isFinite(Number(span?.bearing)) ? Number(span.bearing).toFixed(2) : "";
    const length = Number.isFinite(Number(span?.length)) ? String(Math.round(Number(span.length))) : "";
    return `${normalizedPhysicalPoleId(span?.to)}|${bearing}|${length}`;
  }

  function spanProposalScore(span) {
    return Number("hoa" in span) + Number("fiber" in span) +
      (Number("endDrop" in span) * 2) + (Number("nextHoa" in span) * 2) +
      Number(span.kind !== "O");
  }

  function mergeCompactSpans(existing, incoming) {
    if (!existing) return { ...incoming };
    if (existing.ug || incoming.ug) {
      const preferredGeometry = spanProposalScore(incoming) > spanProposalScore(existing) ? incoming : existing;
      const result = {
        to: preferredGeometry.to || existing.to || incoming.to,
        kind: preferredGeometry.kind || existing.kind || incoming.kind
      };
      if (Number.isFinite(Number(preferredGeometry.bearing))) result.bearing = Number(preferredGeometry.bearing);
      else if (Number.isFinite(Number(existing.bearing))) result.bearing = Number(existing.bearing);
      else if (Number.isFinite(Number(incoming.bearing))) result.bearing = Number(incoming.bearing);
      if (Number.isFinite(Number(preferredGeometry.length))) result.length = Number(preferredGeometry.length);
      else if (Number.isFinite(Number(existing.length))) result.length = Number(existing.length);
      else if (Number.isFinite(Number(incoming.length))) result.length = Number(incoming.length);
      result.ug = true;
      return result;
    }

    const preferred = spanProposalScore(incoming) > spanProposalScore(existing) ? incoming : existing;
    const secondary = preferred === incoming ? existing : incoming;
    const result = { ...preferred };
    ["bearing", "length", "hoa", "fiber", "endDrop", "nextHoa"].forEach(field => {
      if (!(field in result) && field in secondary) result[field] = secondary[field];
    });
    if (result.kind === "O" && secondary.kind && secondary.kind !== "O") result.kind = secondary.kind;
    return result;
  }

  function compactSpansForPole(state, poleId) {
    const primaryProposal = primaryProposalForPole(state, poleId);
    const byPhysicalSpan = new Map();

    Object.values(state?.spans || {})
      .filter(span => span.fromPole === poleId && hasSpanGeometry(span))
      .map(span => compactSpan(state, poleId, span, primaryProposal))
      .forEach(span => {
        const key = physicalSpanKey(span);
        byPhysicalSpan.set(key, mergeCompactSpans(byPhysicalSpan.get(key), span));
      });

    return Array.from(byPhysicalSpan.values())
      .sort((a, b) => `${a.kind}|${a.to}`.localeCompare(`${b.kind}|${b.to}`, undefined, { numeric: true }));
  }

  function movementInstructionKey(candidate) {
    return [
      candidate.owner.toLowerCase(),
      candidate.from,
      candidate.to,
      candidate.service ? "service" : "normal"
    ].join("|");
  }

  function movementCandidates(state, poleId) {
    const candidates = [];
    const byInstruction = new Map();

    Object.values(state?.spanComms || {}).forEach(row => {
      if (row.poleId !== poleId) return;
      const from = inches(row.existingHOA);
      const to = inches(row.existingHOAChange || (row.transferToNewPole ? row.existingHOA : ""));
      if (from === null || to === null) return;
      if (from === to && !row.transferToNewPole) return;

      const candidate = {
        owner: ownerForMove(row),
        from,
        to,
        service: Boolean(row.serviceDrop),
        // Service drops are never relocated with a down guy. Keep this
        // invariant at the movement-candidate boundary so both the MR text
        // and the compact payload stay consistent even when imported data
        // incorrectly carries a DG flag on a drop.
        dg: !row.serviceDrop && Boolean(row.downGuy),
        transfer: Boolean(row.transferToNewPole)
      };
      const key = movementInstructionKey(candidate);
      const existing = byInstruction.get(key);
      if (existing) {
        existing.dg = existing.dg || candidate.dg;
        existing.transfer = existing.transfer || candidate.transfer;
        return;
      }
      byInstruction.set(key, candidate);
      candidates.push(candidate);
    });

    return candidates;
  }

  function compactMoves(state, poleId) {
    return movementCandidates(state, poleId).map(candidate => {
      const item = { owner: candidate.owner, from: candidate.from, to: candidate.to };
      if (candidate.service) item.service = true;
      if (candidate.dg && !candidate.service) item.dg = true;
      return item;
    });
  }

  function buildCompactPayload(state = S()?.getState?.()) {
    if (!state) return { sizes: { messenger: null, fiber: {} }, owner: "", poles: [] };
    applyAttachmentDefaults(state);

    const poles = [];
    Object.keys(state.poles || {})
      .sort((a, b) => a.localeCompare(b, undefined, { numeric: true }))
      .forEach(poleId => {
        const spans = compactSpansForPole(state, poleId);
        const moves = compactMoves(state, poleId);
        const terminalHoa = isPoleFullyUg(state, poleId)
          ? null
          : inches(state.poles[poleId]?.standaloneProposedHOA);
        if (!spans.length && !moves.length && terminalHoa === null) return;

        const pole = { id: poleId };
        if (terminalHoa !== null) pole.terminalHoa = terminalHoa;
        if (spans.length) pole.spans = spans;
        if (moves.length) pole.moves = moves;
        poles.push(pole);
      });

    const usedFibers = new Set();
    poles.forEach(pole => (pole.spans || []).forEach(span => {
      if (!span.ug && Number.isFinite(Number(span.fiber))) usedFibers.add(String(Number(span.fiber)));
    }));

    const fiberSizes = {};
    usedFibers.forEach(count => {
      const key = configuredFiberKey(state, count);
      fiberSizes[count] = numericSize(state.settings?.fiberSizes?.[key]);
    });

    return sanitizeBlockedLocalPayload({
      sizes: {
        messenger: numericSize(state.settings?.attachmentMessengerSize),
        fiber: fiberSizes
      },
      owner: text(state.settings?.proposedOwner || "Wecom") || "Wecom",
      poles
    }, state);
  }

  function validationErrors(payload) {
    const usedFibers = new Set();
    (payload?.poles || []).forEach(pole => (pole.spans || []).forEach(span => {
      if (!span.ug && Number.isFinite(Number(span.fiber))) usedFibers.add(String(Number(span.fiber)));
    }));
    if (!usedFibers.size) return [];

    const errors = [];
    if (numericSize(payload?.sizes?.messenger) === null) errors.push("Missing Messenger Size");
    Array.from(usedFibers).sort((a, b) => Number(a) - Number(b)).forEach(count => {
      if (numericSize(payload?.sizes?.fiber?.[count]) === null) errors.push(`Missing ${count}CT Fiber Size`);
    });
    return errors;
  }

  function safeJobFilePart(value) {
    const raw = text(value || "pole_job")
      .replace(/\.[^.]+$/, "")
      .replace(/_(?:Pole_Calculator|AutoProposed|Debug)$/i, "")
      .replace(/^excel[_\-\s]*/i, "")
      .replace(/[_\-\s]*20\d{2}[-_]\d{2}[-_]\d{2}$/i, "")
      .trim();
    return (raw || "pole_job")
      .replace(/[<>:"/\\|?*\u0000-\u001F]+/g, "_")
      .replace(/\s+/g, "_");
  }

  function notify(message, type = "warning") {
    if (global.document) {
      const host = global.document.getElementById("toastHost");
      if (host) {
        const item = global.document.createElement("div");
        item.className = `toast ${type}`;
        item.textContent = message;
        host.appendChild(item);
        global.setTimeout?.(() => item.remove(), 4200);
        return;
      }
    }
    global.alert?.(message);
  }

  function exportCompactProposedJson() {
    global.Calculations?.recalculateAll?.();
    const state = S()?.getState?.();
    const payload = buildCompactPayload(state);
    const errors = validationErrors(payload);
    if (errors.length) {
      errors.forEach(error => notify(error, "warning"));
      return false;
    }
    const jobName = safeJobFilePart(state?.jobName || state?.importedFileName || "pole_job");
    global.ProjectExport?.downloadJson?.(`${jobName}_AutoProposed.json`, payload);
    return true;
  }

  function movementMrLine(candidate) {
    if (candidate.transfer) return "";
    const from = H()?.formatHeight?.(candidate.from) || String(candidate.from);
    const to = H()?.formatHeight?.(candidate.to) || String(candidate.to);
    if (candidate.service) return `Relocate ${candidate.owner} drop at HOA ${from} to HOA ${to}.`;
    const dg = candidate.dg ? " with DG" : "";
    const verb = candidate.to > candidate.from ? "raise" : "lower";
    return `At HOA ${from} ${verb} ${candidate.owner} to HOA ${to}${dg}.`;
  }

  function applyMrCase(value, state) {
    return text(state?.settings?.mrCase).toUpperCase() === "UPPER" ? value.toUpperCase() : value;
  }

  function normalizeMrLine(value) {
    return text(value).replace(/\s+/g, " ").toLowerCase();
  }

  function generatedMovementLineKeys(state, poleId) {
    return new Set(
      movementCandidates(state, poleId)
        .map(movementMrLine)
        .map(normalizeMrLine)
        .filter(Boolean)
    );
  }

  function sanitizeBlockedLocalMr(state, poleId) {
    if (!state || !Array.isArray(state.mr) || !blocksLocalActions(state, poleId)) {
      return state?.mr?.filter(item => item.poleId === poleId) || [];
    }

    const generatedMovementLines = generatedMovementLineKeys(state, poleId);
    state.mr = state.mr.reduce((items, item) => {
      if (item?.poleId !== poleId) {
        items.push(item);
        return items;
      }

      const remainingLines = text(item.text)
        .split(/\r?\n/)
        .map(line => line.trim())
        .filter(Boolean)
        .filter(line => !generatedMovementLines.has(normalizeMrLine(line)));

      if (remainingLines.length) items.push({ ...item, text: remainingLines.join("\n") });
      return items;
    }, []);

    return state.mr.filter(item => item.poleId === poleId);
  }

  function augmentPoleMakeReady(poleId) {
    const state = S()?.getState?.();
    if (!state) return [];
    const generated = movementCandidates(state, poleId)
      .map(movementMrLine)
      .filter(Boolean)
      .map(line => applyMrCase(line, state));
    if (!generated.length) return state.mr?.filter(item => item.poleId === poleId) || [];

    state.mr = Array.isArray(state.mr) ? state.mr : [];
    let item = state.mr.find(row => row.poleId === poleId);
    if (!item) {
      item = { poleId, spanId: "", owner: "MR", text: "", imported: false };
      state.mr.push(item);
    }
    const lines = text(item.text).split(/\n+/).map(line => line.trim()).filter(Boolean);
    const seen = new Set(lines.map(line => line.toLowerCase()));
    generated.forEach(line => {
      const key = line.toLowerCase();
      if (seen.has(key)) return;
      const withoutDg = key.replace(/ with dg\.$/, ".");
      const baseIndex = lines.findIndex(existing => existing.toLowerCase() === withoutDg);
      if (baseIndex >= 0 && withoutDg !== key) {
        seen.delete(lines[baseIndex].toLowerCase());
        lines[baseIndex] = line;
        seen.add(key);
        return;
      }
      seen.add(key);
      lines.push(line);
    });
    item.text = lines.join("\n");
    return state.mr.filter(row => row.poleId === poleId);
  }

  function patchMrLogic() {
    const logic = global.MRLogic;
    if (!logic || logic.__compactAutoProposedPatched || typeof logic.generateMRForPole !== "function") return;
    const originalGenerateMRForPole = logic.generateMRForPole.bind(logic);
    logic.generateMRForPole = function (poleId) {
      originalGenerateMRForPole(poleId);
      augmentPoleMakeReady(poleId);
      return sanitizeBlockedLocalMr(S()?.getState?.(), poleId);
    };
    logic.generateAllMR = function () {
      const state = S()?.getState?.();
      if (!state) return [];
      state.mr = [];
      Object.keys(state.poles || {}).forEach(poleId => logic.generateMRForPole(poleId));
      return state.mr;
    };
    logic.__compactAutoProposedPatched = true;
  }

  function markSaveDirty() {
    if (!global.document) return;
    const button = global.document.getElementById("saveLocalBtn");
    if (button && !/\*$/.test(button.textContent || "")) button.textContent = `${text(button.textContent) || "Save"} *`;
  }

  function updateInjectedSetting(input) {
    const state = S()?.getState?.();
    if (!state) return;
    state.settings = state.settings || {};
    if (input.dataset.field === "attachmentMessengerSize") {
      state.settings.attachmentMessengerSize = text(input.value);
    } else if (input.dataset.field === "fiberSize" && input.dataset.fiber) {
      state.settings.fiberSizes = state.settings.fiberSizes && typeof state.settings.fiberSizes === "object" ? state.settings.fiberSizes : {};
      state.settings.fiberSizes[input.dataset.fiber] = text(input.value);
    }
    markSaveDirty();
  }

  function createFiberSection(state, fibers) {
    const section = global.document.createElement("div");
    section.className = "settings-section settings-section-fiber";
    section.dataset.compactFiberSection = "true";
    section.innerHTML = `<h3>Fiber</h3><div class="settings-grid fiber-grid">
      <label class="clearance-row"><span>Messenger Size</span><input class="input" data-scope="attachmentSettings" data-field="attachmentMessengerSize" value="${escapeHtml(state.settings?.attachmentMessengerSize || "")}" /></label>
      ${fibers.map(item => {
        const key = configuredFiberKey(state, item.count);
        return `<label class="clearance-row"><span>${escapeHtml(item.name)} Size</span><input class="input" data-scope="attachmentSettings" data-field="fiberSize" data-fiber="${escapeHtml(key)}" value="${escapeHtml(state.settings?.fiberSizes?.[key] || "")}" /></label>`;
      }).join("")}
    </div>`;
    section.querySelectorAll("input[data-scope='attachmentSettings']").forEach(input => {
      input.addEventListener("input", () => updateInjectedSetting(input));
      input.addEventListener("change", () => updateInjectedSetting(input));
    });
    return section;
  }

  let syncingUi = false;
  function syncFiberSettingsUi() {
    if (syncingUi || !global.document || !S()?.getState) return;
    syncingUi = true;
    try {
      const state = S().getState();
      const changed = applyAttachmentDefaults(state);
      if (changed) markSaveDirty();
      const container = global.document.getElementById("clearanceSettings");
      if (!container) return;
      const fibers = fiberEntries(state);
      const supported = SUPPORTED_PROFILES.has(profileName(state));
      let section = container.querySelector(".settings-section-fiber");
      if (!supported || !fibers.length) {
        if (section?.dataset.compactFiberSection === "true") section.remove();
        return;
      }
      if (!section) {
        section = createFiberSection(state, fibers);
        container.appendChild(section);
      }
      const messenger = section.querySelector('[data-field="attachmentMessengerSize"]');
      if (messenger && global.document.activeElement !== messenger) messenger.value = state.settings?.attachmentMessengerSize || "";
      fibers.forEach(item => {
        const key = configuredFiberKey(state, item.count);
        const input = Array.from(section.querySelectorAll('[data-field="fiberSize"]'))
          .find(row => fiberCount(row.dataset.fiber) === item.count);
        if (input && global.document.activeElement !== input) input.value = state.settings?.fiberSizes?.[key] || "";
      });
    } finally {
      syncingUi = false;
    }
  }

  function patchStore() {
    const store = S();
    if (!store || store.__compactAutoProposedPatched) return;
    if (typeof store.setState === "function") {
      const originalSetState = store.setState.bind(store);
      store.setState = function (nextState) {
        const result = originalSetState(nextState);
        applyAttachmentDefaults(store.getState?.());
        return result;
      };
    }
    if (typeof store.updateSetting === "function") {
      const originalUpdateSetting = store.updateSetting.bind(store);
      store.updateSetting = function (field, value) {
        const result = originalUpdateSetting(field, value);
        if (field === "projectProfile" || field === "proposedOwner") applyAttachmentDefaults(store.getState?.());
        return result;
      };
    }
    store.__compactAutoProposedPatched = true;
  }

  function installUiSync() {
    if (!global.document) return;
    const start = () => {
      syncFiberSettingsUi();
      const container = global.document.getElementById("clearanceSettings");
      if (container && global.MutationObserver) {
        const observer = new global.MutationObserver(() => {
          if (global.queueMicrotask) global.queueMicrotask(syncFiberSettingsUi);
          else global.setTimeout?.(syncFiberSettingsUi, 0);
        });
        observer.observe(container, { childList: true, subtree: true });
      }
      global.document.addEventListener("input", event => {
        if (event.target?.dataset?.scope === "attachmentSettings") markSaveDirty();
      });
      global.document.addEventListener("change", event => {
        if (["projectProfile", "proposedOwner"].includes(event.target?.dataset?.field) || event.target?.dataset?.scope === "attachmentSettings") {
          global.setTimeout?.(syncFiberSettingsUi, 0);
        }
      });
    };
    if (global.document.readyState === "loading") global.document.addEventListener("DOMContentLoaded", start);
    else start();
  }

  patchStore();
  patchMrLogic();
  applyAttachmentDefaults();
  if (global.ProjectExport) global.ProjectExport.exportProposedJson = exportCompactProposedJson;
  installUiSync();

  global.CompactAutoProposed = {
    applyAttachmentDefaults,
    detectedReferenceFiberCounts,
    fiberEntries,
    buildCompactPayload,
    validationErrors,
    compactMoves,
    compactSpansForPole,
    movementCandidates,
    augmentPoleMakeReady,
    exportCompactProposedJson,
    isPoleFullyUg,
    isPolePco,
    blocksLocalActions,
    sanitizeBlockedLocalPayload,
    sanitizeBlockedLocalMr,
    sanitizeFullyUgPayload: sanitizeBlockedLocalPayload,
    isUgSpan,
    spanKind,
    spanLengthInches,
    directionFromPole,
    directionTokensForReference
  };
})(window);
