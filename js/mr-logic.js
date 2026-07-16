(function (global) {
  "use strict";

  // MRLogic turns calculated state into the final Make Ready text shown per
  // pole. Ordering matters because crews read these instructions top to bottom.
  const H = () => global.HeightUtils;
  const S = () => global.AppStore;

  function detectRaiseLower(spanComm) {
    const existing = H().parseHeight(spanComm.existingHOA);
    const changed = H().parseHeight(spanComm.existingHOAChange);
    if (existing === null || changed === null || existing === changed) return null;
    return changed > existing ? "Raise" : "Lower";
  }

  function detectAttachFromSpanSide(spanSide) {
    return Boolean(spanSide && spanSide.proposedHOA);
  }

  function detectOverlash(spanComm) {
    return /overlash/i.test(`${spanComm.notes || ""} ${spanComm.mr || ""}`);
  }

  function detectAnchor(spanSide) {
    return /\banc\b|anchor/i.test(`${spanSide.notes || ""}`);
  }

  function detectRiser(spanSide) {
    return /riser/i.test(`${spanSide.notes || ""}`);
  }

  function detectDownGuy(value) {
    return /\bdg\b|down\s*guy/i.test(`${value || ""}`);
  }

  function detectOHG(spanSide) {
    return /\bohg\b|overhead\s*guy/i.test(`${spanSide.notes || ""}`);
  }

  function ownerForMR(spanComm) {
    const raw = spanComm.rawOwner || spanComm.ownerBase || spanComm.owner || "COMM";
    if (/century\s*link|centurylink|\bctl\b/i.test(raw)) return "CTL";
    return String(raw).replace(/^COMMUNICATION\s*>\s*/i, "").replace(/,\s*.*$/, "").trim() || "COMM";
  }

  function ownerForIntecMovementMR(spanComm) {
    const raw = spanComm.rawOwner || spanComm.ownerBase || spanComm.owner || "COMM";
    if (/century\s*link|centurylink|\bctl\b/i.test(raw)) return "CTL";
    return String(raw).replace(/^COMMUNICATION\s*>\s*/i, "").replace(/,\s*.*$/, "").trim() || "COMM";
  }

  function proposedOwnerForMR() {
    const settings = S().getState().settings || {};
    return String(settings.proposedOwner || "Wecom").trim() || "Wecom";
  }

  function mrHeight(value) {
    return String(value || "");
  }

  function applyCase(text) {
    const settings = S().getState().settings || {};
    return (settings.mrCase || "LOWER") === "UPPER" ? text.toUpperCase() : text;
  }

  function isMetronetMR() {
    const settings = S().getState().settings || {};
    return String(settings.mrTemplate || settings.projectProfile || "").toUpperCase() === "METRONET";
  }

  function shouldAddLowPowerMidspanMR() {
    const settings = S().getState().settings || {};
    return settings.allowLowPowerMidspanAdjustment !== false && !isMetronetMR();
  }

  function metronetAnchorMR(spanSide, direction) {
    const notes = String(spanSide?.notes || "");
    const hoa = mrHeight(spanSide?.proposedHOA || "");
    if (!hoa) return "";
    const dir = direction || "";
    if (detectOHG(spanSide)) {
      return `Pl new OHG${dir ? ` ${dir}` : ""} at HOA ${hoa} and pl DG on${dir ? ` ${dir}` : ""} pole.`;
    }
    const sizeMatch = notes.match(/\b(8|10)\s*(?:"|”|in\b|inch\b)?/i);
    const size = sizeMatch ? `${sizeMatch[1]}"` : `8"`;
    const sidewalk = /sidewalk/i.test(notes) ? " SIDEWALK" : "";
    const distanceMatch = notes.match(/\b(\d+)\s*'\s*([NSEW])\b/i);
    const distance = distanceMatch
      ? `${distanceMatch[1]}' ${distanceMatch[2].toUpperCase()}`
      : `15'${dir ? ` ${dir}` : ""}`;
    return `Pl new ${size}${sidewalk} ANC ${distance} and pl new DG at HOA ${hoa}.`;
  }

  function directionForSpanComm(spanComm) {
    const span = S().getSpan(spanComm?.spanId || "");
    if (!span) return "";
    return span.fromPole === spanComm.poleId
      ? (span.direction || "")
      : oppositeDirection(span.direction || "");
  }

  function generateResagServiceDropMR(spanComm) {
    if (!spanComm?.serviceDrop || !spanComm?.resagServiceDrop || isMetronetMR()) return "";
    const span = S().getSpan(spanComm.spanId || "");
    if (/back\s*span|backspan/i.test(`${span?.type || ""} ${span?.rawType || ""}`)) return "";
    const originalMidspan = H().parseHeight(spanComm.calculatedMidspan || spanComm.midspan || spanComm.ocalcMS || "");
    const target = H().parseHeight("15'6\"");
    if (originalMidspan === null || target === null || originalMidspan >= target) return "";
    const direction = directionForSpanComm(spanComm);
    return `Re-sag ${ownerForMR(spanComm)} comm drop${direction ? ` ${direction}` : ""}, ensure 15'6\" at midspan.`;
  }

  function commRowsInSamePoleGroup(spanComm) {
    const owner = ownerForMR(spanComm).toLowerCase();
    const existingHOA = H().parseHeight(spanComm?.existingHOA || "");
    return S().getSpanCommsForPole(spanComm?.poleId || "").filter(row => {
      const rowHOA = H().parseHeight(row?.existingHOA || "");
      return ownerForMR(row).toLowerCase() === owner && rowHOA === existingHOA;
    });
  }

  function commGroupTransferContext(spanComm) {
    const rows = commRowsInSamePoleGroup(spanComm);
    return {
      enabled: rows.some(row => row.transferToNewPole),
      downGuy: rows.some(row => row.downGuy || detectDownGuy(`${row.notes || ""} ${row.mr || ""}`))
    };
  }

  function generateTransferMRForPole(poleId) {
    const groups = new Map();
    S().getSpanCommsForPole(poleId).forEach(row => {
      const context = commGroupTransferContext(row);
      const height = H().parseHeight(row.existingHOAChange || row.existingHOA || "");
      if (!context.enabled || height === null) return;
      const owner = ownerForMR(row);
      const key = owner.toLowerCase();
      if (!groups.has(key)) groups.set(key, { owner, heights: new Set(), downGuy: false });
      const group = groups.get(key);
      group.heights.add(height);
      group.downGuy = group.downGuy || context.downGuy;
    });

    return Array.from(groups.values())
      .map(group => {
        const heights = Array.from(group.heights).sort((a, b) => a - b);
        return {
          minimum: heights[0] ?? Infinity,
          text: `Transfer ${group.owner} to new pole at HOA ${heights.map(value => H().formatHeight(value)).join(" and ")}${group.downGuy ? " with DG" : ""}.`
        };
      })
      .sort((a, b) => a.minimum - b.minimum)
      .map(item => item.text);
  }

  function generateMRForComm(spanComm) {
    if (!spanComm) return "";
    if (spanComm.mr && spanComm.mr.trim()) return spanComm.mr.trim();
    const resag = generateResagServiceDropMR(spanComm);
    const owner = ownerForMR(spanComm);
    const transferContext = commGroupTransferContext(spanComm);
    const dg = (transferContext.enabled ? transferContext.downGuy : (
      spanComm.downGuy || detectDownGuy(`${spanComm.notes || ""} ${spanComm.mr || ""}`)
    )) ? " with DG" : "";
    if (transferContext.enabled) {
      const transferHeight = spanComm.existingHOAChange || spanComm.existingHOA;
      if (!transferHeight) return resag;
      const transfer = `Transfer ${owner} to new pole at HOA ${mrHeight(transferHeight)}${dg}.`;
      return [transfer, resag].filter(Boolean).join("\n");
    }
    const action = detectRaiseLower(spanComm);
    if (!action) return resag;
    if (isMetronetMR()) {
      const verb = action === "Lower" ? "lower" : "raise";
      const movement = `At HOA ${mrHeight(spanComm.existingHOA)} ${verb} ${owner} to HOA ${mrHeight(spanComm.existingHOAChange)}${dg}.`;
      return [movement, resag].filter(Boolean).join("\n");
    }
    // Service drops use different MR wording than regular comm movement.
    const settings = S().getState().settings || {};
    if (spanComm.serviceDrop && settings.showServiceDrop !== false) {
      const relocation = `Relocate ${owner} drop at HOA ${mrHeight(spanComm.existingHOA)} to HOA ${mrHeight(spanComm.existingHOAChange)}.`;
      return [relocation, resag].filter(Boolean).join("\n");
    }
    const verb = action === "Lower" ? "lower" : "raise";
    const movement = `At HOA ${mrHeight(spanComm.existingHOA)} ${verb} ${ownerForIntecMovementMR(spanComm)} to HOA ${mrHeight(spanComm.existingHOAChange)}${dg}.`;
    return [movement, resag].filter(Boolean).join("\n");
  }

  function generateMRForSpanSide(spanSide) {
    if (!spanSide || !spanSide.proposedHOA) return "";
    const span = S().getSpan(spanSide.spanId);
    const dir = span && span.direction ? ` ${span.direction}` : "";
    const items = [];
    if (spanSide.clearanceMSReason === "LOW_POWER" && spanSide.clearanceMSIssue && shouldAddLowPowerMidspanMR()) {
      items.push(`Ensure min 30" to low power at midspan.`);
    }
    if (isMetronetMR()) {
      if (detectAnchor(spanSide) || detectDownGuy(spanSide.notes) || detectOHG(spanSide)) {
        const anchor = metronetAnchorMR(spanSide, span?.direction || "");
        if (anchor) items.push(anchor);
      }
      if (detectRiser(spanSide)) items.push(`Pl new riser${dir}.`.replace("  ", " "));
      return items.join("\n");
    }
    // Slack is selected in the PLA model, not inferred by the calculator from
    // free-form notes. Excel Review accepts that model-owned instruction.
    if (detectAnchor(spanSide)) items.push(`PL NEW ANC${dir}.`.replace("  ", " "));
    if (detectRiser(spanSide)) items.push(`PL NEW RISER${dir}.`.replace("  ", " "));
    return items.join("\n");
  }

  function generateAttachMRForPole(poleId) {
    if (isMetronetMR()) return "";
    const heights = S().getSpanSidesForPole(poleId)
      .map(side => H().parseHeight(side.proposedHOA || ""))
      .concat(H().parseHeight(S().getPole(poleId)?.standaloneProposedHOA || ""))
      .filter(value => value !== null)
      .sort((a, b) => a - b);
    const uniqueHeights = Array.from(new Set(heights)).map(value => mrHeight(H().formatHeight(value)));
    if (!uniqueHeights.length) return "";
    const joinedHeights = uniqueHeights.join(" and ");
    return `Attach ${proposedOwnerForMR()} at HOA ${joinedHeights}.`;
  }

  function ugReplacementMR() {
    if (isMetronetMR()) {
      return [
        "Suggest going UG due to [clearance violation / 2 span aerial requirement]."
      ];
    }
    return ["Unable to attach due to proposed pole overloaded."];
  }

  function pcoReplacementMR() {
    if (isMetronetMR()) {
      return [
        "Replace pole to 45ft Class 3 due to (failing clearances on pole / failing clearances at midspan / failing load capacity) & (keeping 1/2 span aerial / keeping 2 poles aerial)."
      ];
    }
    return [
      "(Existing/proposed) clearance violations (specify violation), replace pole. Transfer (existing comm 1 ie Fiber, CATV, Telco) & (existing comm 2 type) to new pole.",
      "(Existing/proposed) pole overloaded by (who is causing overload), replace pole. Transfer (existing comm 1 ie Fiber, CATV, Telco) & (existing comm 2 type) to new pole."
    ];
  }

  function oppositeDirection(direction) {
    return {
      N: "S",
      NE: "SW",
      E: "W",
      SE: "NW",
      S: "N",
      SW: "NE",
      W: "E",
      NW: "SE"
    }[direction] || direction || "";
  }

  function connectedUGInstructions(poleId) {
    const spans = S().getConnectedSpans(poleId);
    const byOtherPole = new Map();

    function normalizeSpanRelation(rawType, isReverseSide) {
      const clean = String(rawType || "").replace(/\s+/g, "").toLowerCase();
      const relation = clean === "forespan" ? "Forespan"
        : clean === "backspan" ? "Backspan"
          : "Otherspan";
      if (!isReverseSide) return relation;
      if (relation === "Forespan") return "Backspan";
      if (relation === "Backspan") return "Forespan";
      return relation;
    }

    spans.forEach(span => {
      const otherPoleId = S().getOtherPoleId(span, poleId);
      const otherPole = S().getPole(otherPoleId);
      if (!otherPole?.ugActive || !otherPoleId) return;
      const current = byOtherPole.get(otherPoleId);
      // Excel already defines whether the span is Fore, Back, or Other from
      // the imported side. Reverse it only when we are reading the far end.
      const relation = normalizeSpanRelation(span.rawType || span.type, span.fromPole !== poleId);
      const directionFromThisPole = span.fromPole === poleId
        ? span.direction
        : oppositeDirection(span.direction);
      const candidate = { relation, direction: directionFromThisPole || "", spanId: span.spanId, otherPoleId };
      // A physical connection can appear twice in Span. Prefer the Back Span
      // row because it is the relation owned by this pole for the UG handoff.
      if (!current || (current.relation !== "Backspan" && relation === "Backspan")) {
        byOtherPole.set(otherPoleId, candidate);
      }
    });

    return Array.from(byOtherPole.values()).flatMap(item => {
      const direction = item.direction ? ` ${item.direction}` : "";
      if (isMetronetMR()) {
        const relation = item.relation === "Otherspan" ? "Other Span" : item.relation;
        return [`${relation} going UG due to [clearance violation/insert other reason]. Pl new ANC/DG for deadending lines. Pl new riser for UG transfer${direction}.`];
      }
      const lines = [`${item.relation} to go UG${direction} due to existing pole overloaded.`];
      if (item.relation !== "Backspan") return lines;

      const proposedSides = S().getSpanSidesForPole(poleId)
        .filter(side => side.proposedHOA)
        .sort((a, b) => Number(Boolean(a.isAdditionalProposed)) - Number(Boolean(b.isAdditionalProposed)));
      const proposedSide = proposedSides.find(side => {
        if (!side.proposedHOA) return false;
        if (side.spanId === item.spanId) return true;
        const sideSpan = S().getSpan(side.spanId);
        return Boolean(sideSpan && [sideSpan.fromPole, sideSpan.toPole].includes(item.otherPoleId));
      }) || proposedSides[0];
      const proposed = H().parseHeight(proposedSide?.proposedHOA || S().getPole(poleId)?.standaloneProposedHOA || "");
      if (proposed !== null) {
        // The riser uses a pole-face cardinal. Keep the full octant on the UG
        // line, but use the first component for diagonals (SE -> S, NW -> N).
        const riserDirection = (item.direction || "").charAt(0).toUpperCase();
        lines.push(`Pl riser${riserDirection ? ` ${riserDirection}` : ""} at HOA ${H().formatHeight(proposed - 12)}.`);
      }
      return lines;
    });
  }

  function generateMRForSpan(spanId) {
    const sideItems = S().getSpanSidesForSpan(spanId).map(generateMRForSpanSide).filter(Boolean);
    const commItems = S().getSpanCommsForSpan(spanId)
      .map(generateMRForComm)
      .flatMap(text => String(text || "").split(/\n+/).filter(Boolean));
    return [...sideItems, ...commItems];
  }

  /**
   * Replaces the generated Make Ready block for one pole using current state.
   * @param {string} poleId
   * @returns {Array<Object>} Generated MR records for the pole.
   */
  function generateMRForPole(poleId) {
    const state = S().getState();
    state.mr = state.mr.filter(item => item.poleId !== poleId);

    const ug = [];
    const power = [];
    const commMoves = [];
    const dropMoves = [];
    const proposed = [];
    const ensure = [];
    const pole = S().getPole(poleId);
    if (pole?.ugActive || pole?.pcoActive) {
      const lines = pole.ugActive ? ugReplacementMR() : pcoReplacementMR();
      const text = lines.map(applyCase).join("\n");
      state.mr.push({ poleId, spanId: "", owner: "MR", text, imported: false });
      return state.mr.filter(item => item.poleId === poleId);
    }
    ug.push(...connectedUGInstructions(poleId));
    commMoves.push(...generateTransferMRForPole(poleId));
    S().getSpanSidesForPole(poleId).forEach(side => {
      const text = generateMRForSpanSide(side);
      if (text) text.split(/\n+/).filter(Boolean).forEach(line => {
        if (/ensure min 30/i.test(line)) ensure.push(line);
        else proposed.push(line);
      });
    });
    S().getSpanCommsForPole(poleId)
      .slice()
      .sort((a, b) => (H().parseHeight(getEffectiveCommHOAForMR(b)) ?? -Infinity) - (H().parseHeight(getEffectiveCommHOAForMR(a)) ?? -Infinity))
      .forEach(sc => {
      if (commGroupTransferContext(sc).enabled) {
        const resag = generateResagServiceDropMR(sc);
        if (resag) dropMoves.push(resag);
        return;
      }
      const text = generateMRForComm(sc);
      if (!text) return;
      text.split(/\n+/).filter(Boolean).forEach(line => {
        if (/^Re-sag\b/i.test(line) || (sc.serviceDrop && !sc.transferToNewPole)) dropMoves.push(line);
        else commMoves.push(line);
      });
    });
    const attach = generateAttachMRForPole(poleId);
    if (attach) proposed.unshift(attach);

    const lines = [...ug, ...power, ...commMoves, ...dropMoves, ...proposed, ...ensure].map(applyCase);
    const unique = Array.from(new Set(lines.map(line => line.trim()).filter(Boolean)));
    if (unique.length) state.mr.push({ poleId, spanId: "", owner: "MR", text: unique.join("\n"), imported: false });
    return state.mr.filter(item => item.poleId === poleId);
  }

  /** @returns {Array<Object>} Fresh generated Make Ready records for all poles. */
  function generateAllMR() {
    const state = S().getState();
    state.mr = [];
    Object.keys(state.poles).forEach(generateMRForPole);
    return state.mr;
  }

  function getEffectiveCommHOAForMR(spanComm) {
    return spanComm?.existingHOAChange || spanComm?.existingHOA || "";
  }

  /** @namespace MRLogic */
  global.MRLogic = {
    generateMRForPole,
    generateMRForSpan,
    generateMRForComm,
    generateResagServiceDropMR,
    generateMRForSpanSide,
    generateAllMR,
    detectAttach: detectAttachFromSpanSide,
    detectRaiseLower,
    detectOverlash,
    detectAnchor,
    detectRiser
  };
})(window);
