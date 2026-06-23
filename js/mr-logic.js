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

  function detectSlack(spanSide) {
    return /slack/i.test(`${spanSide.notes || ""}`);
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

  function generateMRForComm(spanComm) {
    if (!spanComm) return "";
    if (spanComm.mr && spanComm.mr.trim()) return spanComm.mr.trim();
    const action = detectRaiseLower(spanComm);
    if (!action) return "";
    const owner = ownerForMR(spanComm);
    if (isMetronetMR()) {
      const verb = action === "Lower" ? "lower" : "raise";
      const dg = spanComm.downGuy || detectDownGuy(`${spanComm.notes || ""} ${spanComm.mr || ""}`) ? " with DG" : "";
      return `At HOA ${mrHeight(spanComm.existingHOA)} ${verb} ${owner} to HOA ${mrHeight(spanComm.existingHOAChange)}${dg}.`;
    }
    // Service drops use different MR wording than regular comm movement.
    const settings = S().getState().settings || {};
    if (spanComm.serviceDrop && settings.showServiceDrop !== false) return `Relocate ${owner} drop at HOA ${mrHeight(spanComm.existingHOA)} to HOA ${mrHeight(spanComm.existingHOAChange)}.`;
    const verb = action === "Lower" ? "lower" : "raise";
    const dg = spanComm.downGuy || detectDownGuy(`${spanComm.notes || ""} ${spanComm.mr || ""}`) ? " with DG" : "";
    return `At HOA ${mrHeight(spanComm.existingHOA)} ${verb} ${ownerForIntecMovementMR(spanComm)} to HOA ${mrHeight(spanComm.existingHOAChange)}${dg}.`;
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
    if (detectSlack(spanSide)) items.push(`Proposed slack span${dir}.`.replace("  ", " "));
    if (detectAnchor(spanSide)) items.push(`PL NEW ANC${dir}.`.replace("  ", " "));
    if (detectRiser(spanSide)) items.push(`PL NEW RISER${dir}.`.replace("  ", " "));
    return items.join("\n");
  }

  function generateAttachMRForPole(poleId) {
    if (isMetronetMR()) return "";
    const heights = S().getSpanSidesForPole(poleId)
      .map(side => H().parseHeight(side.proposedHOA || ""))
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
    return [
      "Unable to attach due to (reasoning).",
      "Red tag",
      "Inability to place ANC",
      "TDU replace required",
      "Existing neutral / multiplex above 26'9\"",
      "PCO neutral / multiplex exceeds 26'9\""
    ];
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
      const candidate = { relation, direction: directionFromThisPole || "" };
      if (!current || (current.relation === "Backspan" && relation !== "Backspan")) {
        byOtherPole.set(otherPoleId, candidate);
      }
    });

    return Array.from(byOtherPole.values()).map(item => {
      const direction = item.direction ? ` ${item.direction}` : "";
      if (isMetronetMR()) {
        const relation = item.relation === "Otherspan" ? "Other Span" : item.relation;
        return `${relation} going UG due to [clearance violation/insert other reason]. Pl new ANC/DG for deadending lines. Pl new riser for UG transfer${direction}.`;
      }
      return `${item.relation} to go UG${direction} due to (Red tag / TDU Replace required / PCO neutral (inclusive of triplex/quadruplexes as noted above) exceeds 26'9" / inability to place ANC).`;
    });
  }

  function generateMRForSpan(spanId) {
    const sideItems = S().getSpanSidesForSpan(spanId).map(generateMRForSpanSide).filter(Boolean);
    const commItems = S().getSpanCommsForSpan(spanId).map(generateMRForComm).filter(Boolean);
    return [...sideItems, ...commItems];
  }

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
      const text = generateMRForComm(sc);
      if (!text) return;
      if (sc.serviceDrop) dropMoves.push(text);
      else commMoves.push(text);
    });
    const attach = generateAttachMRForPole(poleId);
    if (attach) proposed.unshift(attach);

    const lines = [...ug, ...power, ...commMoves, ...dropMoves, ...proposed, ...ensure].map(applyCase);
    const unique = Array.from(new Set(lines.map(line => line.trim()).filter(Boolean)));
    if (unique.length) state.mr.push({ poleId, spanId: "", owner: "MR", text: unique.join("\n"), imported: false });
    return state.mr.filter(item => item.poleId === poleId);
  }

  function generateAllMR() {
    const state = S().getState();
    state.mr = [];
    Object.keys(state.poles).forEach(generateMRForPole);
    return state.mr;
  }

  function getEffectiveCommHOAForMR(spanComm) {
    return spanComm?.existingHOAChange || spanComm?.existingHOA || "";
  }

  global.MRLogic = {
    generateMRForPole,
    generateMRForSpan,
    generateMRForComm,
    generateMRForSpanSide,
    generateAllMR,
    detectAttach: detectAttachFromSpanSide,
    detectRaiseLower,
    detectOverlash,
    detectSlack,
    detectAnchor,
    detectRiser
  };
})(window);
