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

  // Make Ready height lists follow normal English punctuation while keeping
  // the caller's already-sorted order: A, A and B, or A, B and C.
  function joinMRList(values) {
    const items = (values || []).filter(Boolean);
    if (items.length < 2) return items[0] || "";
    if (items.length === 2) return `${items[0]} and ${items[1]}`;
    return `${items.slice(0, -1).join(", ")} and ${items.at(-1)}`;
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
          highest: heights[heights.length - 1] ?? -Infinity,
          lowest: heights[0] ?? -Infinity,
          text: `Transfer ${group.owner} to new pole at HOA ${joinMRList(heights.map(value => H().formatHeight(value)))}${group.downGuy ? " with DG" : ""}.`
        };
      })
      // Crew instructions follow the physical pole stack from top to bottom.
      // A grouped owner is positioned by its highest effective attachment.
      .sort((a, b) => (b.highest - a.highest) || (b.lowest - a.lowest))
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
    const joinedHeights = joinMRList(uniqueHeights);
    return `Attach ${proposedOwnerForMR()} at HOA ${joinedHeights}.`;
  }

  function cleanUGReason(value) {
    return String(value || "").trim().replace(/[.]+$/, "") || "(reasoning)";
  }

  function defaultIntecUGLines(pole) {
    return [
      `Unable to attach due to ${cleanUGReason(pole?.ugReason)}.`,
      "Red tag",
      "Inability to place ANC",
      "TDU replace required",
      "Existing neutral / multiplex above 26'9\"",
      "PCO neutral / multiplex exceeds 26'9\""
    ];
  }

  function editableUGTemplate(pole) {
    return String(pole?.ugMRText || "").trim() || defaultIntecUGLines(pole).join("\n");
  }

  function ugReasonFromPole(pole) {
    const firstLine = editableUGTemplate(pole).split(/\r?\n/).find(line => /unable\s+to\s+attach\s+due\s+to/i.test(line));
    const match = String(firstLine || "").match(/unable\s+to\s+attach\s+due\s+to\s+(.+?)[.]*\s*$/i);
    return cleanUGReason(match?.[1] || pole?.ugReason);
  }

  function ugReplacementMR(pole) {
    if (isMetronetMR()) {
      return [
        "Suggest going UG due to [clearance violation / 2 span aerial requirement]."
      ];
    }
    return editableUGTemplate(pole).split(/\r?\n/).map(line => line.trim()).filter(Boolean);
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

  // Riser direction starts with the model-owned Make Ready/IO value. The pole
  // can store a user override. When neither exists, the imported span relation
  // supplies a practical default. Both Fore Span and Back Span use the UG
  // span direction from the current pole. Other spans stay blank for review.
  function importedRiserDirection(poleId) {
    const refs = (S().getState().makeReadyReferences || []).filter(ref => ref.poleId === poleId);
    const directionPattern = "(NE|NW|SE|SW|N|E|S|W)";
    for (const ref of refs) {
      const sources = [
        ref.makeReadyNotes,
        ...Object.values(ref.raw || {}).filter(value => typeof value === "string")
      ];
      for (const source of sources) {
        const text = String(source || "");
        const match = text.match(new RegExp(`\\b(?:pl(?:ace)?\\s+(?:new\\s+)?riser(?:\\s+for\\s+ug\\s+transfer)?|riser)\\s+${directionPattern}\\b`, "i"));
        if (match) return match[1].toUpperCase();
      }
    }
    return "";
  }

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

  function defaultRiserDirection(relation, direction) {
    const normalizedDirection = String(direction || "").toUpperCase();
    if (!normalizedDirection) return "";
    if (relation === "Forespan" || relation === "Backspan") return normalizedDirection;
    return "";
  }

  // Builds one canonical UG connection per adjacent pole. Excel can contain
  // both directions of the same physical span, so Back Span wins when both
  // rows describe the same adjacent UG pole.
  function connectedUGSpanItems(poleId) {
    const spans = S().getConnectedSpans(poleId);
    const byOtherPole = new Map();

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

    return Array.from(byOtherPole.values());
  }

  function resolvedRiserDirection(poleId, connection = null) {
    const pole = S().getPole(poleId);
    const saved = String(pole?.ugRiserDirection || "").toUpperCase();
    if (saved) return saved;
    const imported = importedRiserDirection(poleId);
    if (imported) return imported;
    const item = connection || riserConnectionForPole(poleId);
    return item ? defaultRiserDirection(item.relation, item.direction) : "";
  }

  function riserConnectionForPole(poleId) {
    const items = connectedUGSpanItems(poleId);
    const adjacentUG = items.find(item => item.relation === "Backspan")
      || items.find(item => item.relation === "Forespan")
      || items[0]
      || null;
    if (adjacentUG) return adjacentUG;

    // A manually selected Riser does not require an adjacent UG pole. Use the
    // primary Proposed span to supply its default direction when available.
    const side = S().getSpanSidesForPole(poleId)
      .filter(item => item.proposedHOA)
      .sort((a, b) => Number(Boolean(a.isAdditionalProposed)) - Number(Boolean(b.isAdditionalProposed)))[0];
    const span = S().getSpan(side?.spanId || "");
    if (!span) return null;
    return {
      relation: normalizeSpanRelation(span.rawType || span.type, span.fromPole !== poleId),
      direction: span.fromPole === poleId ? (span.direction || "") : oppositeDirection(span.direction || ""),
      spanId: span.spanId,
      otherPoleId: S().getOtherPoleId(span, poleId) || ""
    };
  }

  function isRiserAvailable(poleId) {
    const pole = S().getPole(poleId);
    return Boolean(pole && !pole.ugActive && !pole.pcoActive);
  }

  // null preserves the legacy automatic Back Span behavior. true/false is an
  // explicit user override from the Riser action button.
  function isRiserEnabled(poleId, connection = null) {
    const pole = S().getPole(poleId);
    if (pole?.riserActive === true || pole?.riserActive === false) return pole.riserActive;
    // Only an actual adjacent Back Span UG turns the legacy automatic riser
    // on. A normal Back Span or Proposed does not activate it by itself.
    const items = connectedUGSpanItems(poleId);
    const item = items.find(candidate => candidate.relation === "Backspan")
      || items.find(candidate => candidate.relation === "Forespan")
      || (connection && items.find(candidate => candidate.spanId === connection.spanId));
    return item?.relation === "Backspan" || item?.relation === "Forespan";
  }

  function generateRiserInstruction(poleId) {
    if (isMetronetMR() || !isRiserAvailable(poleId) || !isRiserEnabled(poleId)) return "";
    const connection = riserConnectionForPole(poleId);
    const proposedSides = S().getSpanSidesForPole(poleId)
      .filter(side => side.proposedHOA)
      .sort((a, b) => Number(Boolean(a.isAdditionalProposed)) - Number(Boolean(b.isAdditionalProposed)));
    const proposedSide = proposedSides.find(side => {
      if (!connection) return false;
      if (side.spanId === connection.spanId) return true;
      const sideSpan = S().getSpan(side.spanId);
      return Boolean(sideSpan && [sideSpan.fromPole, sideSpan.toPole].includes(connection.otherPoleId));
    }) || proposedSides[0];
    const proposed = H().parseHeight(proposedSide?.proposedHOA || S().getPole(poleId)?.standaloneProposedHOA || "");
    if (proposed === null) return "";
    const riserDirection = resolvedRiserDirection(poleId, connection);
    const direction = riserDirection ? ` ${riserDirection}` : "";
    return `Pl riser${direction} at HOA ${H().formatHeight(proposed - 12)}.`;
  }

  function connectedUGInstructions(poleId) {
    const items = connectedUGSpanItems(poleId);

    const lines = items.flatMap(item => {
      const direction = item.direction ? ` ${item.direction}` : "";
      if (isMetronetMR()) {
        const relation = item.relation === "Otherspan" ? "Other Span" : item.relation;
        return [`${relation} going UG due to [clearance violation/insert other reason]. Pl new ANC/DG for deadending lines. Pl new riser for UG transfer${direction}.`];
      }
      const adjacentReason = ugReasonFromPole(S().getPole(item.otherPoleId));
      return [`${item.relation} to go UG${direction} due to on adj pole ${adjacentReason}.`];
    });

    const riser = generateRiserInstruction(poleId);
    if (riser) lines.push(riser);
    return lines;
  }

  function generateMRForSpan(spanId) {
    const sideItems = S().getSpanSidesForSpan(spanId).map(generateMRForSpanSide).filter(Boolean);
    const commItems = S().getSpanCommsForSpan(spanId)
      .map(generateMRForComm)
      .flatMap(text => String(text || "").split(/\n+/).filter(Boolean));
    return [...sideItems, ...commItems];
  }

  /**
   * Converts user-selected Power Equipment actions into crew instructions.
   * Imported heights remain the source HOA; actionHeight is the destination.
   */
  function generatePowerEquipmentMRForPole(poleId) {
    const rows = S().getPole(poleId)?.metadata?.powerEquipment || [];
    return rows.flatMap(row => {
      const category = String(row.category || row.type || "").toUpperCase();
      if (category.includes("STREETLIGHT")) {
        const lines = [];
        if (row.actionActive) {
          lines.push(isMetronetMR()
            ? "MNT GROUND STREETLIGHT"
            : "Install flex conduit to STLT circuit. bond STLT housing to pole GRND/NEUT.");
        }
        if (!isMetronetMR() && row.raiseActive) {
          const source = H().parseHeight(row.attachmentHeight || "");
          const target = H().parseHeight(row.raiseHeight || "");
          if (source !== null && target !== null && target > source && target <= source + 12) {
            lines.push(`Raise streetlight from HOA ${H().formatHeight(source)} to ${H().formatHeight(target)}.`);
          }
        }
        return lines;
      }

      if (!row.actionActive) return [];

      const target = H().parseHeight(row.actionHeight || "");
      if (target === null) return [];
      const targetText = H().formatHeight(target);
      if (category.includes("TRANSFORMER")) {
        return [isMetronetMR()
          ? `POWER REDRESS TRANSFORMER DRIP LOOP TO HOA ${targetText}.`
          : `Secure transformer drip loop to HOA ${targetText}.`];
      }
      if (category.includes("RISER")) {
        const source = H().parseHeight(row.attachmentHeight || "");
        if (source === null || target <= source) return [];
        return [isMetronetMR()
          ? `AT HOA ${H().formatHeight(source)} RAISE POWER RISER TO HOA ${targetText} DUE TO CLEARANCES.`
          : `Raise APS riser from HOA ${H().formatHeight(source)} to HOA ${targetText}.`];
      }
      return [];
    });
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
    const risers = [];
    const pole = S().getPole(poleId);
    if (pole?.ugActive || pole?.pcoActive) {
      const lines = pole.ugActive ? ugReplacementMR(pole) : pcoReplacementMR();
      const text = lines.map(applyCase).join("\n");
      state.mr.push({ poleId, spanId: "", owner: "MR", text, imported: false });
      return state.mr.filter(item => item.poleId === poleId);
    }
    connectedUGInstructions(poleId).forEach(line => {
      // Riser work is always the final instruction crews read in the pole MR.
      // Keep the UG relation at the top, but defer its separate riser line.
      if (/^pl\s+(?:new\s+)?riser\b/i.test(line)) risers.push(line);
      else ug.push(line);
    });
    power.push(...generatePowerEquipmentMRForPole(poleId));
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

    const lines = [...ug, ...power, ...commMoves, ...dropMoves, ...proposed, ...ensure, ...risers].map(applyCase);
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
    generatePowerEquipmentMRForPole,
    getEditableUGTemplate: editableUGTemplate,
    getImportedRiserDirection: importedRiserDirection,
    getResolvedRiserDirection: resolvedRiserDirection,
    getDefaultRiserDirection: defaultRiserDirection,
    isRiserAvailable,
    isRiserEnabled,
    generateRiserInstruction,
    generateMRForSpanSide,
    generateAllMR,
    detectAttach: detectAttachFromSpanSide,
    detectRaiseLower,
    detectOverlash,
    detectAnchor,
    detectRiser
  };
})(window);
