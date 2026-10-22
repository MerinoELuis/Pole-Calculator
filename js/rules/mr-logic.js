(function (global) {
  "use strict";

  // MRLogic turns calculated state into the final Make Ready text shown per
  // pole. Ordering matters because crews read these instructions top to bottom.
  const H = () => global.HeightUtils;
  const S = () => global.AppStore;

  const poleResolutionOf = pole => S().getPoleResolution(pole);
  const ownerCanonicalOf = row => S().getOwnerCanonical(row);
  const ownerImportedOf = row => S().getOwnerImported(row);
  const spanRelationOf = span => S().getSpanRelation(span);

  function detectRaiseLower(spanComm) {
    const existing = H().parseHeight(spanComm.existingHOA);
    const changed = H().parseHeight(spanComm.otherHOA || spanComm.existingHOAChange);
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

  function generateProposedSlackMR(spanSide, span) {
    if (!spanSide || !span) return "";
    const spanIsSlack = Boolean(spanSide.isSlack)
      || S().getSpanSidesForSpan(span.spanId).some(side => Boolean(side.isSlack));
    if (!spanIsSlack) return "";
    const direction = directionForPhysicalSpan(span, spanSide.poleId);
    if (!direction) return "";
    if (isMidAmMR()) return `PL NEW SLACK SPAN ${direction}.`;
    // CSU and ComEd do not have a confirmed Slack MR template in the current
    // project rules, so do not invent MidAm wording for those profiles.
    if (isMetronetMR()) return "";
    return isOlssonMR()
      ? `PROPOSED SLACK SPAN ${direction}.`
      : `Proposed slack span ${direction}.`;
  }

  function directionForPhysicalSpan(span, poleId) {
    const direction = String(span?.direction || "").trim().toUpperCase();
    if (!direction) return "";
    if (span?.fromPole === poleId) return direction;
    if (span?.toPole === poleId) return oppositeDirection(direction);
    return direction;
  }

  function getMROwnerLabel(spanComm) {
    const raw = ownerImportedOf(spanComm) || ownerCanonicalOf(spanComm) || spanComm.ownerDisplay || "COMM";
    if (/century\s*link|centurylink|\bctl\b/i.test(raw)) return "CTL";
    return displayOwnerName(raw);
  }

  function getIntecMovementOwnerLabel(spanComm) {
    const raw = ownerImportedOf(spanComm) || ownerCanonicalOf(spanComm) || spanComm.ownerDisplay || "COMM";
    if (/century\s*link|centurylink|\bctl\b/i.test(raw)) return "CTL";
    return displayOwnerName(raw);
  }

  function displayOwnerName(value) {
    return String(value || "COMM")
      .replace(/^COMMUNICATION\s*>\s*/i, "")
      .replace(/,\s*.*$/, "")
      .replace(/\s+communications?$/i, "")
      .replace(/\s+inc\.?$/i, "")
      .trim() || "COMM";
  }

  function proposedOwnerForMR() {
    const settings = S().getState().settings || {};
    const owner = String(settings.proposedOwner || "Wecom").trim() || "Wecom";
    return /century\s*link|centurylink|\bctl\b/i.test(owner) ? "CTL" : displayOwnerName(owner);
  }

  function mrHeight(value) {
    return String(value || "");
  }

  function formattedMRHeight(value) {
    const parsed = H().parseHeight(value);
    return parsed === null ? mrHeight(value) : H().formatHeight(parsed);
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

  function normalizeMRLine(value) {
    return String(value || "")
      .replace(/[’‘]/g, "'")
      .replace(/[“”]/g, '"')
      .replace(/(\d+)'(\d+)(?:")?(?=\s|[.,]|$)/g, '$1\'$2"')
      .replace(/\s+/g, " ")
      .trim()
      .toLowerCase();
  }

  function isMetronetMR() {
    const settings = S().getState().settings || {};
    return String(settings.mrTemplate || settings.projectProfile || "").toUpperCase() === "METRONET";
  }

  function isMidAmMR() {
    const settings = S().getState().settings || {};
    return isMetronetMR() && String(settings.proposedOwner || "").toUpperCase() === "MIDAM";
  }

  function isOlssonMR() {
    const settings = S().getState().settings || {};
    // The selected project profile is authoritative. Older saved states can
    // retain a stale mrTemplate from a previous project, which must not make
    // an Olsson Service Drop fall back to the Relocate template.
    if (String(settings.projectProfile || "").toUpperCase() === "OLSSON_OPPD") return true;
    return String(settings.mrTemplate || "").toUpperCase() === "OLSSON_OPPD";
  }

  function isIntecMR() {
    return !isMetronetMR() && !isOlssonMR();
  }

  function generateRedTagMR(pole) {
    if (!isOlssonMR() || !pole?.redTagActive) return "";
    const minimumHeight = formattedMRHeight(pole.redTagMinHeight || "");
    const minimumClass = String(pole.redTagClass || "").trim();
    const minimum = minimumHeight && minimumClass
      ? `${minimumHeight}-${minimumClass}`
      : (minimumHeight || minimumClass);
    const reason = String(pole.redTagReason || "").trim().replace(/[.]+$/, "");
    if (!minimum) return "RED TAG PRESENT - POLE IS MARKED FOR OPPD REPLACEMENT";
    return `RED TAG PRESENT - POLE IS MARKED FOR OPPD REPLACEMENT - MIN ${minimum} DUE TO ${reason || "CLEARANCES"}.`;
  }

  function generateOlssonCommClearanceNote(pole) {
    if (!isOlssonMR() || !pole?.lessThan12CommClearanceActive) return "";
    return `LESS THAN 12\" COMM CLEARANCE REQUESTED.`;
  }

  function shouldAddLowPowerMidspanMR() {
    const settings = S().getState().settings || {};
    return settings.allowLowPowerMidspanAdjustment !== false && !isMetronetMR();
  }

  function generateAnchorInstructionFromSpan(spanSide, direction) {
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

  function getPoleSpanDirection(poleId) {
    const sides = S().getSpanSidesForPole(poleId);
    const side = sides.find(item => item.proposedHOA) || sides[0];
    const span = S().getSpan(side?.spanId || "");
    if (!side || !span) return "";
    return span.fromPole === poleId ? String(span.direction || "").toUpperCase() : oppositeDirection(span.direction || "");
  }

  function normalizeAnchorSize(value) {
    const raw = String(value || "8").trim().replace(/[“”″]/g, '"');
    return /"$/.test(raw) ? raw : `${raw}"`;
  }

  function normalizeAnchorDistance(value) {
    const parsed = H().parseHeight(value || "15'");
    return parsed === null ? String(value || "15'").trim() : H().formatHeight(parsed);
  }

  function proposedHoaForAnchor(pole) {
    if (!pole) return "";
    if (pole.standaloneProposedHOA) return String(pole.standaloneProposedHOA).trim();
    return S().getSpanSidesForPole(pole.poleId).find(item => item.proposedHOA)?.proposedHOA || "";
  }

  function defaultAnchorDistance(pole) {
    const proposed = H().parseHeight(proposedHoaForAnchor(pole));
    if (proposed === null) return "";
    // The automatic anchor distance is one-third of Proposed, rounded up to
    // the next whole foot so the field never understates the required reach.
    return H().formatHeight(Math.ceil((proposed / 3) / 12) * 12);
  }

  function getEffectiveAnchorDistance(pole) {
    const manual = String(pole?.ancDistance || "").trim();
    if (pole?.ancDistanceManual === true && manual) return normalizeAnchorDistance(manual);
    return defaultAnchorDistance(pole) || normalizeAnchorDistance(manual || "15'");
  }

  function automaticAnchorDirection(pole) {
    const ugConnection = connectedUGSpanItems(pole?.poleId)
      .find(item => item.adjacentPoleUG);
    return String(ugConnection?.direction || getPoleSpanDirection(pole?.poleId) || "").toUpperCase();
  }

  function getEffectiveAnchorDirection(pole) {
    const manual = String(pole?.ancDirection || "").trim().toUpperCase();
    if (pole?.ancDirectionManual === true && manual) return manual;
    // An ANC that the user enabled manually must wait for a direction choice.
    // Only ANC inferred from an adjacent UG connection receives an automatic
    // direction.
    if (pole?.ancActive === true) return "";
    return automaticAnchorDirection(pole);
  }

  function getEffectiveAnchorDGHOA(pole) {
    if (!isANCEnabled(pole?.poleId)) return "";
    const manual = String(pole?.dgHoa || "").trim();
    return formattedMRHeight(manual || proposedHoaForAnchor(pole));
  }

  function generateConfiguredAnchorInstruction(pole) {
    if (!isANCEnabled(pole?.poleId)) return "";
    const direction = getEffectiveAnchorDirection(pole);
    const distance = getEffectiveAnchorDistance(pole);
    const proposedHoa = proposedHoaForAnchor(pole);
    const dgHoa = getEffectiveAnchorDGHOA(pole) || formattedMRHeight(proposedHoa);
    if (!direction || !dgHoa) return "";
    if (isIntecMR()) {
      return `Pl new ANC ${distance} ${direction} and pl new DG at HOA ${dgHoa}.`;
    }
    return `PL NEW ${normalizeAnchorSize(pole.ancSize)} ANC ${distance} ${direction} AND PL NEW DG AT HOA ${dgHoa}.`;
  }

  function metronetConfiguredOHGMR(pole) {
    if (!pole?.ohgActive) return "";
    const direction = String(pole.ohgDirection || getPoleSpanDirection(pole.poleId) || "").toUpperCase();
    const hoa = formattedMRHeight(pole.ohgHoa || pole.standaloneProposedHOA || (
      S().getSpanSidesForPole(pole.poleId).find(item => item.proposedHOA)?.proposedHOA || ""
    ));
    if (!direction || !hoa) return "";
    return `PL NEW OHG ${direction} AT HOA ${hoa} AND PL DG ON ${direction} POLE.`;
  }

  function directionForSpanComm(spanComm) {
    const span = S().getSpan(spanComm?.spanId || "");
    if (!span) return "";
    return span.fromPole === spanComm.poleId
      ? (span.direction || "")
      : oppositeDirection(span.direction || "");
  }

  function generateResagServiceDropMR(spanComm) {
    const settings = S().getState().settings || {};
    if (!spanComm?.serviceDrop || !spanComm?.resagServiceDrop || settings.showResagServiceDrop === false || isMetronetMR()) return "";
    const span = S().getSpan(spanComm.spanId || "");
    if (/back\s*span|backspan/i.test(spanRelationOf(span))) return "";
    const originalMidspan = H().parseHeight(spanComm.calculatedMidspan || spanComm.midspan || spanComm.ocalcMS || "");
    const target = H().parseHeight("15'6\"");
    if (originalMidspan === null || target === null || originalMidspan >= target) return "";
    const direction = directionForSpanComm(spanComm);
    return `Re-sag ${getMROwnerLabel(spanComm)} comm drop${direction ? ` ${direction}` : ""}, ensure 15'6\" at midspan.`;
  }

  function commRowsInSamePoleGroup(spanComm) {
    const owner = getMROwnerLabel(spanComm).toLowerCase();
    const existingHOA = H().parseHeight(spanComm?.existingHOA || "");
    return S().getSpanCommsForPole(spanComm?.poleId || "").filter(row => {
      const rowHOA = H().parseHeight(row?.existingHOA || "");
      return getMROwnerLabel(row).toLowerCase() === owner && rowHOA === existingHOA;
    });
  }

  function commGroupTransferContext(spanComm) {
    const rows = commRowsInSamePoleGroup(spanComm);
    return {
      enabled: rows.some(row => row.transferToNewPole),
      downGuy: rows.some(row => row.downGuy || detectDownGuy(`${row.notes || ""} ${row.mr || ""}`))
    };
  }

  function generatedCommMovementMR(spanComm, action, dg = "") {
    if (!spanComm || !action) return "";
    const settings = S().getState().settings || {};
    const owner = getMROwnerLabel(spanComm);
    const existing = formattedMRHeight(spanComm.existingHOA);
    const changed = formattedMRHeight(spanComm.otherHOA || spanComm.existingHOAChange);
    if (isMetronetMR()) {
      const verb = action === "Lower" ? "lower" : "raise";
      return `At HOA ${existing} ${verb} ${owner} to HOA ${changed}${dg}.`;
    }
    // Service drops use different MR wording than regular comm movement in
    // the profiles that have a dedicated drop template. Olsson OPPD uses the
    // same raise/lower wording as every other normal communication owner.
    if (spanComm.serviceDrop && settings.showServiceDrop !== false && !isOlssonMR()) {
      return `Relocate ${owner} drop at HOA ${existing} to HOA ${changed}.`;
    }
    const verb = action === "Lower" ? "lower" : "raise";
    return `At HOA ${existing} ${verb} ${getIntecMovementOwnerLabel(spanComm)} to HOA ${changed}${dg}.`;
  }

  function generateTransferMRForPole(poleId) {
    // MidAm treats a checked Transfer as work that has already been handled
    // outside this Make Ready sheet. Do not emit a duplicate transfer MR.
    if (isMidAmMR()) return [];
    const groups = new Map();
    const settings = S().getState().settings || {};
    S().getSpanCommsForPole(poleId).forEach(row => {
      if (row.serviceDrop && settings.showServiceDrop === false) return;
      const context = commGroupTransferContext(row);
      const height = H().parseHeight(row.otherHOA || row.existingHOAChange || row.existingHOA || "");
      if (!context.enabled || height === null) return;
      const owner = getMROwnerLabel(row);
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
    const settings = S().getState().settings || {};
    // Profiles that do not require service drops (including CSU) must not
    // emit movement or re-sag instructions for imported drop rows, even if an
    // old save contains a stale HOA change or custom MR note.
    if (spanComm.serviceDrop && settings.showServiceDrop === false) return "";
    const custom = String(spanComm.mr || "").trim();
    const resag = generateResagServiceDropMR(spanComm);
    const owner = getMROwnerLabel(spanComm);
    const transferContext = commGroupTransferContext(spanComm);
    if (transferContext.enabled && isMidAmMR()) return resag;
    const dg = (transferContext.enabled ? transferContext.downGuy : (
      spanComm.downGuy || detectDownGuy(`${spanComm.notes || ""} ${spanComm.mr || ""}`)
    )) ? " with DG" : "";
    if (transferContext.enabled) {
      const transferHeight = spanComm.otherHOA || spanComm.existingHOAChange || spanComm.existingHOA;
      if (!transferHeight) return resag;
      const transfer = `Transfer ${owner} to new pole at HOA ${mrHeight(transferHeight)}${dg}.`;
      const customContainsTransfer = custom
        && custom.includes(mrHeight(spanComm.otherHOA || spanComm.existingHOAChange || spanComm.existingHOA));
      return [custom || transfer, custom && !customContainsTransfer ? transfer : "", resag].filter(Boolean).join("\n");
    }
    const action = detectRaiseLower(spanComm);
    if (!action) return [custom, resag].filter(Boolean).join("\n");
    const movement = generatedCommMovementMR(spanComm, action, dg);
    // Keep a user/imported MR note, but do not let it hide a real HOA movement
    // that is present in the calculator state. If the note already contains
    // both heights, avoid duplicating the generated instruction.
    const customContainsMovement = custom
      && custom.includes(mrHeight(spanComm.existingHOA))
      && custom.includes(mrHeight(spanComm.otherHOA || spanComm.existingHOAChange));
    return [custom, customContainsMovement ? "" : movement, resag].filter(Boolean).join("\n");
  }

  function generateMRForSpanSide(spanSide) {
    if (!spanSide) return "";
    const span = S().getSpan(spanSide.spanId);
    const slack = generateProposedSlackMR(spanSide, span);
    if (!spanSide.proposedHOA) return slack;
    const dir = span && span.direction ? ` ${span.direction}` : "";
    const items = slack ? [slack] : [];
    if (spanSide.clearanceMSReason === "LOW_POWER" && spanSide.clearanceMSIssue && shouldAddLowPowerMidspanMR()) {
      items.push(`Ensure min 30" to low power at midspan.`);
    }
    if (isMetronetMR()) {
      const pole = S().getPole(spanSide.poleId);
      if ((detectAnchor(spanSide) || detectDownGuy(spanSide.notes) || detectOHG(spanSide)) && pole?.ancActive !== false) {
        const anchor = generateAnchorInstructionFromSpan(spanSide, span?.direction || "");
        if (anchor) items.push(anchor);
      }
      if (detectRiser(spanSide) && pole?.riserActive !== false) items.push(`Pl new riser${dir}.`.replace("  ", " "));
      return items.join("\n");
    }
    // Slack is selected by the Proposed-table checkbox, not inferred from
    // free-form notes. Excel Review accepts that model-owned instruction.
    if (detectAnchor(spanSide)) items.push(`PL NEW ANC${dir}.`.replace("  ", " "));
    if (detectRiser(spanSide)) items.push(`PL NEW RISER${dir}.`.replace("  ", " "));
    return items.join("\n");
  }

  function makeReadyDirectionTokens(reference) {
    if (Array.isArray(reference?.attachmentDirectionTokens) && reference.attachmentDirectionTokens.length) {
      return reference.attachmentDirectionTokens
        .map(value => String(value || "").trim().toUpperCase())
        .filter(Boolean);
    }
    return String(reference?.attachmentDirection || "")
      .toUpperCase()
      .match(/\b(?:NE|NW|SE|SW|N|E|S|W)\b/g) || [];
  }

  function makeReadyReferenceMatchesSpan(reference, span) {
    if (!reference || !span) return false;
    const endpoint = reference.poleId === span.fromPole
      ? span.fromPole
      : reference.poleId === span.toPole ? span.toPole : "";
    if (!endpoint) return false;
    const tokens = makeReadyDirectionTokens(reference);
    if (!tokens.length) return true;
    const direction = span.fromPole === endpoint
      ? String(span.direction || "").toUpperCase()
      : oppositeDirection(String(span.direction || "").toUpperCase());
    return !direction || tokens.includes(direction);
  }

  function isDirectionalFiberOnlyOverlash(reference, span) {
    if (!reference || !span) return false;
    if (/overlash/i.test(String(reference.attachmentType || ""))) return true;
    const endpoint = reference.poleId === span.fromPole
      ? span.fromPole
      : reference.poleId === span.toPole ? span.toPole : "";
    if (!endpoint) return false;
    const direction = span.fromPole === endpoint
      ? String(span.direction || "").toUpperCase()
      : oppositeDirection(String(span.direction || "").toUpperCase());
    if (!direction) return false;
    const raw = String(reference.attachmentSizeRaw || reference.attachmentFiber || "");
    return raw.split(/\s*\+\s*/).some(part => {
      const directionMatch = part.match(/\(([^)]+)\)\s*$/);
      if (!directionMatch) return false;
      const tokens = directionMatch[1].split(/[\/,;\s]+/).map(token => String(token || "").trim().toUpperCase()).filter(Boolean);
      if (!tokens.includes(direction)) return false;
      return !/\b\d+(?:\.\d+)?\s*M\b/i.test(part);
    });
  }

  function overlashReferencesForSpan(span) {
    return (S().getState().makeReadyReferences || [])
      .filter(reference => makeReadyReferenceMatchesSpan(reference, span))
      .filter(reference => isDirectionalFiberOnlyOverlash(reference, span));
  }

  function isBackSpan(span) {
    return /back\s*span|backspan/i.test(spanRelationOf(span));
  }

  function isOverlashSpanSide(spanSide) {
    return Boolean(spanSide && overlashReferencesForSpan(S().getSpan(spanSide.spanId || "")).length);
  }

  function normalizedOwnerToken(value) {
    const normalized = String(value || "")
      .replace(/^COMMUNICATION\s*>\s*/i, "")
      .replace(/\b(?:inc|llc|communications?)\b/gi, "")
      .replace(/[^a-z0-9]+/gi, "")
      .toLowerCase();
    if (/wecom/.test(normalized)) return "wecom";
    return normalized;
  }

  function samePhysicalSpan(first, second) {
    if (!first || !second) return false;
    const left = [first.fromPole, first.toPole].filter(Boolean).sort().join("|");
    const right = [second.fromPole, second.toPole].filter(Boolean).sort().join("|");
    return Boolean(left && right && left === right);
  }

  function existingJobMessengerForSpan(span) {
    const owner = normalizedOwnerToken(proposedOwnerForMR());
    if (!owner || !span) return null;
    return Object.values(S().getState().spanComms || {}).find(row => {
      const rowSpan = S().getSpan(row.spanId || "");
      if (!samePhysicalSpan(rowSpan, span)) return false;
      return normalizedOwnerToken(ownerImportedOf(row) || ownerCanonicalOf(row) || "") === owner
        && /\bmsgr\b|messenger/i.test(`${row.size || ""} ${row.construction || ""}`);
    }) || null;
  }

  function generateOverlashMRForPole(poleId) {
    const heights = new Set();
    S().getConnectedSpans(poleId)
      // Overlash is installed on the proposed/forward span owned by this
      // pole. A reciprocal Back Span is only a reference to the adjacent
      // pole and must not add a second Overlash height to this pole's MR.
      .filter(span => span.fromPole === poleId && !isBackSpan(span))
      .forEach(span => {
        if (!overlashReferencesForSpan(span).length) return;
        const messenger = existingJobMessengerForSpan(span);
        if (!messenger) return;
        const side = S().getSpanSide(span.spanId, poleId);
        const proposed = H().parseHeight(side?.proposedHOA || "");
        const height = proposed !== null ? proposed : H().parseHeight(getEffectiveCommHOAForMR(messenger));
        if (height !== null) heights.add(height);
      });
    if (!heights.size) return [];
    const values = Array.from(heights).sort((a, b) => a - b).map(value => H().formatHeight(value));
    return [`Overlash ${proposedOwnerForMR()} at HOA ${joinMRList(values)}.`];
  }

  function generateAttachMRForPole(poleId) {
    if (isMetronetMR()) return "";
    const heights = S().getSpanSidesForPole(poleId)
      .filter(side => !isOverlashSpanSide(side))
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
    const reason = String(value || "")
      .trim()
      .replace(/[.]+$/, "")
      .replace(/^[[\](){}]+|[[\](){}]+$/g, "")
      .trim();
    if (!reason || /^(?:reason|reasoning|specify reason|insert reason|specific reason|enter specific reason)$/i.test(reason)) return "";
    return reason;
  }

  function defaultIntecUGLines(pole) {
    return [
      `Unable to attach due to ${cleanUGReason(pole?.ugReason) || "[enter specific reason]"}.`,
      "Red tag",
      "Inability to place ANC",
      "TDU replace required",
      "Existing neutral / multiplex above 26'9\"",
      "PCO neutral / multiplex exceeds 26'9\""
    ];
  }

  function defaultOlssonUGLines(pole) {
    const saved = String(pole?.ugMRText || "").trim();
    return saved ? [saved] : [];
  }

  function defaultMetronetUGLines() {
    return ["SUGGEST GOING UG DUE TO [CLEARANCE VIOLATION / 500FT/1500FT AERIAL REQUIREMENT]."];
  }

  function editableUGTemplate(pole) {
    const saved = String(pole?.ugMRText || "").trim();
    if (saved) return saved;
    if (isOlssonMR()) return "";
    return (isMetronetMR() ? defaultMetronetUGLines() : defaultIntecUGLines(pole)).join("\n");
  }

  function ugReasonFromPole(pole) {
    const selectedLine = selectedIntecUGInstruction(pole);
    const selectedMatch = String(selectedLine || "").match(/to\s+go\s+ug(?:\s+[NSEW]{1,2})?\s+due\s+to\s+(.+?)[.]*\s*$/i);
    if (selectedMatch) {
      const selectedReason = cleanUGReason(selectedMatch[1]);
      if (selectedReason) return selectedReason;
    }
    const firstLine = editableUGTemplate(pole).split(/\r?\n/).find(line => /unable\s+to\s+attach\s+due\s+to/i.test(line));
    const match = String(firstLine || "").match(/unable\s+to\s+attach\s+due\s+to\s+(.+?)[.]*\s*$/i);
    return cleanUGReason(match?.[1]) || cleanUGReason(pole?.ugReason);
  }

  function metronetUGReasonFromPole(pole) {
    const lines = ugReplacementMR(pole);
    for (const line of lines) {
      const match = String(line).match(/(?:fore|back|other)span\s+going\s+ug\s+due\s+to\s+(.+?)(?:\.|$)/i)
        || String(line).match(/suggest\s+going\s+ug\s+due\s+to\s+(.+?)(?:\.|$)/i);
      if (!match) continue;
      const reason = cleanUGReason(match[1]);
      if (!reason || /clearance\s+violation\s*\/\s*(?:insert\s+other\s+reason|500ft\s*\/\s*1500ft\s+aerial\s+requirement)/i.test(reason)) return "";
      return reason;
    }
    return cleanUGReason(pole?.ugReason);
  }

  function metronetUGPlaceholderReason() {
    return "[CLEARANCE VIOLATION / INSERT OTHER REASON]";
  }

  function ugReplacementMR(pole) {
    return editableUGTemplate(pole).split(/\r?\n/).map(line => line.trim()).filter(Boolean);
  }

  function defaultPCOLines() {
    if (isOlssonMR()) {
      return [
        "OPPD REPL POLE TO 45FT CLASS 2 DUE TO (FAILING POLE CAPACITY/CLEARANCE VIOLATIONS/BROKEN POLE/ETC)"
      ];
    }
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

  // PCO follows the same operator-owned workflow as UG: start with the
  // profile template, then persist and generate exactly the edited lines.
  function editablePCOTemplate(pole) {
    return String(pole?.pcoMRText || "").trim() || defaultPCOLines().join("\n");
  }

  function pcoReplacementMR(pole) {
    return editablePCOTemplate(pole).split(/\r?\n/).map(line => line.trim()).filter(Boolean);
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

  function ugSpanOptionsForPole(poleId) {
    const byPhysicalConnection = new Map();
    S().getConnectedSpans(poleId).forEach(span => {
        const otherPoleId = S().getOtherPoleId(span, poleId);
        const relation = normalizeSpanRelation(spanRelationOf(span), span.fromPole !== poleId);
        const direction = span.fromPole === poleId
          ? String(span.direction || "").toUpperCase()
          : oppositeDirection(String(span.direction || "").toUpperCase());
        if (!span?.spanId || !otherPoleId) return;
        const physicalKey = `${otherPoleId}::${direction}`;
        const candidate = {
          spanId: span.spanId,
          otherPoleId,
          relation,
          direction,
          label: `${relation}${direction ? ` ${direction}` : ""} · ${otherPoleId}`,
          ownedByPole: span.fromPole === poleId
        };
        const existing = byPhysicalConnection.get(physicalKey);
        // Excel can contain the reciprocal record for the same physical
        // connection. For the current pole, prefer the row owned by this pole
        // so a reciprocal row does not create a duplicate UG option.
        if (!existing || (candidate.ownedByPole && !existing.ownedByPole)) {
          byPhysicalConnection.set(physicalKey, candidate);
        }
      });
    return Array.from(byPhysicalConnection.values())
      .sort((a, b) => `${a.relation} ${a.direction} ${a.otherPoleId}`.localeCompare(`${b.relation} ${b.direction} ${b.otherPoleId}`));
  }

  function selectedIntecUGSpan(pole) {
    const spanId = String(pole?.ugSpanId || "ALL").trim();
    if (poleResolutionOf(pole) !== "UG" || !spanId || spanId.toUpperCase() === "ALL") return null;
    const span = S().getSpan(spanId);
    if (!span || (span.fromPole !== pole.poleId && span.toPole !== pole.poleId)) return null;
    return span;
  }

  function defaultIntecUGSpanLines(pole) {
    const span = selectedIntecUGSpan(pole);
    if (!span) return [];
    const relation = normalizeSpanRelation(spanRelationOf(span), span.fromPole !== pole.poleId);
    const direction = span.fromPole === pole.poleId
      ? String(span.direction || "").toUpperCase()
      : oppositeDirection(String(span.direction || "").toUpperCase());
    const reason = cleanUGReason(pole.ugReason) || "[enter specific reason]";
    return [
      `${relation} to go UG${direction ? ` ${direction}` : ""} due to ${reason}.`,
      "Red tag",
      "TDU Replace required",
      "PCO neutral (inclusive of triplex/quadruplexes as noted above) exceeds 26'9\"",
      "inability to place ANC"
    ];
  }

  function getEditableUGSpanTemplate(pole) {
    const spanId = String(pole?.ugSpanId || "ALL").trim();
    if (!spanId || spanId.toUpperCase() === "ALL") return "";
    const saved = String(pole?.ugSpanMRText || "").trim();
    if (saved && (!pole?.ugSpanMRSpanId || pole.ugSpanMRSpanId === spanId)) return saved;
    return defaultIntecUGSpanLines(pole).join("\n");
  }

  function selectedIntecUGLines(pole) {
    return getEditableUGSpanTemplate(pole).split(/\r?\n/).map(line => line.trim()).filter(Boolean);
  }

  function selectedIntecUGInstruction(pole) {
    return selectedIntecUGLines(pole)[0] || "";
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
    const explicitSpanUG = new Map();

    spans.forEach(span => {
      const otherPoleId = S().getOtherPoleId(span, poleId);
      const otherPole = S().getPole(otherPoleId);
      if (!otherPoleId || (span.ugActive !== true && poleResolutionOf(otherPole) !== "UG")) return;
      const current = byOtherPole.get(otherPoleId);
      // Excel already defines whether the span is Fore, Back, or Other from
      // the imported side. Reverse it only when we are reading the far end.
      const relation = normalizeSpanRelation(spanRelationOf(span), span.fromPole !== poleId);
      const directionFromThisPole = span.fromPole === poleId
        ? span.direction
        : oppositeDirection(span.direction);
      const candidate = {
        relation,
        direction: directionFromThisPole || "",
        spanId: span.spanId,
        otherPoleId,
        isSpanUG: span.ugActive === true,
        adjacentPoleUG: poleResolutionOf(otherPole) === "UG",
        reason: String(span.ugReason || "").trim()
      };
      // Case 2 is a span-level decision. Keep every explicitly selected UG
      // span, even when two physical spans share the same adjacent pole.
      if (candidate.isSpanUG) {
        explicitSpanUG.set(`${otherPoleId}::${span.spanId}`, candidate);
        return;
      }
      // A physical connection can appear twice in Span. Prefer the Back Span
      // row because it is the relation owned by this pole for the UG handoff.
      if (!current || (current.relation !== "Backspan" && relation === "Backspan")) {
        byOtherPole.set(otherPoleId, candidate);
      }
    });

    return [...explicitSpanUG.values(), ...Array.from(byOtherPole.values())];
  }

  function isANCEnabled(poleId) {
    const pole = S().getPole(poleId);
    if (!pole || ["UG", "PCO"].includes(poleResolutionOf(pole)) || (!isMetronetMR() && !isIntecMR())) return false;
    if (pole.ancActive === true) return true;
    if (pole.ancActive === false) return false;
    // Aerial poles receiving an adjacent UG transfer need an ANC/DG in both
    // MidAm and INTEC.  Keep an explicit `false` override authoritative, but
    // otherwise enable the automatic anchor whenever the connected UG pole is
    // the reason this pole is being worked.
    return (isMetronetMR() || isIntecMR())
      && connectedUGSpanItems(poleId).some(item => item.adjacentPoleUG);
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
      relation: normalizeSpanRelation(spanRelationOf(span), span.fromPole !== poleId),
      direction: span.fromPole === poleId ? (span.direction || "") : oppositeDirection(span.direction || ""),
      spanId: span.spanId,
      otherPoleId: S().getOtherPoleId(span, poleId) || ""
    };
  }

  function isRiserAvailable(poleId) {
    const pole = S().getPole(poleId);
    return Boolean(pole && !["UG", "PCO"].includes(poleResolutionOf(pole)));
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

  function generateRiserInstruction(poleId, options = {}) {
    const allowReplacement = options.allowReplacement === true;
    const pole = S().getPole(poleId);
    if (!pole || poleResolutionOf(pole) === "PCO" || (!allowReplacement && !isRiserAvailable(poleId))) return "";
    if (pole.riserActive === false && !allowReplacement) return "";
    if (!allowReplacement && !isRiserEnabled(poleId)) return "";
    const connection = riserConnectionForPole(poleId);
    if (allowReplacement && !connection) return "";
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

  function metronetUGReplacementMR(pole) {
    const lines = ugReplacementMR(pole);
    const hasUGRelation = lines.some(line => /\b(?:fore|back|other)span\b.*\bgoing\s+ug\b/i.test(line));
    if (hasUGRelation && !lines.some(line => /\bnew\s+anc\b.*\b(?:dg|deadending)\b/i.test(line))) {
      const relationMatch = lines
        .map(line => String(line).match(/\b(?:fore|back|other)span\b\s+going\s+ug(?:\s+([NSEW]{1,2}))?/i))
        .find(Boolean);
      const side = S().getSpanSidesForPole(pole?.poleId)
        .find(item => item.proposedHOA) || null;
      const anchor = generateConfiguredAnchorInstruction(pole) || metronetUGAnchorInstruction(pole?.poleId, {
        spanId: side?.spanId || "",
        direction: relationMatch?.[1] || resolvedRiserDirection(pole?.poleId)
      });
      lines.push(anchor || "PL NEW ANC/DG FOR DEADENDING LINES.");
    }
    if (hasUGRelation && !lines.some(line => /\briser\b/i.test(line))) {
      const directionMatch = lines
        .map(line => String(line).match(/\b(?:fore|back|other)span\b\s+going\s+ug(?:\s+([NSEW]{1,2}))?\b/i))
        .find(Boolean);
      const direction = String(resolvedRiserDirection(pole?.poleId) || directionMatch?.[1] || "").toUpperCase();
      const riser = metronetUGRiserInstruction(direction);
      if (riser) lines.push(riser);
    }
    return lines;
  }

  function metronetUGRiserInstruction(direction) {
    const normalized = String(direction || "").trim().toUpperCase();
    if (!/^[NSEW]{1,2}$/.test(normalized)) return "";
    return `PL NEW RISER FOR UG TRANSFER ${normalized}.`;
  }

  function metronetUGAnchorInstruction(poleId, connection) {
    const side = S().getSpanSide?.(connection?.spanId || "", poleId)
      || S().getSpanSidesForPole(poleId).find(item => item.spanId === connection?.spanId)
      || null;
    const pole = S().getPole(poleId);
    const proposedHOA = side?.proposedHOA || pole?.standaloneProposedHOA || "";
    if (!proposedHOA) return "";
    return generateAnchorInstructionFromSpan({ ...(side || {}), proposedHOA }, connection?.direction || "");
  }

  function buildConnectedUGInstructions(poleId) {
    const items = connectedUGSpanItems(poleId);

    const lines = items.flatMap(item => {
      const direction = item.direction ? ` ${item.direction}` : "";
      if (isMetronetMR()) {
        const relation = item.relation === "Otherspan" ? "Other Span" : item.relation;
        const adjacentReason = item.isSpanUG
          ? (cleanUGReason(item.reason) || metronetUGPlaceholderReason())
          : (metronetUGReasonFromPole(S().getPole(item.otherPoleId)) || metronetUGPlaceholderReason());
        const anchor = item.isSpanUG ? "" : (isANCEnabled(poleId)
          ? (generateConfiguredAnchorInstruction(S().getPole(poleId)) || metronetUGAnchorInstruction(poleId, item) || "Pl new ANC/DG for deadending lines.")
          : "");
        const transferDirection = resolvedRiserDirection(poleId, item) || item.direction;
        const riser = S().getPole(poleId)?.riserActive === false
          ? ""
          : metronetUGRiserInstruction(transferDirection);
        return [
          item.isSpanUG
            ? `${relation} going UG due to ${adjacentReason}.`
            : `${relation} going UG${direction} due to on adj pole ${adjacentReason}.`,
          ...(anchor ? [anchor] : []),
          ...(riser ? [riser] : [])
        ];
      }
      const adjacentReason = item.isSpanUG
        ? (cleanUGReason(item.reason) || "[insert other reason]")
        : ugReasonFromPole(S().getPole(item.otherPoleId));
      if (!adjacentReason) return [];
      return [item.isSpanUG
        ? `${item.relation} to go UG due to ${adjacentReason}.`
        : `${item.relation} to go UG${direction} due to on adj pole ${adjacentReason}.`];
    });

    const hasMetronetUGTransferRiser = lines.some(line => /^pl\s+new\s+riser\s+for\s+ug\s+transfer\b/i.test(line));
    const riser = hasMetronetUGTransferRiser ? "" : generateRiserInstruction(poleId);
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
        if (row.raiseActive) {
          const source = H().parseHeight(row.attachmentHeight || "");
          const target = H().parseHeight(row.raiseHeight || "");
          if (source !== null && target !== null && target > source && target <= source + 12) {
            lines.push(isMidAmMR()
              ? `AT HOA ${H().formatHeight(source)} RAISE STREETLIGHT TO HOA ${H().formatHeight(target)} DUE TO CLEARANCE VIOLATIONS.`
              : `Raise streetlight from HOA ${H().formatHeight(source)} to ${H().formatHeight(target)}.`);
          }
        }
        return lines;
      }

      const target = H().parseHeight(row.actionHeight || "");
      if (category.includes("TRANSFORMER")) {
        if (!row.actionActive || target === null) return [];
        const targetText = H().formatHeight(target);
        return [isMetronetMR()
          ? `POWER REDRESS TRANSFORMER DRIP LOOP TO HOA ${targetText}.`
          : `Secure transformer drip loop to HOA ${targetText}.`];
      }
      if (category.includes("RISER")) {
        const source = H().parseHeight(row.attachmentHeight || "");
        const lines = [];
        if (row.actionActive && source !== null && target !== null && target > source) {
          const targetText = H().formatHeight(target);
          lines.push(isMetronetMR()
            ? `AT HOA ${H().formatHeight(source)} RAISE POWER RISER TO HOA ${targetText} DUE TO CLEARANCES.`
            : isOlssonMR()
              ? `AT HOA ${H().formatHeight(source)} RAISE OPPD RISER TO HOA ${targetText} DUE TO CLEARANCES.`
              : `Raise APS riser from HOA ${H().formatHeight(source)} to HOA ${targetText}.`);
        }
        if (row.secureActive && source !== null) {
          const effectiveRiser = row.actionActive && target !== null && target > source ? target : source;
          lines.push(`Secure riser drip loop to HOA ${H().formatHeight(effectiveRiser - 8)}.`);
        }
        return lines;
      }
      if (!row.actionActive || target === null) return [];
      return [];
    });
  }

  function generateSecondaryDripLoopMR(pole) {
    if (!pole?.secondaryDripLoopActive) return "";
    const height = H().parseHeight(pole.secondaryDripLoopHoa || "");
    return height === null ? "" : `Secure secondary drip loop to HOA ${H().formatHeight(height)}.`;
  }

  function generatePoleInsetMR(pole) {
    if (!pole?.poleInsetActive || ["UG", "PCO"].includes(poleResolutionOf(pole))) return "";
    const settings = S().getState().settings || {};
    if (["METRONET", "OLSSON_OPPD"].includes(String(settings.projectProfile || "").toUpperCase())) return "";
    const reason = String(pole.poleInsetReason || "").toUpperCase() === "FAILING_CLEARANCES"
      ? "failing clearances"
      : "overloaded";
    const details = global.Calculations?.getPoleInsetDetails?.(pole.poleId) || {};
    const direction = String(details.direction || "").trim().toUpperCase();
    const directionSuffix = direction ? ` ${direction}` : "";
    return `Pole ${reason} unless comm inset pole placed at midspan to reduce span length${directionSuffix}.`;
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
    const poleActionNotes = [];
    const poleInset = [];
    const commMoves = [];
    const dropMoves = [];
    const resagGroups = new Map();
    const proposed = [];
    const slackGroups = new Map();
    const ensure = [];
    const ensureDirections = [];
    let ensureSpanCount = 0;
    const risers = [];
    const anc = [];
    const pole = S().getPole(poleId);
    const redTag = generateRedTagMR(pole);
    if (["UG", "PCO"].includes(poleResolutionOf(pole))) {
      const selectedUGLines = poleResolutionOf(pole) === "UG" && isIntecMR() ? selectedIntecUGLines(pole) : [];
      const lines = poleResolutionOf(pole) === "UG"
        ? (selectedUGLines.length
          ? selectedUGLines
          : (isMetronetMR() ? metronetUGReplacementMR(pole) : (isOlssonMR() ? defaultOlssonUGLines(pole) : ugReplacementMR(pole))))
        : pcoReplacementMR(pole);
      const text = [redTag, ...lines].filter(Boolean).map(applyCase).join("\n");
      if (text) state.mr.push({ poleId, spanId: "", owner: "MR", text, imported: false });
      return state.mr.filter(item => item.poleId === poleId);
    }
    buildConnectedUGInstructions(poleId).forEach(line => {
      // Riser work is always the final instruction crews read in the pole MR.
      // Keep the UG relation at the top, place ANC after power work, and defer
      // the separate riser line until the end.
      if (/^pl\s+(?:new\s+)?riser\b/i.test(line)) risers.push(line);
      else if (/^pl\s+new\b.*\banc\b/i.test(line)) anc.push(line);
      else ug.push(line);
    });
    const inset = generatePoleInsetMR(pole);
    if (inset) poleInset.push(inset);
    power.push(...generatePowerEquipmentMRForPole(poleId));
    const olssonCommClearanceNote = generateOlssonCommClearanceNote(pole);
    if (olssonCommClearanceNote) poleActionNotes.push(olssonCommClearanceNote);
    if (isMetronetMR() || isIntecMR() || isOlssonMR()) {
      const configuredAnchor = generateConfiguredAnchorInstruction(pole);
      const configuredOHG = metronetConfiguredOHGMR(pole);
      if ((isMetronetMR() || isIntecMR()) && configuredAnchor) anc.push(configuredAnchor);
      if (configuredOHG) anc.push(configuredOHG);
    }
    const secondaryDripLoop = generateSecondaryDripLoopMR(pole);
    if (secondaryDripLoop) power.push(secondaryDripLoop);
    commMoves.push(...generateTransferMRForPole(poleId));
    S().getSpanSidesForPole(poleId).forEach(side => {
      const text = generateMRForSpanSide(side);
      if (text) text.split(/\n+/).filter(Boolean).forEach(line => {
        const slackMatch = line.match(/^(Proposed slack span|PROPOSED SLACK SPAN|PL NEW SLACK SPAN)\s+([NSEW]{1,2})\.$/i);
        if (slackMatch) {
          const prefix = slackMatch[1];
          const key = prefix.toUpperCase();
          if (!slackGroups.has(key)) slackGroups.set(key, { prefix, directions: [] });
          const group = slackGroups.get(key);
          const direction = slackMatch[2].toUpperCase();
          if (!group.directions.includes(direction)) group.directions.push(direction);
          return;
        }
        if (/ensure min 30/i.test(line)) {
          ensure.push(line);
          ensureSpanCount += 1;
          if (isIntecMR()) {
            const span = S().getSpan(side.spanId);
            const direction = String(span?.direction || "").trim().toUpperCase();
            if (direction && !ensureDirections.includes(direction)) ensureDirections.push(direction);
          }
        }
        else proposed.push(line);
      });
    });
    slackGroups.forEach(group => {
      proposed.unshift(`${group.prefix} ${joinMRList(group.directions)}.`);
    });
    // Every INTEC Ensure instruction identifies the affected span direction.
    // When several proposed spans exist, only directions whose own span has a
    // low-power issue are collected, so a single affected span remains a
    // single direction while multiple affected spans are joined naturally.
    if (isIntecMR() && ensureSpanCount > 0 && ensureDirections.length) {
      for (let index = ensure.length - 1; index >= 0; index -= 1) {
        if (/ensure min 30/i.test(ensure[index])) ensure.splice(index, 1);
      }
      ensure.push(`Ensure min 30" to low power at midspan ${joinMRList(ensureDirections)}.`);
    }
    const collectResag = spanComm => {
      const resag = generateResagServiceDropMR(spanComm);
      if (!resag) return;
      const owner = getMROwnerLabel(spanComm);
      const key = owner.toLowerCase();
      if (!resagGroups.has(key)) resagGroups.set(key, { owner, directions: [], undirected: false });
      const group = resagGroups.get(key);
      const direction = directionForSpanComm(spanComm);
      if (direction) {
        if (!group.directions.includes(direction)) group.directions.push(direction);
      } else {
        group.undirected = true;
      }
    };
    S().getSpanCommsForPole(poleId)
      .slice()
      .sort((a, b) => (H().parseHeight(getEffectiveCommHOAForMR(b)) ?? -Infinity) - (H().parseHeight(getEffectiveCommHOAForMR(a)) ?? -Infinity))
      .forEach(sc => {
      if (commGroupTransferContext(sc).enabled) {
        collectResag(sc);
        return;
      }
      const text = generateMRForComm(sc);
      if (!text) return;
      text.split(/\n+/).filter(Boolean).forEach(line => {
        if (/^Re-sag\b/i.test(line)) collectResag(sc);
        else if (sc.serviceDrop && !sc.transferToNewPole) dropMoves.push(line);
        else commMoves.push(line);
      });
    });
    resagGroups.forEach(group => {
      const directions = group.directions.slice();
      if (group.undirected && !directions.length) directions.push("");
      const directionText = directions.filter(Boolean).length
        ? ` ${joinMRList(directions.filter(Boolean))}`
        : "";
      dropMoves.push(`Re-sag ${group.owner} comm drop${directionText}, ensure 15'6" at midspan.`);
    });
    proposed.push(...generateOverlashMRForPole(poleId));
    const attach = generateAttachMRForPole(poleId);
    if (attach) proposed.unshift(attach);
    // Keep the new attachment immediately before ANC work.  Other proposed
    // instructions remain in their existing position after the ANC block.
    const attachmentLine = attach ? proposed.shift() : "";

    // Ensure reminders belong after all existing communication movements but
    // before the new attachment/proposed work they protect.
    const lines = [redTag, ...poleInset, ...ug, ...power, ...poleActionNotes, ...commMoves, ...dropMoves, ...ensure, attachmentLine, ...anc, ...proposed, ...risers].filter(Boolean).map(applyCase);
    const seen = new Set();
    const unique = lines.map(line => line.trim()).filter(Boolean).filter(line => {
      const key = normalizeMRLine(line);
      if (seen.has(key)) return false;
      seen.add(key);
      return true;
    });
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

  // Editing the UG reason changes the relation-specific instruction on the
  // adjacent pole's Forespan/Backspan. Regenerate the UG pole and every
  // directly connected pole so those handoff instructions never remain stale.
  function generateMRForUGChange(poleId) {
    const affectedPoleIds = new Set([poleId]);
    S().getConnectedSpans(poleId).forEach(span => {
      const otherPoleId = S().getOtherPoleId(span, poleId);
      if (otherPoleId) affectedPoleIds.add(otherPoleId);
    });
    affectedPoleIds.forEach(generateMRForPole);
    return Array.from(affectedPoleIds);
  }

  function getEffectiveCommHOAForMR(spanComm) {
    return spanComm?.otherHOA || spanComm?.existingHOAChange || spanComm?.existingHOA || "";
  }

  /** @namespace MRLogic */
  global.MRLogic = {
    generateMRForPole,
    generateMRForSpan,
    generateMRForComm,
    generateResagServiceDropMR,
    generatePowerEquipmentMRForPole,
    generateOverlashMRForPole,
    generatePoleInsetMR,
    generateRedTagMR,
    getEditableUGTemplate: editableUGTemplate,
    getEditablePCOTemplate: editablePCOTemplate,
    getImportedRiserDirection: importedRiserDirection,
    getResolvedRiserDirection: resolvedRiserDirection,
    getDefaultRiserDirection: defaultRiserDirection,
    getUGSpanOptions: ugSpanOptionsForPole,
    getEditableUGSpanTemplate,
    getDefaultUGSpanTemplate: pole => defaultIntecUGSpanLines(pole).join("\n"),
    getSelectedIntecUGInstruction: selectedIntecUGInstruction,
    isRiserAvailable,
    isRiserEnabled,
    isANCEnabled,
    getEffectiveANCDistance: poleId => getEffectiveAnchorDistance(S().getPole(poleId)),
    getEffectiveANCDirection: poleId => getEffectiveAnchorDirection(S().getPole(poleId)),
    getEffectiveANCDGHOA: poleId => getEffectiveAnchorDGHOA(S().getPole(poleId)),
    generateRiserInstruction,
    generateMRForSpanSide,
    generateAllMR,
    generateMRForUGChange,
    detectAttach: detectAttachFromSpanSide,
    detectRaiseLower,
    detectOverlash,
    detectAnchor,
    detectRiser
  };
})(window);
