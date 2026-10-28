(function (global) {
  "use strict";

  function create(context = {}) {
    const S = context.S || global.AppStore;
    const ownerCanonicalOf = context.ownerCanonicalOf;
    const safeJobFilePart = context.safeJobFilePart;
  function spanSideHasUserWork(side) {
    return Boolean(side?.proposedHOA || side?.proposedHOAChange || side?.proposedMidspan || side?.ocalcMS || side?.notes || side?.isManualProposed || side?.isSlack);
  }

  function spanCommHasUserWork(row) {
    return Boolean(
      row?.existingHOAChange ||
      row?.otherHOA ||
      row?.otherHOAActive ||
      row?.notes ||
      row?.mr ||
      row?.serviceDrop ||
      row?.downGuy ||
      row?.transferToNewPole ||
      row?.resagServiceDrop ||
      String(row?.wireId || "").startsWith("manual-")
    );
  }

  // Imported HOA and midspan values remain calculation inputs even when the
  // user has not edited the row. Update Data must not delete those baselines
  // merely because a later workbook omits the corresponding Span.Wire row.
  function spanCommHasCalculationData(row) {
    return Boolean(
      row?.existingHOA ||
      row?.midspan ||
      row?.ocalcMS ||
      row?.calculatedMidspan ||
      row?.msProposed ||
      row?.finalMidspan
    );
  }

  function normalizedCommOwner(value) {
    return String(value || "")
      .replace(/^communication\s*>\s*/i, "")
      .replace(/century\s*link/g, "centurylink")
      .trim()
      .toLowerCase();
  }

  // Excel updates can change Wire Id while describing the same physical comm.
  // Match exact keys first, then reconcile by span, pole and owner so saved HOA
  // movements are applied to the newly imported row instead of creating a stale
  // duplicate beside it.
  function findImportedCommMatch(spanComms, oldKey, oldRow, claimedKeys) {
    if (spanComms[oldKey] && !claimedKeys.has(oldKey)) {
      return { key: oldKey, row: spanComms[oldKey], hadCandidates: true, matchType: "exact" };
    }

    const owner = normalizedCommOwner(ownerCanonicalOf(oldRow));
    const logicalCandidates = Object.entries(spanComms).filter(([, row]) =>
      row.spanId === oldRow.spanId &&
      row.poleId === oldRow.poleId &&
      normalizedCommOwner(ownerCanonicalOf(row)) === owner
    );
    const candidates = logicalCandidates.filter(([key]) => !claimedKeys.has(key));
    if (!candidates.length) return { key: "", row: null, hadCandidates: logicalCandidates.length > 0, matchType: "" };

    candidates.sort(([, a], [, b]) => {
      const score = row =>
        (oldRow.wireId && row.wireId === oldRow.wireId ? 100 : 0) +
        (oldRow.wireIndex && row.wireIndex === oldRow.wireIndex ? 40 : 0) +
        (oldRow.existingHOA && row.existingHOA === oldRow.existingHOA ? 20 : 0) +
        (oldRow.size && row.size === oldRow.size ? 10 : 0);
      return score(b) - score(a);
    });
    return { key: candidates[0][0], row: candidates[0][1], hadCandidates: true, matchType: "logical" };
  }

  function mergeCommUserWork(importedRow, oldRow) {
    // Service Drop is an imported physical-wire classification. Do not carry
    // a stale flag from an older row when the refreshed Span.Wire size says
    // that the cable is a regular CATV/Telco bundle (or vice versa).
    const importedSize = String(importedRow.size || "").trim();
    const importedServiceDrop = importedSize
      ? /communication\s*drops?|service\s*drop|\bdrop\b/i.test(importedSize)
      : null;
    return {
      ...importedRow,
      existingHOAChange: oldRow.existingHOAChange || importedRow.existingHOAChange || "",
      otherHOA: oldRow.otherHOA || importedRow.otherHOA || "",
      otherHOAActive: Boolean(oldRow.otherHOAActive || importedRow.otherHOAActive),
      serviceDrop: importedServiceDrop === null
        ? Boolean(oldRow.serviceDrop || importedRow.serviceDrop)
        : importedServiceDrop,
      downGuy: Boolean(oldRow.downGuy || importedRow.downGuy),
      transferToNewPole: Boolean(oldRow.transferToNewPole),
      resagServiceDrop: Boolean(oldRow.resagServiceDrop),
      notes: oldRow.notes || importedRow.notes || "",
      mr: oldRow.mr || importedRow.mr || "",
      // These fields are calculated again from the newly imported base data.
      remotePoleId: "",
      remoteHOA: "",
      calculatedMidspan: "",
      difference: "",
      msProposed: "",
      finalMidspan: "",
      clearanceMSStatus: "",
      clearanceMSMessage: "",
      clearanceMSIssue: false,
      flaggingStatus: "",
      flaggingMessage: ""
    };
  }

  function isBlankUpdatedValue(value) {
    return value === undefined || value === null || (typeof value === "string" && !value.trim());
  }

  function hasStoredValue(value) {
    if (isBlankUpdatedValue(value)) return false;
    if (Array.isArray(value)) return value.length > 0;
    if (value && typeof value === "object") return Object.keys(value).length > 0;
    return true;
  }

  // Update Data is additive. A populated cell from the new workbook replaces
  // the old value, while an empty imported cell keeps the last known value.
  // The raw Excel Review snapshot is intentionally excluded and always
  // describes the newly selected workbook exactly as it arrived.
  function preserveValuesMissingFromUpdate(importedEntity, oldEntity, reconciliation) {
    const next = { ...(importedEntity || {}) };
    Object.entries(oldEntity || {}).forEach(([field, oldValue]) => {
      if (!isBlankUpdatedValue(next[field]) || !hasStoredValue(oldValue)) return;
      next[field] = oldValue && typeof oldValue === "object"
        ? JSON.parse(JSON.stringify(oldValue))
        : oldValue;
      reconciliation.blankValuesPreserved += 1;
    });
    return next;
  }

  function normalizeReviewHeader(value) {
    return String(value || "")
      .toLowerCase()
      .normalize("NFD")
      .replace(/[\u0300-\u036f]/g, "")
      .replace(/[^a-z0-9]/g, "");
  }

  function reviewCell(row, names) {
    const wanted = names.map(normalizeReviewHeader);
    const key = Object.keys(row || {}).find(candidate => wanted.includes(normalizeReviewHeader(candidate)));
    return key ? String(row[key] ?? "").trim() : "";
  }

  // Raw review sheets need their own stable identities because Update Data
  // may contain only one pole. These keys replace matching source rows while
  // retaining every unrelated row from the previously loaded job workbook.
  function reviewRowIdentity(sheetName, row) {
    const poleId = reviewCell(row, ["Id", "Pole ID", "PoleId", "Pole", "Structure Number"]);
    const poleKey = S.canonicalPoleIdentity(poleId) || poleId.toUpperCase();
    const collectionId = reviewCell(row, ["collectionId", "Collection ID"]);
    const spanId = reviewCell(row, ["Span Id", "Span ID", "spanId", "Wire Span ID"]);
    const spanIndex = reviewCell(row, ["Span Index"]);
    const wireId = reviewCell(row, ["Wire Id", "Wire ID", "wireId"]);
    const wireIndex = reviewCell(row, ["Wire Index"]);
    const owner = reviewCell(row, ["Owner"]);
    const type = reviewCell(row, ["Type", "Equipment Type", "Attachment Type"]);
    const attachment = reviewCell(row, ["Attachment Height.display", "Attachment Height Display", "Attachment Height"]);
    const orientation = reviewCell(row, ["Orientation", "Direction"]);
    const genericId = reviewCell(row, ["Make Ready Id", "Make Ready ID", "MR ID", "Equipment Id", "Equipment ID", "Anchor Id", "Anchor ID", "Guy Id", "Guy ID"]);
    const genericIndex = reviewCell(row, ["Make Ready Index", "MR Index", "Equipment Index", "Anchor Index", "Guy Index"]);

    if (sheetName === "collection") return poleKey || collectionId;
    if (sheetName === "spans") return spanId || [collectionId, spanIndex, poleKey].join("|");
    if (sheetName === "spanWires") return [spanId, poleKey, wireId || wireIndex || `${owner}|${type}|${attachment}`].join("|");
    return [poleKey || collectionId, genericId || genericIndex || `${type}|${owner}|${attachment}|${orientation}`].join("|");
  }

  function mergeReviewSheetSnapshot(sheetName, previousSnapshot, importedSnapshot, reconciliation) {
    const rows = (importedSnapshot?.rows || []).map(row => ({ ...row }));
    const importedByIdentity = new Map();
    rows.forEach((row, index) => {
      const identity = reviewRowIdentity(sheetName, row) || JSON.stringify(row);
      if (!importedByIdentity.has(identity)) importedByIdentity.set(identity, []);
      importedByIdentity.get(identity).push(index);
    });
    const oldOccurrences = new Map();
    (previousSnapshot?.rows || []).forEach(oldRow => {
      const identity = reviewRowIdentity(sheetName, oldRow) || JSON.stringify(oldRow);
      const occurrence = oldOccurrences.get(identity) || 0;
      oldOccurrences.set(identity, occurrence + 1);
      const importedIndex = importedByIdentity.get(identity)?.[occurrence];
      if (importedIndex !== undefined) {
        rows[importedIndex] = preserveValuesMissingFromUpdate(rows[importedIndex], oldRow, reconciliation);
        return;
      }
      rows.push({ ...oldRow });
      reconciliation.reviewRowsPreserved += 1;
    });
    return {
      headers: Array.from(new Set([...(importedSnapshot?.headers || []), ...(previousSnapshot?.headers || [])])),
      rows
    };
  }

  function mergeExcelReviewSources(previousSource, importedSource, reconciliation) {
    const sheetNames = ["collection", "spans", "spanWires", "equipment", "anchors", "anchorGuys", "makeReady", "commTransfers"];
    return Object.fromEntries(sheetNames.map(sheetName => [
      sheetName,
      mergeReviewSheetSnapshot(sheetName, previousSource?.[sheetName], importedSource?.[sheetName], reconciliation)
    ]));
  }

  function findImportedPowerMatch(spanPower, oldKey, oldRow, claimedKeys) {
    if (spanPower[oldKey] && !claimedKeys.has(oldKey)) return { key: oldKey, row: spanPower[oldKey] };
    const oldOwner = normalizedCommOwner(oldRow.owner);
    const oldSize = String(oldRow.size || "").trim().toLowerCase();
    if (!oldOwner && !oldSize) return { key: "", row: null };
    const candidates = Object.entries(spanPower).filter(([key, row]) => (
      !claimedKeys.has(key)
      && row.spanId === oldRow.spanId
      && row.poleId === oldRow.poleId
      && normalizedCommOwner(row.owner) === oldOwner
      && String(row.size || "").trim().toLowerCase() === oldSize
    ));
    return candidates.length ? { key: candidates[0][0], row: candidates[0][1] } : { key: "", row: null };
  }

  function referenceMatchScore(importedRef, oldRef) {
    if (String(importedRef.poleId || "").trim() !== String(oldRef.poleId || "").trim()) return -1;
    let score = 0;
    if (oldRef.makeReadyId && importedRef.makeReadyId === oldRef.makeReadyId) score += 100;
    if (oldRef.makeReadyIndex && importedRef.makeReadyIndex === oldRef.makeReadyIndex) score += 80;
    if (oldRef.attachmentSizeRaw && importedRef.attachmentSizeRaw === oldRef.attachmentSizeRaw) score += 40;
    if (oldRef.attachmentDirection && importedRef.attachmentDirection === oldRef.attachmentDirection) score += 20;
    if (oldRef.attachmentHeight && importedRef.attachmentHeight === oldRef.attachmentHeight) score += 10;
    return score;
  }

  function mergeMakeReadyReferences(previousRows, importedRows, reconciliation) {
    const mergedRows = Array.isArray(importedRows) ? importedRows.map(row => ({ ...row })) : [];
    const claimed = new Set();
    (previousRows || []).forEach(oldRef => {
      const candidates = mergedRows
        .map((row, index) => ({ row, index, score: claimed.has(index) ? -1 : referenceMatchScore(row, oldRef) }))
        .filter(item => item.score >= 0)
        .sort((a, b) => b.score - a.score);
      const selected = candidates.find(item => item.score > 0) || (candidates.length === 1 ? candidates[0] : null);
      if (!selected) return;
      claimed.add(selected.index);
      mergedRows[selected.index] = preserveValuesMissingFromUpdate(selected.row, oldRef, reconciliation);
    });
    return mergedRows;
  }

  function buildPoleAliasMap(previousPoles, importedPoles) {
    const importedByCanonical = new Map();
    Object.keys(importedPoles || {}).forEach(poleId => {
      const canonical = S.canonicalPoleIdentity(poleId);
      if (!importedByCanonical.has(canonical)) importedByCanonical.set(canonical, []);
      importedByCanonical.get(canonical).push(poleId);
    });
    const aliases = new Map();
    Object.keys(previousPoles || {}).forEach(oldPoleId => {
      if (importedPoles?.[oldPoleId]) {
        aliases.set(oldPoleId, oldPoleId);
        return;
      }
      const candidates = importedByCanonical.get(S.canonicalPoleIdentity(oldPoleId)) || [];
      if (candidates.length === 1) aliases.set(oldPoleId, candidates[0]);
    });
    return aliases;
  }

  function remapPoleId(poleId, aliases) {
    return aliases.get(poleId) || poleId;
  }

  function remapSpanEndpoints(span, aliases) {
    return {
      ...span,
      fromPole: remapPoleId(span.fromPole, aliases),
      toPole: remapPoleId(span.toPole, aliases)
    };
  }

  function physicalSpanKey(span, aliases) {
    if (!span) return "";
    const endpoints = [
      remapPoleId(span.fromPole, aliases),
      remapPoleId(span.toPole, aliases)
    ].filter(Boolean).map(value => S.canonicalPoleIdentity(value)).sort();
    return endpoints.length === 2 ? endpoints.join("<->") : "";
  }

  function endpointPlaceholderGroupKey(row, spans, aliases) {
    const spanKey = physicalSpanKey(spans?.[row.spanId], aliases);
    const poleKey = S.canonicalPoleIdentity(remapPoleId(row.poleId, aliases));
    const ownerKey = normalizedCommOwner(ownerCanonicalOf(row));
    return spanKey && poleKey && ownerKey ? `${poleKey}|${spanKey}|${ownerKey}` : "";
  }

  // normalizeState() creates owner-only rows at a completely empty endpoint.
  // During Update Data those helpers must not become new blank cables when the
  // previous state already owns the same physical connection. Match exact wire
  // IDs first, then reconcile remaining rows one-for-one within the same pole,
  // physical span and owner. Extra new wires remain available to import.
  function discardSupersededEndpointPlaceholders(previous, merged, aliases, reconciliation) {
    const oldGroups = new Map();
    Object.values(previous.spanComms || {}).forEach(source => {
      if (source.isEndpointPlaceholder && !spanCommHasCalculationData(source) && !spanCommHasUserWork(source)) return;
      const row = { ...source, poleId: remapPoleId(source.poleId, aliases) };
      const groupKey = endpointPlaceholderGroupKey(row, previous.spans || {}, aliases);
      if (!groupKey) return;
      if (!oldGroups.has(groupKey)) oldGroups.set(groupKey, []);
      oldGroups.get(groupKey).push(row);
    });

    const placeholderGroups = new Map();
    Object.entries(merged.spanComms || {}).forEach(([key, row]) => {
      if (!row.isEndpointPlaceholder) return;
      const groupKey = endpointPlaceholderGroupKey(row, merged.spans || {}, aliases);
      if (!groupKey) return;
      if (!placeholderGroups.has(groupKey)) placeholderGroups.set(groupKey, []);
      placeholderGroups.get(groupKey).push({ key, row });
    });

    placeholderGroups.forEach((placeholders, groupKey) => {
      const oldRows = [...(oldGroups.get(groupKey) || [])];
      if (!oldRows.length) return;
      placeholders.forEach(placeholder => {
        if (!oldRows.length) return;
        const exactIndex = placeholder.row.wireId
          ? oldRows.findIndex(row => row.wireId && row.wireId === placeholder.row.wireId)
          : -1;
        oldRows.splice(exactIndex >= 0 ? exactIndex : 0, 1);
        delete merged.spanComms[placeholder.key];
        reconciliation.endpointPlaceholdersDiscarded += 1;
      });
    });
  }

  function equipmentMatchScore(importedRow, oldRow) {
    if (!importedRow || !oldRow) return -1;
    const clean = value => String(value || "").trim().toLowerCase();
    if (clean(oldRow.equipmentId) && clean(importedRow.equipmentId) === clean(oldRow.equipmentId)) return 100;
    const samePhysicalRow = ["category", "type", "owner", "orientation", "attachmentHeight", "bottomHeight", "dripLoopHeight"]
      .every(field => clean(importedRow[field]) === clean(oldRow[field]));
    if (samePhysicalRow) return 80;
    if (clean(oldRow.equipmentIndex) && clean(importedRow.equipmentIndex) === clean(oldRow.equipmentIndex)) return 60;
    return -1;
  }

  // Equipment actions are user work. Reimport fresh physical data, then put
  // each saved action back on its matching row. If Excel temporarily omits an
  // acted-on row, keep it so Update Data cannot silently delete that work.
  function mergePowerEquipmentUserWork(importedRows, oldRows, reconciliation) {
    const rows = Array.isArray(importedRows) ? importedRows.map(row => ({ ...row })) : [];
    // A partial Update Data file often has no Equipment sheet at all. In that
    // case the empty imported list means "not supplied", not "delete power".
    if (!rows.length && (oldRows || []).length) return oldRows.map(row => ({ ...row }));
    const claimed = new Set();
    (oldRows || []).filter(row => row.actionActive || row.actionHeight || row.raiseActive || row.secureActive || row.raiseHeight).forEach(oldRow => {
      const match = rows
        .map((row, index) => ({ index, score: claimed.has(index) ? -1 : equipmentMatchScore(row, oldRow) }))
        .filter(item => item.score >= 0)
        .sort((a, b) => b.score - a.score)[0];
      if (match) {
        claimed.add(match.index);
        rows[match.index] = {
          ...rows[match.index],
          actionActive: Boolean(oldRow.actionActive),
          actionHeight: oldRow.actionHeight || "",
          raiseActive: Boolean(oldRow.raiseActive),
          secureActive: Boolean(oldRow.secureActive),
          raiseHeight: oldRow.raiseHeight || ""
        };
        reconciliation.equipmentActionsPreserved += 1;
      } else {
        rows.push({ ...oldRow });
        reconciliation.missingEquipmentActionsPreserved += 1;
      }
    });
    return rows;
  }

  // A deleted pole is job-level user intent. Apply its canonical tombstone to
  // the freshly imported graph so Update Data cannot restore it as a renamed
  // Collection row (for example, with STEEL newly appended).
  function suppressDeletedPoles(merged, deletedPoleIds) {
    const deleted = new Set((deletedPoleIds || []).map(S.canonicalPoleIdentity).filter(Boolean));
    if (!deleted.size) return 0;
    const removedPoleIds = Object.keys(merged.poles || {}).filter(id => deleted.has(S.canonicalPoleIdentity(id)));
    if (!removedPoleIds.length) return 0;
    const removed = new Set(removedPoleIds);
    const spanIds = new Set(Object.values(merged.spans || {})
      .filter(span => removed.has(span.fromPole) || removed.has(span.toPole))
      .map(span => span.spanId));
    let foundDependent = true;
    while (foundDependent) {
      foundDependent = false;
      Object.values(merged.spans || {}).forEach(span => {
        if (spanIds.has(span.spanId) || !span.sourceSpanId || !spanIds.has(span.sourceSpanId)) return;
        spanIds.add(span.spanId);
        foundDependent = true;
      });
    }
    removedPoleIds.forEach(id => delete merged.poles[id]);
    spanIds.forEach(id => delete merged.spans[id]);
    ["spanSides", "spanComms", "spanPower"].forEach(mapName => {
      Object.keys(merged[mapName] || {}).forEach(key => {
        const row = merged[mapName][key];
        if (removed.has(row.poleId) || spanIds.has(row.spanId)) delete merged[mapName][key];
      });
    });
    merged.makeReadyReferences = (merged.makeReadyReferences || []).filter(row => !removed.has(row.poleId));
    merged.poleClassChecks = (merged.poleClassChecks || []).filter(row => !removed.has(row.poleId));
    merged.movements = (merged.movements || []).filter(row => !removed.has(row.poleId));
    merged.mr = (merged.mr || []).filter(row => !removed.has(row.poleId));
    merged.warnings = (merged.warnings || []).filter(row => !removed.has(row.poleId) && !spanIds.has(row.spanId));
    return removedPoleIds.length;
  }

  function mergeImportedUpdate(previous, imported) {
    const merged = JSON.parse(JSON.stringify(imported));
    // Update Data refreshes workbook facts, but the operator-owned job label
    // remains stable for subsequent Save and export operations.
    merged.jobName = previous.jobName || merged.jobName
      || (S.jobNameFromFileName ? S.jobNameFromFileName(merged.importedFileName) : safeJobFilePart(merged.importedFileName));
    merged.poles = merged.poles || {};
    merged.spans = merged.spans || {};
    merged.spanSides = merged.spanSides || {};
    merged.spanComms = merged.spanComms || {};
    merged.spanPower = merged.spanPower || {};
    const reconciliation = {
      exactCommMatches: 0,
      logicalCommMatches: 0,
      staleDuplicateCommsDiscarded: 0,
      endpointPlaceholdersDiscarded: 0,
      unmatchedUserCommsPreserved: 0,
      baselineCommRowsPreserved: 0,
      equipmentActionsPreserved: 0,
      missingEquipmentActionsPreserved: 0,
      reviewRowsPreserved: 0,
      blankValuesPreserved: 0,
      missingRowsPreserved: 0,
      deletedPolesSuppressed: 0
    };
    const poleAliases = buildPoleAliasMap(previous.poles || {}, merged.poles);
    discardSupersededEndpointPlaceholders(previous, merged, poleAliases, reconciliation);

    Object.entries(previous.spans || {}).forEach(([spanId, oldSpan]) => {
      const mappedOldSpan = remapSpanEndpoints(oldSpan, poleAliases);
      if (merged.spans?.[spanId]) {
        merged.spans[spanId] = preserveValuesMissingFromUpdate(merged.spans[spanId], mappedOldSpan, reconciliation);
      }
    });

    Object.values(previous.spanSides || {}).forEach(oldSideSource => {
      const oldSide = { ...oldSideSource, poleId: remapPoleId(oldSideSource.poleId, poleAliases) };
      const key = S.keyForSpanSide(oldSide.spanId, oldSide.poleId);
      const newSide = merged.spanSides?.[key];
      if (newSide) {
        const preservedSide = preserveValuesMissingFromUpdate(newSide, oldSide, reconciliation);
        merged.spanSides[key] = {
          ...preservedSide,
          proposedHOA: oldSide.proposedHOA || preservedSide.proposedHOA || "",
          proposedHOAChange: oldSide.proposedHOAChange || preservedSide.proposedHOAChange || "",
          nextPoleProposedAuto: oldSide.proposedHOAChange ? oldSide.nextPoleProposedAuto : preservedSide.nextPoleProposedAuto,
          proposedMidspan: oldSide.proposedMidspan || preservedSide.proposedMidspan || "",
          ocalcMS: oldSide.ocalcMS || preservedSide.ocalcMS || "",
          notes: oldSide.notes || preservedSide.notes || "",
          isSlack: Boolean(oldSide.isSlack || preservedSide.isSlack),
          isManualProposed: Boolean(oldSide.isManualProposed || preservedSide.isManualProposed),
          isAdditionalProposed: Boolean(oldSide.isAdditionalProposed || preservedSide.isAdditionalProposed)
        };
        return;
      }
      if (!spanSideHasUserWork(oldSide)) return;
      merged.spanSides[key] = oldSide;
      reconciliation.missingRowsPreserved += 1;
      if (!merged.spans[oldSide.spanId] && previous.spans?.[oldSide.spanId]) {
        merged.spans[oldSide.spanId] = remapSpanEndpoints(previous.spans[oldSide.spanId], poleAliases);
      }
    });

    const claimedImportedCommKeys = new Set();
    Object.values(previous.spanComms || {}).forEach(oldRowSource => {
      const oldRow = {
        ...oldRowSource,
        poleId: remapPoleId(oldRowSource.poleId, poleAliases),
        remotePoleId: remapPoleId(oldRowSource.remotePoleId, poleAliases)
      };
      const key = S.keyForSpanComm(oldRow.spanId, oldRow.poleId, oldRow.owner, oldRow.wireId || "");
      const match = findImportedCommMatch(merged.spanComms || {}, key, oldRow, claimedImportedCommKeys);
      if (match.row) {
        if (match.matchType === "exact") reconciliation.exactCommMatches += 1;
        else reconciliation.logicalCommMatches += 1;
        claimedImportedCommKeys.add(match.key);
        const preservedRow = preserveValuesMissingFromUpdate(match.row, oldRow, reconciliation);
        merged.spanComms[match.key] = spanCommHasUserWork(oldRow)
          ? mergeCommUserWork(preservedRow, oldRow)
          : preservedRow;
        return;
      }
      // A second old row with the same logical identity is stale data, not a
      // separate cable. Truly unmatched rows are retained so Update Data does
      // not silently delete a previously imported cable.
      if (match.hadCandidates && !String(oldRow.wireId || "").startsWith("manual-")) {
        reconciliation.staleDuplicateCommsDiscarded += 1;
        return;
      }
      const hasUserWork = spanCommHasUserWork(oldRow);
      const hasCalculationData = spanCommHasCalculationData(oldRow);
      if (!hasUserWork && !hasCalculationData) return;
      if (hasUserWork) reconciliation.unmatchedUserCommsPreserved += 1;
      else reconciliation.baselineCommRowsPreserved += 1;
      merged.spanComms[key] = oldRow;
      reconciliation.missingRowsPreserved += 1;
      if (!merged.spans[oldRow.spanId] && previous.spans?.[oldRow.spanId]) {
        merged.spans[oldRow.spanId] = remapSpanEndpoints(previous.spans[oldRow.spanId], poleAliases);
      }
    });

    const claimedImportedPowerKeys = new Set();
    Object.entries(previous.spanPower || {}).forEach(([key, oldRowSource]) => {
      const oldRow = { ...oldRowSource, poleId: remapPoleId(oldRowSource.poleId, poleAliases) };
      const match = findImportedPowerMatch(merged.spanPower, key, oldRow, claimedImportedPowerKeys);
      if (match.row) {
        claimedImportedPowerKeys.add(match.key);
        merged.spanPower[match.key] = preserveValuesMissingFromUpdate(match.row, oldRow, reconciliation);
        return;
      }
      // Update Data can be a partial workbook. If its Power sheet omits a
      // different pole, keep that imported power row instead of deleting it.
      // This mirrors the non-destructive merge used for comms and review rows.
      let preservedKey = key;
      if (oldRow.poleId !== oldRowSource.poleId) {
        preservedKey = [oldRow.spanId, oldRow.poleId, oldRow.wireId || oldRow.label || "preserved"].join("__");
      }
      while (merged.spanPower[preservedKey]) preservedKey = `${preservedKey}__preserved`;
      merged.spanPower[preservedKey] = oldRow;
      reconciliation.missingRowsPreserved += 1;
    });

    Object.entries(previous.poles || {}).forEach(([poleId, oldPole]) => {
      const targetPoleId = remapPoleId(poleId, poleAliases);
      const mappedOldPole = { ...oldPole, poleId: targetPoleId };
      if (!merged.poles[targetPoleId]) {
        merged.poles[targetPoleId] = mappedOldPole;
        reconciliation.missingRowsPreserved += 1;
        return;
      }
      const importedPole = merged.poles[targetPoleId];
      const preservedPole = preserveValuesMissingFromUpdate(importedPole, mappedOldPole, reconciliation);
      const importedEquipment = importedPole.metadata?.powerEquipment || [];
      const oldEquipment = mappedOldPole.metadata?.powerEquipment || [];
      const powerEquipment = mergePowerEquipmentUserWork(importedEquipment, oldEquipment, reconciliation);
      const lowPowerBaseline = importedPole.lowPower
        || mappedOldPole.metadata?.lowPowerBaseline
        || preservedPole.lowPower
        || "";
      merged.poles[targetPoleId] = {
        ...preservedPole,
        poleId: targetPoleId,
        ugActive: Boolean(mappedOldPole.ugActive || preservedPole.ugActive),
        pcoActive: Boolean(mappedOldPole.pcoActive || preservedPole.pcoActive),
        poleInsetActive: Boolean(mappedOldPole.poleInsetActive || preservedPole.poleInsetActive),
        poleInsetReason: mappedOldPole.poleInsetReason || preservedPole.poleInsetReason || "FAILING_CLEARANCES",
        poleInsetSpanId: mappedOldPole.poleInsetSpanId || preservedPole.poleInsetSpanId || "",
        redTagActive: Boolean(mappedOldPole.redTagActive || preservedPole.redTagActive),
        redTagMinHeight: mappedOldPole.redTagMinHeight || preservedPole.redTagMinHeight || "",
        redTagClass: mappedOldPole.redTagClass || preservedPole.redTagClass || "",
        redTagReason: mappedOldPole.redTagReason || preservedPole.redTagReason || "",
        lessThan12CommClearanceActive: Boolean(mappedOldPole.lessThan12CommClearanceActive || preservedPole.lessThan12CommClearanceActive),
        riserActive: mappedOldPole.riserActive === true || mappedOldPole.riserActive === false
          ? mappedOldPole.riserActive
          : preservedPole.riserActive,
        secondaryDripLoopActive: Boolean(mappedOldPole.secondaryDripLoopActive || preservedPole.secondaryDripLoopActive),
        secondaryDripLoopHoa: mappedOldPole.secondaryDripLoopHoa || preservedPole.secondaryDripLoopHoa || "",
        standaloneProposedHOA: mappedOldPole.standaloneProposedHOA || preservedPole.standaloneProposedHOA || "",
        notes: mappedOldPole.notes || preservedPole.notes || "",
        metadata: {
          ...(preservedPole.metadata || {}),
          lowPowerBaseline,
          powerEquipment
        }
      };
    });

    merged.makeReadyReferences = mergeMakeReadyReferences(
      (previous.makeReadyReferences || []).map(row => ({ ...row, poleId: remapPoleId(row.poleId, poleAliases) })),
      merged.makeReadyReferences || [],
      reconciliation
    );

    const importedPoleClassById = new Map((merged.poleClassChecks || []).map(row => [S.canonicalPoleIdentity(row.poleId), row]));
    (previous.poleClassChecks || []).forEach(oldRowSource => {
      const oldRow = { ...oldRowSource, poleId: remapPoleId(oldRowSource.poleId, poleAliases) };
      const identity = S.canonicalPoleIdentity(oldRow.poleId);
      const importedRow = importedPoleClassById.get(identity);
      if (!importedRow) {
        importedPoleClassById.set(identity, oldRow);
        reconciliation.missingRowsPreserved += 1;
        return;
      }
      const preserved = preserveValuesMissingFromUpdate(importedRow, oldRow, reconciliation);
      // Birthmark activation is an operator choice, not a workbook fact. Keep
      // an explicit manual on/off value when Update Data refreshes Collection.
      if (Object.prototype.hasOwnProperty.call(oldRow, "birthmarkActive")) {
        preserved.birthmarkActive = Boolean(oldRow.birthmarkActive);
      }
      importedPoleClassById.set(
        identity,
        global.ExcelImport?.recalculatePoleClassCheck
          ? global.ExcelImport.recalculatePoleClassCheck(preserved)
          : preserved
      );
    });
    merged.poleClassChecks = Array.from(importedPoleClassById.values());
    merged.excelReviewSource = mergeExcelReviewSources(
      previous.excelReviewSource || {},
      merged.excelReviewSource || {},
      reconciliation
    );

    merged.settings = {
      ...(merged.settings || {}),
      ...(previous.settings || {}),
      fiberSizes: {
        ...(merged.settings?.fiberSizes || {}),
        ...(previous.settings?.fiberSizes || {})
      }
    };
    // Review decisions belong to the job, not to one workbook snapshot.
    merged.excelReviewIgnoredChecks = { ...(previous.excelReviewIgnoredChecks || {}) };
    merged.deletedPoleIds = Array.from(new Set(previous.deletedPoleIds || []));
    reconciliation.deletedPolesSuppressed = suppressDeletedPoles(merged, merged.deletedPoleIds);
    merged.updateDiagnostics = {
      updatedAt: new Date().toISOString(),
      ...reconciliation
    };
    merged.ui = {
      ...(merged.ui || {}),
      ...(previous.ui || {}),
      hiddenPoleIds: (previous.ui?.hiddenPoleIds || []).map(id => remapPoleId(id, poleAliases)),
      revealedUnknownPoleIds: (previous.ui?.revealedUnknownPoleIds || []).map(id => remapPoleId(id, poleAliases))
    };
    return merged;
  }

  function updateLogValue(value) {
    if (value === undefined) return "(missing)";
    if (value === null) return "null";
    if (typeof value === "object") return JSON.stringify(value);
    return String(value);
  }

  function collectUpdateMapChanges(previousMap, finalMap, entity, fields) {
    const changes = [];
    const keys = new Set([...Object.keys(previousMap || {}), ...Object.keys(finalMap || {})]);
    keys.forEach(key => {
      const before = previousMap?.[key];
      const after = finalMap?.[key];
      if (!before && after) {
        changes.push({ entity, key, action: "Added", field: "(row)", before: "", after: "Present" });
        return;
      }
      if (before && !after) {
        changes.push({ entity, key, action: "Removed", field: "(row)", before: "Present", after: "" });
        return;
      }
      fields.forEach(field => {
        const beforeValue = updateLogValue(before?.[field]);
        const afterValue = updateLogValue(after?.[field]);
        if (beforeValue === afterValue) return;
        changes.push({ entity, key, action: "Changed", field, before: beforeValue, after: afterValue });
      });
    });
    return changes;
  }

  // Update Data is intentionally observable in DevTools. The summary explains
  // reconciliation decisions; the table lists final field-level state changes
  // after aliases, preserved baselines and recalculated values are applied.
  function logExcelUpdateChanges(fileName, previous, finalState) {
    const specs = [
      ["Pole", "poles", ["poleHeight", "lowPower", "poleType", "standaloneProposedHOA", "ugActive", "ugSpanId", "ugSpanMRSpanId", "ugSpanMRText", "ugReason", "ugMRText", "pcoActive", "pcoMRText", "riserActive", "secondaryDripLoopActive", "secondaryDripLoopHoa", "ugRiserDirection", "ancActive", "ancSize", "ancDistance", "ancDistanceManual", "ancDirection", "ancDirectionManual", "dgHoa", "ohgActive", "ohgHoa", "ohgDirection", "redTagActive", "redTagMinHeight", "redTagClass", "redTagReason"]],
      ["Span", "spans", ["fromPole", "toPole", "type", "direction", "bearingDegrees", "lengthDisplay", "environment"]],
      ["Proposed", "spanSides", ["proposedHOA", "proposedHOAChange", "isSlack", "endDrop", "ocalcMS", "msProposed", "finalMidspan", "clearanceMSStatus", "proposedFlaggingStatus"]],
      ["Comm", "spanComms", ["owner", "existingHOA", "existingHOAChange", "otherHOA", "otherHOAActive", "remoteHOA", "midspan", "calculatedMidspan", "finalMidspan", "flaggingStatus", "serviceDrop", "downGuy", "transferToNewPole", "resagServiceDrop", "isEndpointPlaceholder"]],
      ["Power", "spanPower", ["attachmentHeight", "attachmentHeightChange", "midspan", "calculatedMidspan", "owner", "size"]]
    ];
    const changes = specs.flatMap(([entity, mapName, fields]) =>
      collectUpdateMapChanges(previous?.[mapName], finalState?.[mapName], entity, fields)
    );
    const counts = mapName => ({
      before: Object.keys(previous?.[mapName] || {}).length,
      after: Object.keys(finalState?.[mapName] || {}).length
    });

    console.groupCollapsed(`[PoleCalc Update Data] ${fileName || "Excel update"}`);
    console.info("Reconciliation summary", finalState.updateDiagnostics || {});
    console.table({
      Poles: counts("poles"),
      Spans: counts("spans"),
      Proposed: counts("spanSides"),
      Comms: counts("spanComms"),
      Power: counts("spanPower")
    });
    if (changes.length) console.table(changes);
    else console.info("No final state fields changed.");
    console.groupEnd();
  }

    return { mergeImportedUpdate, logExcelUpdateChanges };
  }

  global.UpdateDataReconciler = { create };
})(window);
