(function (global) {
  "use strict";

  // ExcelReview audits the imported workbook against the current calculator
  // state. It never mutates calculator entities and deliberately excludes the
  // clearance and Pole Type rules owned by their existing modules.
  const S = () => global.AppStore;
  const H = () => global.HeightUtils;
  const I = () => global.ExcelImport;

  const STATUS_PRIORITY = { ERROR: 0, WARNING: 1, PASS: 2, NOT_READY: 3 };
  const ALLOWED_COMM_OWNERS = new Set([
    "communication",
    "3j communications",
    "cable one, prescott",
    "cable one, show low",
    "century link communications",
    "commnet",
    "cox communications",
    "mci metro",
    "wecom inc",
    "catv",
    "telco",
    "fiber",
    "cable one"
  ]);
  const COMM_INSULATORS = ["single bolt", "three bolt", "j-hook"];
  const PRIMARY_INSULATORS = ["deadend aps", "horizontal line post", "post 15\""];
  const SECONDARY_INSULATORS = ["pin 7.5in", "deadend aps", "spool 3in", "suspension aps"];

  let current = emptyResults();

  function emptyResults() {
    return {
      reviewedAt: "",
      results: [],
      globalChecks: [],
      summary: { errors: 0, warnings: 0, passed: 0, finalNotReady: 0, total: 0 }
    };
  }

  function text(value) {
    return String(value ?? "").trim();
  }

  function normalizedText(value) {
    return text(value)
      .toLowerCase()
      .normalize("NFD")
      .replace(/[\u0300-\u036f]/g, "")
      .replace(/\s+/g, " ");
  }

  function normalizedOwner(value) {
    const clean = normalizedText(value).replace(/^communication\s*>\s*/i, "").trim();
    if (/century\s*link|centurylink|\bctl\b/.test(clean)) return "ctl";
    if (/cable\s*one/.test(clean)) return clean.includes("show low") ? "cable one show low" : "cable one";
    if (/cox/.test(clean)) return "cox";
    if (/\bcatv\b/.test(clean)) return "catv";
    if (/\btelco\b/.test(clean)) return "telco";
    if (/fiber/.test(clean)) return "fiber";
    return clean;
  }

  function normalizeSpanType(value) {
    const clean = normalizedText(value).replace(/[^a-z]/g, "");
    if (clean === "forespan") return "FORE";
    if (clean === "backspan") return "BACK";
    if (clean === "other" || clean === "otherspan") return "OTHER";
    return clean.toUpperCase();
  }

  function check(data) {
    return {
      poleId: text(data.poleId),
      phase: data.phase === "FINAL" ? "FINAL" : "HOA",
      section: text(data.section),
      code: text(data.code),
      status: ["PASS", "WARNING", "ERROR", "NOT_READY"].includes(data.status) ? data.status : "PASS",
      level: data.level === "low" ? "low" : "normal",
      title: text(data.title),
      message: text(data.message),
      expected: text(data.expected),
      actual: text(data.actual),
      applicable: data.applicable !== false,
      details: Array.isArray(data.details) ? data.details : []
    };
  }

  function add(result, data) {
    result.checks.push(check({ poleId: result.poleId, ...data }));
  }

  function source() {
    return S().getState().excelReviewSource || {};
  }

  function rows(sheetName) {
    return Array.isArray(source()[sheetName]?.rows) ? source()[sheetName].rows : [];
  }

  function headers(sheetName) {
    return Array.isArray(source()[sheetName]?.headers) ? source()[sheetName].headers : [];
  }

  function pick(row, names, options) {
    return I().pick(row || {}, names, options || {});
  }

  function hasHeader(sheetName, names) {
    const wanted = names.map(I().normalizeHeaderName);
    return headers(sheetName).some(header => wanted.includes(I().normalizeHeaderName(header)));
  }

  function exactDisplayLowPower(row) {
    const key = Object.keys(row || {}).find(header => I().normalizeHeaderName(header) === "lowpowerattachmentdisplay");
    return key ? row[key] : "";
  }

  function collectionPoleId(row) {
    return text(pick(row, ["Id", "Pole ID", "PoleId", "PoleName", "Structure Number", "Pole"]));
  }

  function collectionSequence(row) {
    return text(pick(row, ["Sequence", "Seq"]));
  }

  function collectionEntries() {
    return rows("collection").map((row, index) => {
      const poleId = collectionPoleId(row);
      return {
        row,
        sourceRow: index + 2,
        poleId,
        displayPoleId: poleId || `Collection row ${index + 2}`,
        sequence: collectionSequence(row),
        collectionId: text(pick(row, ["collectionId", "Collection ID"]))
      };
    });
  }

  function collectionMaps(entries) {
    const byPole = new Map();
    const byCollection = new Map();
    entries.forEach(entry => {
      if (entry.poleId) byPole.set(normalizedText(entry.poleId), entry.poleId);
      if (entry.collectionId && entry.poleId) byCollection.set(normalizedText(entry.collectionId), entry.poleId);
    });
    return { byPole, byCollection };
  }

  function resolvePole(value, maps) {
    const candidate = normalizedText(value);
    if (!candidate) return "";
    if (maps.byPole.has(candidate)) return maps.byPole.get(candidate);
    if (maps.byCollection.has(candidate)) return maps.byCollection.get(candidate);
    for (const [key, poleId] of maps.byPole.entries()) {
      if (candidate.includes(key) || key.includes(candidate)) return poleId;
    }
    return "";
  }

  function spanModels(entries, maps) {
    return rows("spans").map((row, index) => {
      const sourceId = text(pick(row, ["Id", "Pole ID", "Pole", "From Pole"]));
      const sourceCollection = text(pick(row, ["collectionId", "Collection ID"]));
      const linkedTitle = text(pick(row, ["Linked Collection.Title", "Linked Collection Title", "Other Pole", "To Pole", "Remote Pole"]));
      const linkedCollectionId = text(pick(row, ["Linked Collection.ID", "Linked Collection ID"]));
      const fromPole = resolvePole(sourceId, maps) || maps.byCollection.get(normalizedText(sourceCollection)) || sourceId;
      const toPole = resolvePole(linkedTitle, maps) || maps.byCollection.get(normalizedText(linkedCollectionId)) || "";
      return {
        row,
        sourceRow: index + 2,
        spanId: text(pick(row, ["Span Id", "Span ID", "spanId", "Wire Span ID"])),
        spanIndex: text(pick(row, ["Span Index"])),
        typeRaw: text(pick(row, ["Type"])),
        type: normalizeSpanType(pick(row, ["Type"])),
        fromPole,
        toPole,
        linkedTitle,
        linkedCollectionId,
        environment: normalizedText(pick(row, ["Environment"])).replace(/[^a-z0-9]+/g, "_"),
        hasConnectionData: Boolean(linkedTitle || linkedCollectionId),
        details: {
          spanIndex: text(pick(row, ["Span Index"])),
          spanId: text(pick(row, ["Span Id", "Span ID", "spanId", "Wire Span ID"])),
          type: text(pick(row, ["Type"])),
          linkedCollectionTitle: linkedTitle
        }
      };
    }).filter(span => span.fromPole);
  }

  function spanDescription(span) {
    const label = span.spanIndex ? `Span ${span.spanIndex}` : (span.spanId ? `Span ${span.spanId}` : `source row ${span.sourceRow}`);
    return `${label}${span.toPole ? ` to ${span.toPole}` : ""}`;
  }

  function rowsForPole(sheetName, poleId) {
    const wanted = normalizedText(poleId);
    return rows(sheetName).filter(row => normalizedText(pick(row, ["Id", "Pole ID", "PoleId", "Pole", "Structure Number"])) === wanted);
  }

  function statusFromChecks(checks, phase) {
    const applicable = checks.filter(item => item.phase === phase && item.applicable !== false);
    if (applicable.some(item => item.status === "ERROR")) return "ERROR";
    if (applicable.some(item => item.status === "WARNING")) return "WARNING";
    if (applicable.some(item => item.status === "NOT_READY")) return "NOT_READY";
    return "PASS";
  }

  function addCollectionChecks(result, entry) {
    if (!entry.poleId) {
      add(result, {
        phase: "HOA", section: "Collection", code: "MISSING_POLE_ID", status: "ERROR",
        title: "Pole Id", message: `Id is empty on Collection row ${entry.sourceRow}.`, expected: "Non-empty Id", actual: "Empty"
      });
    }

    if (!entry.sequence) {
      add(result, {
        phase: "HOA", section: "Collection", code: "MISSING_SEQUENCE", status: "ERROR",
        title: "Sequence", message: "Sequence is empty.", expected: "Sequence matching the start of Id", actual: "Empty"
      });
    } else if (entry.poleId && !normalizedText(entry.poleId).startsWith(normalizedText(entry.sequence))) {
      add(result, {
        phase: "HOA", section: "Collection", code: "SEQUENCE_ID_MISMATCH", status: "ERROR",
        title: "Sequence", message: `Sequence ${entry.sequence} does not match the start of Id ${entry.poleId}.`,
        expected: `${entry.sequence}...`, actual: entry.poleId
      });
    }

    const lowPower = exactDisplayLowPower(entry.row);
    if (!text(lowPower)) {
      add(result, {
        phase: "HOA", section: "Collection", code: "MISSING_LOW_POWER", status: "ERROR",
        title: "Low Power", message: "Low Power Attachment.display is empty.", expected: "Non-empty value", actual: "Empty"
      });
    }

    const year = pick(entry.row, ["Year Installed"], { contains: false });
    if (!text(year)) {
      add(result, {
        phase: "HOA", section: "Collection", code: "MISSING_YEAR_INSTALLED", status: "WARNING",
        title: "Year Installed", message: "Year Installed is missing. Review the required project loading manually.",
        expected: "Non-empty value", actual: "Empty"
      });
    }
  }

  function addSpanCountChecks(result, poleSpans) {
    [["FORE", "Fore Span"], ["BACK", "Back Span"]].forEach(([type, label]) => {
      const found = poleSpans.filter(span => span.type === type);
      if (found.length === 1) return;
      const code = found.length ? `MULTIPLE_${type}_SPANS` : `MISSING_${type}_SPAN`;
      const listed = found.map(span => span.spanIndex || span.spanId || `row ${span.sourceRow}`).join(", ");
      add(result, {
        phase: "HOA", section: "Span", code, status: "ERROR", title: label,
        message: `Expected exactly one ${label}. Found ${found.length}${listed ? `: ${listed}` : ""}.`,
        expected: "1", actual: String(found.length), details: found.map(span => span.details)
      });
    });
  }

  function addLinkedCollectionChecks(result, poleSpans, maps) {
    poleSpans.forEach(span => {
      if (!span.linkedTitle) {
        add(result, {
          phase: "HOA", section: "Span", code: "MISSING_LINKED_COLLECTION_TITLE", status: "WARNING", level: "low",
          title: "Linked Collection", message: "Linked Collection.Title is empty. Review the span connection.",
          expected: "Existing Collection Id", actual: "Empty", details: [span.details]
        });
        return;
      }
      if (!resolvePole(span.linkedTitle, maps)) {
        add(result, {
          phase: "HOA", section: "Span", code: "UNKNOWN_LINKED_COLLECTION", status: "WARNING", level: "low",
          title: "Linked Collection", message: `Linked Collection.Title ${span.linkedTitle} does not match an Id in Collection.`,
          expected: "Existing Collection Id", actual: span.linkedTitle, details: [span.details]
        });
      }
    });
  }

  function addReciprocalChecks(result, poleSpans, allSpans) {
    poleSpans.filter(span => span.type === "FORE" || span.type === "BACK").forEach(span => {
      if (!span.hasConnectionData || !span.toPole) return;
      const expectedType = span.type === "FORE" ? "BACK" : "FORE";
      const reverse = allSpans.filter(candidate => (
        normalizedText(candidate.fromPole) === normalizedText(span.toPole)
        && normalizedText(candidate.toPole) === normalizedText(span.fromPole)
        && candidate.type === expectedType
      ));
      if (reverse.length === 1) return;
      const wrongReverse = allSpans.filter(candidate => (
        normalizedText(candidate.fromPole) === normalizedText(span.toPole)
        && normalizedText(candidate.toPole) === normalizedText(span.fromPole)
        && candidate.type !== expectedType
      ));
      const expectedLabel = expectedType === "BACK" ? "Back Span" : "Fore Span";
      const actual = wrongReverse.length
        ? wrongReverse.map(candidate => candidate.typeRaw || candidate.type).join(", ")
        : "No reciprocal span";
      add(result, {
        phase: "HOA", section: "Span", code: "RECIPROCAL_SPAN_WARNING", status: "WARNING",
        title: "Fore / Back Relationship",
        message: `${spanDescription(span)} expects exactly one ${expectedLabel} from ${span.toPole} back to ${span.fromPole}. Found ${reverse.length}.`,
        expected: expectedLabel, actual, details: [span.details, ...reverse.map(candidate => candidate.details), ...wrongReverse.map(candidate => candidate.details)]
      });
    });
  }

  function addEnvironmentChecks(result, poleSpans, allSpans, environmentPairsSeen) {
    poleSpans.forEach(span => {
      if (!span.toPole || !span.environment) return;
      const reverse = allSpans.filter(candidate => (
        normalizedText(candidate.fromPole) === normalizedText(span.toPole)
        && normalizedText(candidate.toPole) === normalizedText(span.fromPole)
        && ((span.type === "FORE" && candidate.type === "BACK") || (span.type === "BACK" && candidate.type === "FORE") || (span.type === "OTHER" && candidate.type === "OTHER"))
      ));
      reverse.forEach(candidate => {
        if (!candidate.environment) return;
        const pairKey = `${normalizedText(result.sourcePoleId)}|${[span.fromPole, span.toPole].map(normalizedText).sort().join("|")}|${span.type}|${candidate.type}`;
        if (environmentPairsSeen.has(pairKey)) return;
        environmentPairsSeen.add(pairKey);
        if (span.environment === candidate.environment) return;
        add(result, {
          phase: "HOA", section: "Span", code: "ENVIRONMENT_MISMATCH", status: "ERROR",
          title: "Environment",
          message: `Environment mismatch on ${span.fromPole} to ${span.toPole}: ${span.environment || "empty"} vs ${candidate.environment || "empty"}.`,
          expected: span.environment, actual: candidate.environment, details: [span.details, candidate.details]
        });
      });
    });
  }

  function normalizedInsulator(value) {
    return normalizedText(value).replace(/\s+/g, " ").replace(/\.+$/g, "");
  }

  function addIntecWireChecks(result, poleId) {
    if (text(S().getState().settings?.projectProfile).toUpperCase() !== "INTEC") return;
    rowsForPole("spanWires", poleId).forEach((row, index) => {
      const ownerRaw = text(pick(row, ["Owner", "owner"]));
      const owner = normalizedText(ownerRaw).replace(/^communication\s*>\s*/i, "").trim();
      const size = text(pick(row, ["Size", "Size.display", "Wire Size"]));
      const construction = text(pick(row, ["Construction"]));
      const insulator = normalizedInsulator(pick(row, ["Insulator"]));
      const descriptor = `${ownerRaw || "No owner"} / ${size || `row ${index + 2}`}`;
      const power = I().isPowerWire ? I().isPowerWire(row) : /^utility\s*>/i.test(ownerRaw);
      const communication = I().isCommunicationWire ? I().isCommunicationWire(row) : !power;

      if (normalizedText(construction).replace(/\s+/g, "") === "davit") {
        add(result, {
          phase: "HOA", section: "Span.Wire", code: "DAVIT_NOT_ALLOWED", status: "ERROR",
          title: "Construction", message: "Construction DAVIT is not allowed for INTEC.", expected: "Not DAVIT", actual: construction
        });
      }

      if (communication && !ALLOWED_COMM_OWNERS.has(owner)) {
        add(result, {
          phase: "HOA", section: "Span.Wire", code: "UNKNOWN_COMM_OWNER", status: "WARNING",
          title: "Communication Owner", message: `Communication owner is not in the INTEC list: ${ownerRaw || "empty"}.`,
          expected: "Allowed INTEC communication owner", actual: ownerRaw
        });
      }

      if (communication && !COMM_INSULATORS.includes(insulator)) {
        add(result, {
          phase: "HOA", section: "Span.Wire", code: "INVALID_COMM_INSULATOR", status: "ERROR",
          title: "Communication Insulator", message: `Invalid communication insulator for ${descriptor}.`,
          expected: COMM_INSULATORS.join(", "), actual: text(pick(row, ["Insulator"])) || "Empty"
        });
      }

      if (!power) return;
      const powerType = /primary/i.test(size) ? "PRIMARY" : /secondary/i.test(size) ? "SECONDARY" : /neutral/i.test(size) ? "NEUTRAL" : "";
      if (!powerType) return;
      if (normalizedText(ownerRaw) !== "utility > aps (joint use)") {
        add(result, {
          phase: "HOA", section: "Span.Wire", code: "INVALID_POWER_OWNER", status: "ERROR",
          title: `${powerType} Owner`, message: `${powerType} owner must be UTILITY > APS (Joint use).`,
          expected: "UTILITY > APS (Joint use)", actual: ownerRaw || "Empty"
        });
      }
      const validInsulator = powerType === "PRIMARY"
        ? (insulator.includes("pin 8.38") || PRIMARY_INSULATORS.includes(insulator))
        : SECONDARY_INSULATORS.includes(insulator);
      if (!validInsulator) {
        add(result, {
          phase: "HOA", section: "Span.Wire", code: `INVALID_${powerType}_INSULATOR`, status: "ERROR",
          title: `${powerType} Insulator`, message: `Invalid ${powerType.toLowerCase()} insulator for ${descriptor}.`,
          expected: powerType === "PRIMARY" ? "Pin 8.38, Deadend APS, Horizontal Line Post, or Post 15\"" : SECONDARY_INSULATORS.join(", "),
          actual: text(pick(row, ["Insulator"])) || "Empty"
        });
      }
    });
  }

  function calculatorWorkForPole(poleId) {
    const state = S().getState();
    const pole = state.poles[poleId];
    const commWork = S().getSpanCommsForPole(poleId).some(row => text(row.existingHOAChange));
    const proposedWork = S().getSpanSidesForPole(poleId).some(side => [
      side.proposedHOA, side.proposedHOAChange, side.ocalcMS, side.proposedMidspan,
      side.msProposed, side.finalMidspan, side.endDrop
    ].some(value => text(value)));
    const standalone = text(pole?.standaloneProposedHOA);
    // Legacy or imported state can retain generated instructions without every
    // originating edit field. A non-empty generated MR is still final work.
    const generatedMR = text((state.mr || []).find(item => normalizedText(item.poleId) === normalizedText(poleId))?.text);
    return Boolean(commWork || proposedWork || standalone || pole?.ugActive || generatedMR);
  }

  function excelFinalRowsForPole(poleId) {
    return rowsForPole("makeReady", poleId).filter(row => [
      pick(row, ["Attachment Size"], { contains: true }),
      pick(row, ["Attachment Type"], { contains: true }),
      pick(row, ["Attachment Height.display", "Attachment Height Display"], { contains: true }),
      pick(row, ["Proposed Mid Span.display", "Proposed Midspan.display", "Proposed Mid Span Display"], { contains: true }),
      pick(row, ["Make Ready Notes", "MR Notes", "Notes"], { contains: true }),
      pick(row, ["Comm Transfers"], { contains: true })
    ].some(value => text(value)));
  }

  function excelTransferRowsForPole(poleId) {
    return rowsForPole("commTransfers", poleId).filter(row => [
      pick(row, ["Owner"]), pick(row, ["Height.display", "Height Display", "Height"], { contains: true })
    ].some(value => text(value)));
  }

  function excelHasFinalForPole(poleId) {
    return excelFinalRowsForPole(poleId).length > 0 || excelTransferRowsForPole(poleId).length > 0;
  }

  function poleById(poleId) {
    const wanted = normalizedText(poleId);
    return Object.values(S().getState().poles || {}).find(pole => normalizedText(pole.poleId) === wanted) || null;
  }

  function finalMidspan(side) {
    return side.finalMidspan || side.msProposed || side.proposedMidspan || side.ocalcMS || "";
  }

  function heightValue(row, displayNames, fallbackNames) {
    const display = pick(row, displayNames, { contains: true });
    return text(display) ? display : pick(row, fallbackNames || [], { contains: true });
  }

  function heightMatches(a, b) {
    const left = H().parseHeight(a);
    const right = H().parseHeight(b);
    return left !== null && right !== null && left === right;
  }

  function attachmentDirection(value) {
    const match = text(value).match(/\(([^)]+)\)\s*$/);
    return match ? match[1].toUpperCase().split(/[\s/,;]+/).filter(Boolean) : [];
  }

  function selectMakeReadyRow(expected, candidates, used) {
    const expectedHeight = H().parseHeight(expected.proposedHOA);
    return candidates
      .map((row, index) => ({ row, index }))
      .filter(item => !used.has(item.index))
      .map(item => {
        const rawSpanId = text(pick(item.row, ["Span Id", "Span ID", "spanId"]));
        const rawConnection = text(pick(item.row, ["Linked Collection.Title", "Other Pole", "To Pole"], { contains: true }));
        const directions = attachmentDirection(pick(item.row, ["Attachment Size"], { contains: true }));
        const actualHeight = H().parseHeight(heightValue(item.row, ["Attachment Height.display", "Attachment Height Display"], ["Attachment Height"]));
        let score = 0;
        if (rawSpanId && rawSpanId === expected.spanId) score += 100;
        if (rawConnection && normalizedText(rawConnection).includes(normalizedText(expected.otherPole))) score += 70;
        if (expected.direction && directions.includes(expected.direction.toUpperCase())) score += 40;
        if (expectedHeight !== null && actualHeight !== null) score -= Math.abs(expectedHeight - actualHeight) / 100;
        return { ...item, score };
      })
      .sort((a, b) => b.score - a.score)[0] || null;
  }

  function proposedExpectations(poleId) {
    const items = S().getSpanSidesForPole(poleId)
      .filter(side => text(side.proposedHOA))
      .map(side => {
        const span = S().getSpan(side.spanId);
        return {
          spanId: side.spanId,
          otherPole: span ? S().getOtherPoleId(span, poleId) : "",
          direction: span?.fromPole === poleId ? span.direction : oppositeDirection(span?.direction),
          proposedHOA: side.proposedHOA,
          proposedMidspan: finalMidspan(side)
        };
      });
    const standalone = poleById(poleId)?.standaloneProposedHOA;
    if (text(standalone)) items.push({ spanId: "", otherPole: "", direction: "", proposedHOA: standalone, proposedMidspan: "" });
    return items;
  }

  function oppositeDirection(direction) {
    return { N: "S", NE: "SW", E: "W", SE: "NW", S: "N", SW: "NE", W: "E", NW: "SE" }[text(direction).toUpperCase()] || text(direction);
  }

  function addProposedComparisons(result, poleId, makeReadyRows) {
    const expected = proposedExpectations(poleId);
    const used = new Set();
    expected.forEach(item => {
      const selected = selectMakeReadyRow(item, makeReadyRows, used);
      if (!selected) {
        add(result, {
          phase: "FINAL", section: "Make Ready", code: "MISSING_PROPOSED_ATTACHMENT", status: "ERROR",
          title: "Proposed HOA", message: `Expected Proposed HOA ${item.proposedHOA} was not found in Excel Make Ready.`,
          expected: item.proposedHOA, actual: "Missing"
        });
        return;
      }
      used.add(selected.index);
      const excelHOA = heightValue(selected.row, ["Attachment Height.display", "Attachment Height Display"], ["Attachment Height"]);
      if (!text(excelHOA) || !heightMatches(item.proposedHOA, excelHOA)) {
        add(result, {
          phase: "FINAL", section: "Make Ready", code: "PROPOSED_HOA_MISMATCH", status: "ERROR",
          title: "Proposed HOA", message: `Proposed HOA mismatch. Calculator: ${item.proposedHOA}. Excel: ${text(excelHOA) || "empty"}.`,
          expected: item.proposedHOA, actual: text(excelHOA) || "Empty"
        });
      }

      const excelMS = heightValue(selected.row, ["Proposed Mid Span.display", "Proposed Midspan.display", "Proposed Mid Span Display"], ["Proposed Mid Span", "Proposed Midspan"]);
      if (text(item.proposedMidspan) && (!text(excelMS) || !heightMatches(item.proposedMidspan, excelMS))) {
        add(result, {
          phase: "FINAL", section: "Make Ready", code: "PROPOSED_MIDSPAN_MISMATCH", status: "ERROR",
          title: "Proposed Midspan", message: `Proposed Midspan mismatch. Calculator: ${item.proposedMidspan}. Excel: ${text(excelMS) || "empty"}.`,
          expected: item.proposedMidspan, actual: text(excelMS) || "Empty"
        });
      } else if (!text(item.proposedMidspan) && text(excelMS)) {
        add(result, {
          phase: "FINAL", section: "Make Ready", code: "UNMATCHED_EXCEL_MIDSPAN", status: "WARNING",
          title: "Proposed Midspan", message: `Excel has Proposed Midspan ${excelMS}, but Calculator has no comparable final midspan.`,
          expected: "No comparable value", actual: excelMS
        });
      }

      const size = text(pick(selected.row, ["Attachment Size"], { contains: true }));
      const directions = attachmentDirection(size);
      if (item.direction && directions.length && !directions.includes(item.direction.toUpperCase())) {
        add(result, {
          phase: "FINAL", section: "Make Ready", code: "ATTACHMENT_DIRECTION_MISMATCH", status: "ERROR",
          title: "Attachment Size", message: `Attachment direction does not include ${item.direction}.`,
          expected: item.direction, actual: directions.join("/")
        });
      } else {
        add(result, {
          phase: "FINAL", section: "Make Ready", code: "ATTACHMENT_SIZE_NOT_APPLICABLE", status: "PASS", applicable: false,
          title: "Attachment Size", message: "Not applicable: Calculator does not store a reliable per-Proposed fiber and messenger description for exact Attachment Size comparison."
        });
      }
    });

    makeReadyRows.forEach((row, index) => {
      if (used.has(index)) return;
      const hoa = heightValue(row, ["Attachment Height.display", "Attachment Height Display"], ["Attachment Height"]);
      const ms = heightValue(row, ["Proposed Mid Span.display", "Proposed Midspan.display", "Proposed Mid Span Display"], ["Proposed Mid Span", "Proposed Midspan"]);
      if (!text(hoa) && !text(ms)) return;
      add(result, {
        phase: "FINAL", section: "Make Ready", code: "UNMATCHED_EXCEL_ATTACHMENT", status: "WARNING",
        title: "Excel Attachment", message: "Excel has an attachment or Proposed Midspan without a comparable Calculator Proposed.",
        expected: "Comparable Calculator Proposed", actual: [text(hoa), text(ms)].filter(Boolean).join(" / ")
      });
    });
  }

  function splitInstructions(value) {
    return text(value)
      .split(/\r?\n+|(?<=[.!?])\s+(?=[A-Za-z(])/)
      .map(line => line.trim())
      .filter(Boolean);
  }

  function instructionSignature(value, poleId) {
    const raw = text(value).replace(/[’‘]/g, "'").replace(/[“”]/g, "\"");
    const clean = normalizedText(raw).replace(/[^a-z0-9'\".\s/-]/g, " ").replace(/\s+/g, " ");
    const action = /re[-\s]?sag/.test(clean) ? "resag"
      : /relocate/.test(clean) ? "relocate"
        : /transfer/.test(clean) ? "transfer"
          : /lower/.test(clean) ? "lower"
            : /raise|raising/.test(clean) ? "raise"
              : /attach/.test(clean) ? "attach"
                : /underground|\bgoing ug\b|\bgo ug\b|\bug transfer\b/.test(clean) ? "ug"
                  : /anchor|\banc\b/.test(clean) ? "anchor"
                    : /riser/.test(clean) ? "riser"
                      : /slack/.test(clean) ? "slack" : "other";
    const heightTokens = raw.match(/\d+(?:\.\d+)?\s*'\s*\d*(?:\.\d+)?\s*\"?|\d+\.\d+/g) || [];
    const heights = heightTokens.map(token => H().parseHeight(token)).filter(value => value !== null);
    const ownerCandidates = S().getSpanCommsForPole(poleId)
      .flatMap(row => [row.rawOwner, row.ownerBase, row.owner])
      .concat([S().getState().settings?.proposedOwner, "CTL", "CATV", "TELCO", "FIBER"])
      .filter(Boolean);
    const ownerIsMentioned = owner => {
      if (owner === "ctl") return /century\s*link|centurylink|\bctl\b/.test(clean);
      if (owner === "cable one" || owner === "cable one show low") return /cable\s*one/.test(clean);
      if (owner === "cox") return /\bcox\b/.test(clean);
      if (owner === "catv") return /\bcatv\b/.test(clean);
      if (owner === "telco") return /\btelco\b/.test(clean);
      if (owner === "fiber") return /\bfiber\b/.test(clean);
      return clean.includes(owner.replace(/[^a-z0-9]/g, " "));
    };
    const owners = Array.from(new Set(ownerCandidates.map(normalizedOwner).filter(owner => owner && ownerIsMentioned(owner))));
    const directionMatch = clean.toUpperCase().match(/(?:^|\s)(NE|SE|SW|NW|N|E|S|W)(?:\s|$)/);
    return {
      raw,
      action,
      heights,
      owners,
      direction: directionMatch ? directionMatch[1] : "",
      dg: /\bdg\b|down\s*guy/.test(clean),
      slack: /slack/.test(clean),
      anchor: /anchor|\banc\b/.test(clean),
      riser: /riser/.test(clean)
    };
  }

  function instructionMatches(expected, actual) {
    if (expected.action !== actual.action) return false;
    if (expected.heights.some(height => !actual.heights.includes(height))) return false;
    if (expected.owners.some(owner => !actual.owners.includes(owner))) return false;
    if (expected.dg && !actual.dg) return false;
    if (expected.slack && !actual.slack) return false;
    if (expected.anchor && !actual.anchor) return false;
    if (expected.riser && !actual.riser) return false;
    if (expected.direction && ["ug", "anchor", "riser", "slack"].includes(expected.action) && expected.direction !== actual.direction) return false;
    return true;
  }

  function addMakeReadyNotesComparison(result, poleId, makeReadyRows) {
    const generated = S().getState().mr.find(item => normalizedText(item.poleId) === normalizedText(poleId));
    const expectedLines = splitInstructions(generated?.text || "");
    const actualLines = makeReadyRows.flatMap(row => splitInstructions(pick(row, ["Make Ready Notes", "MR Notes", "Notes"], { contains: true })));
    const used = new Set();

    expectedLines.forEach(line => {
      const expected = instructionSignature(line, poleId);
      const matchIndex = actualLines.findIndex((actualLine, index) => !used.has(index) && instructionMatches(expected, instructionSignature(actualLine, poleId)));
      if (matchIndex >= 0) {
        used.add(matchIndex);
        return;
      }
      add(result, {
        phase: "FINAL", section: "Make Ready", code: "MISSING_MR_INSTRUCTION", status: "ERROR",
        title: "Make Ready Notes", message: `Expected instruction not found: ${line}`,
        expected: line, actual: actualLines.join(" | ") || "Empty"
      });
    });

    actualLines.forEach((line, index) => {
      if (used.has(index)) return;
      add(result, {
        phase: "FINAL", section: "Make Ready", code: "ADDITIONAL_MR_INSTRUCTION", status: "WARNING",
        title: "Make Ready Notes", message: `Excel instruction has no Calculator equivalent: ${line}`,
        expected: "Calculator instruction", actual: line
      });
    });
  }

  function expectedTransfers(poleId) {
    const seen = new Set();
    return S().getSpanCommsForPole(poleId).filter(row => row.transferToNewPole && text(row.existingHOAChange)).reduce((items, row) => {
      const owner = normalizedOwner(row.rawOwner || row.ownerBase || row.owner);
      const height = H().parseHeight(row.existingHOAChange);
      const key = `${owner}|${height}`;
      if (!owner || height === null || seen.has(key)) return items;
      seen.add(key);
      items.push({ owner, ownerDisplay: text(row.rawOwner || row.ownerBase || row.owner), height, heightDisplay: row.existingHOAChange });
      return items;
    }, []);
  }

  function addTransferComparisons(result, poleId) {
    const expected = expectedTransfers(poleId);
    const sheetExists = headers("commTransfers").length > 0;
    if (!expected.length) {
      add(result, {
        phase: "FINAL", section: "Make Ready", code: "COMM_TRANSFERS_NOT_APPLICABLE", status: "PASS", applicable: false,
        title: "Comm Transfers", message: "Not applicable: this pole has no Calculator transfer."
      });
      return;
    }
    if (!sheetExists) {
      add(result, {
        phase: "FINAL", section: "Make Ready", code: "COMM_TRANSFERS_NOT_COMPARABLE", status: "PASS", applicable: false,
        title: "Comm Transfers", message: "Not applicable: Make Ready.Comm Transfers is not available for structured comparison."
      });
      return;
    }
    const actual = excelTransferRowsForPole(poleId).map(row => ({
      owner: normalizedOwner(pick(row, ["Owner"])),
      ownerDisplay: text(pick(row, ["Owner"])),
      height: H().parseHeight(heightValue(row, ["Height.display", "Height Display"], ["Height"])),
      heightDisplay: text(heightValue(row, ["Height.display", "Height Display"], ["Height"]))
    }));
    const used = new Set();
    expected.forEach(item => {
      const exact = actual.findIndex((candidate, index) => !used.has(index) && candidate.owner === item.owner && candidate.height === item.height);
      if (exact >= 0) {
        used.add(exact);
        return;
      }
      const sameOwner = actual.findIndex((candidate, index) => !used.has(index) && candidate.owner === item.owner);
      add(result, {
        phase: "FINAL", section: "Make Ready", code: sameOwner >= 0 ? "TRANSFER_HEIGHT_MISMATCH" : "MISSING_COMM_TRANSFER", status: "ERROR",
        title: "Comm Transfers",
        message: sameOwner >= 0
          ? `Transfer height mismatch for ${item.ownerDisplay}. Calculator: ${item.heightDisplay}. Excel: ${actual[sameOwner].heightDisplay || "empty"}.`
          : `Expected transfer not found for ${item.ownerDisplay} at ${item.heightDisplay}.`,
        expected: `${item.ownerDisplay} ${item.heightDisplay}`,
        actual: sameOwner >= 0 ? `${actual[sameOwner].ownerDisplay} ${actual[sameOwner].heightDisplay}` : "Missing"
      });
      if (sameOwner >= 0) used.add(sameOwner);
    });
    actual.forEach((item, index) => {
      if (used.has(index)) return;
      add(result, {
        phase: "FINAL", section: "Make Ready", code: "UNEXPECTED_COMM_TRANSFER", status: "WARNING",
        title: "Comm Transfers", message: `Excel transfer has no Calculator equivalent: ${item.ownerDisplay} ${item.heightDisplay}.`,
        expected: "No additional transfer", actual: `${item.ownerDisplay} ${item.heightDisplay}`
      });
    });
  }

  function addFinalChecks(result, entry) {
    if (!entry.poleId) {
      add(result, {
        phase: "FINAL", section: "Make Ready", code: "FINAL_NOT_READY", status: "NOT_READY",
        title: "Final Review", message: "Final review cannot match a Collection row without Id."
      });
      return;
    }
    const calculatorHasFinal = calculatorWorkForPole(entry.poleId);
    const excelHasFinal = excelHasFinalForPole(entry.poleId);
    if (!calculatorHasFinal && !excelHasFinal) {
      add(result, {
        phase: "FINAL", section: "Make Ready", code: "FINAL_NOT_READY", status: "NOT_READY",
        title: "Final Review", message: "Final Review is not ready because neither Calculator nor Excel contains final work."
      });
      return;
    }
    if (calculatorHasFinal && !excelHasFinal) {
      add(result, {
        phase: "FINAL", section: "Make Ready", code: "CALCULATOR_WORK_EXCEL_EMPTY", status: "ERROR",
        title: "Final Review", message: "Calculator has final work, but Excel Make Ready is empty.",
        expected: "Excel Make Ready data", actual: "Empty"
      });
    }
    if (excelHasFinal && !calculatorHasFinal) {
      add(result, {
        phase: "FINAL", section: "Make Ready", code: "EXCEL_WORK_CALCULATOR_EMPTY", status: "WARNING",
        title: "Final Review", message: "Excel has final Make Ready data, but there is no calculator work available for comparison.",
        expected: "Calculator work", actual: "Excel only"
      });
    }

    const pla = text(pick(entry.row, ["PLA STATUS"], { contains: false }));
    if (!pla) {
      add(result, {
        phase: "FINAL", section: "PLA", code: "MISSING_PLA_STATUS", status: "ERROR",
        title: "PLA STATUS", message: "PLA STATUS is empty.", expected: "Non-empty value", actual: "Empty"
      });
    }

    const construction = text(pick(entry.row, ["MRE Construction Type"], { contains: false }));
    const normalizedConstruction = normalizedText(construction);
    if (!construction) {
      add(result, {
        phase: "FINAL", section: "Collection", code: "MISSING_MRE_CONSTRUCTION_TYPE", status: "ERROR",
        title: "MRE Construction Type", message: "MRE Construction Type is empty.", expected: "Aerial or Underground", actual: "Empty"
      });
    } else if (!["aerial", "underground"].includes(normalizedConstruction)) {
      add(result, {
        phase: "FINAL", section: "Collection", code: "UNKNOWN_MRE_CONSTRUCTION_TYPE", status: "WARNING",
        title: "MRE Construction Type", message: `Unexpected MRE Construction Type: ${construction}.`,
        expected: "Aerial or Underground", actual: construction
      });
    }

    const pole = poleById(entry.poleId);
    const expectedConstruction = pole?.ugActive ? "underground" : calculatorHasFinal ? "aerial" : "";
    if (expectedConstruction && normalizedConstruction && normalizedConstruction !== expectedConstruction) {
      add(result, {
        phase: "FINAL", section: "Collection", code: "MRE_CONSTRUCTION_MISMATCH", status: "ERROR",
        title: "MRE Construction Type",
        message: `Calculator solution is ${expectedConstruction === "underground" ? "Underground" : "Aerial"}, but Excel says ${construction}.`,
        expected: expectedConstruction === "underground" ? "Underground" : "Aerial", actual: construction
      });
    }

    if (!calculatorHasFinal || !excelHasFinal) return;
    const makeReadyRows = excelFinalRowsForPole(entry.poleId);
    if (pole?.ugActive) {
      const notes = makeReadyRows.map(row => text(pick(row, ["Make Ready Notes", "MR Notes", "Notes"], { contains: true }))).join("\n");
      if (!/underground|\bgoing\s+ug\b|\bgo\s+ug\b|\bug\s+transfer\b/i.test(notes)) {
        add(result, {
          phase: "FINAL", section: "Make Ready", code: "MISSING_UG_INSTRUCTION", status: "ERROR",
          title: "Underground Make Ready", message: "Calculator is UG, but Excel Make Ready Notes do not contain a UG instruction.",
          expected: "UG instruction", actual: notes || "Empty"
        });
      }
      addMakeReadyNotesComparison(result, entry.poleId, makeReadyRows);
      addTransferComparisons(result, entry.poleId);
      return;
    }

    addProposedComparisons(result, entry.poleId, makeReadyRows);
    addMakeReadyNotesComparison(result, entry.poleId, makeReadyRows);
    addTransferComparisons(result, entry.poleId);
  }

  function finalizeResult(result) {
    result.hoaStatus = statusFromChecks(result.checks, "HOA");
    result.finalStatus = statusFromChecks(result.checks, "FINAL");
    const statuses = [result.hoaStatus, result.finalStatus].filter(status => status !== "NOT_READY");
    result.overallStatus = statuses.includes("ERROR") ? "ERROR" : statuses.includes("WARNING") ? "WARNING" : "PASS";
    const applicableProblems = result.checks.filter(item => item.applicable !== false && ["ERROR", "WARNING"].includes(item.status));
    if (!applicableProblems.length) {
      result.checks.push(check({
        poleId: result.poleId, phase: "HOA", section: "Review", code: "ALL_APPLICABLE_CHECKS_PASSED", status: "PASS",
        title: "Review", message: "PASS — All applicable checks passed."
      }));
    }
    return result;
  }

  function naturalCompare(a, b) {
    return text(a).localeCompare(text(b), undefined, { numeric: true, sensitivity: "base" });
  }

  function sortResults(results) {
    return results.sort((a, b) => {
      const status = STATUS_PRIORITY[a.overallStatus] - STATUS_PRIORITY[b.overallStatus];
      if (status) return status;
      return naturalCompare(a.sequence || a.poleId, b.sequence || b.poleId);
    });
  }

  function buildSummary(results, globalChecks) {
    return {
      errors: results.filter(result => result.overallStatus === "ERROR").length + globalChecks.filter(item => item.status === "ERROR").length,
      warnings: results.filter(result => result.overallStatus === "WARNING").length + globalChecks.filter(item => item.status === "WARNING").length,
      passed: results.filter(result => result.overallStatus === "PASS").length,
      finalNotReady: results.filter(result => result.finalStatus === "NOT_READY").length,
      total: results.length
    };
  }

  function runReview() {
    current = emptyResults();
    const entries = collectionEntries();
    const maps = collectionMaps(entries);
    const spans = spanModels(entries, maps);
    const environmentPairsSeen = new Set();

    if (!hasHeader("collection", ["Id"])) {
      current.globalChecks.push(check({
        phase: "HOA", section: "Collection", code: "MISSING_ID_COLUMN", status: "ERROR",
        title: "Collection Id Column", message: "Collection is missing the Id column.", expected: "Id column", actual: "Missing"
      }));
    }

    current.results = entries.map(entry => {
      const result = {
        poleId: entry.displayPoleId,
        sourcePoleId: entry.poleId,
        sequence: entry.sequence,
        sourceRow: entry.sourceRow,
        hoaStatus: "PASS",
        finalStatus: "NOT_READY",
        overallStatus: "PASS",
        checks: []
      };
      const poleSpans = spans.filter(span => normalizedText(span.fromPole) === normalizedText(entry.poleId));
      addCollectionChecks(result, entry);
      addSpanCountChecks(result, poleSpans);
      addLinkedCollectionChecks(result, poleSpans, maps);
      addReciprocalChecks(result, poleSpans, spans);
      addEnvironmentChecks(result, poleSpans, spans, environmentPairsSeen);
      if (entry.poleId) addIntecWireChecks(result, entry.poleId);
      addFinalChecks(result, entry);
      return finalizeResult(result);
    });

    current.results = sortResults(current.results);
    current.reviewedAt = new Date().toISOString();
    current.summary = buildSummary(current.results, current.globalChecks);
    return current;
  }

  function reviewPole(poleId) {
    const wanted = normalizedText(poleId);
    return current.results.find(result => normalizedText(result.sourcePoleId || result.poleId) === wanted) || null;
  }

  function getResults() {
    return current.results;
  }

  function getSummary() {
    return current.summary;
  }

  function getReviewState() {
    return current;
  }

  function clearResults() {
    current = emptyResults();
  }

  /** @namespace ExcelReview */
  global.ExcelReview = {
    runReview,
    reviewPole,
    getResults,
    getSummary,
    getReviewState,
    clearResults
  };
})(window);
