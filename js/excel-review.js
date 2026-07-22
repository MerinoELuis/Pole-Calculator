(function (global) {
  "use strict";

  // ExcelReview audits the imported workbook against the current calculator
  // state. It never mutates calculator entities and deliberately excludes the
  // clearance and Pole Type rules owned by their existing modules.
  const S = () => global.AppStore;
  const H = () => global.HeightUtils;
  const I = () => global.ExcelImport;

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
  const MIDAM_POWER_SIZES = {
    PRIMARY: "primary > aac 477.0 kcm 19 strand cosmos > static",
    SECONDARY: "secondary > triplex 2 awg > static",
    NEUTRAL: "neutral > aac 477.0 kcm 19 strand cosmos > static"
  };
  const MIDAM_COMM_INSULATORS = ["single bolt", "three bolt"];
  const MIDAM_POWER_INSULATORS = ["spool 2.5\"", "deadend 12.75\"", "suspension 11.50\"", "pin 7.5\""];

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
    if (/century\s*link|centurylink|\bctl\b|\btelco\b/.test(clean)) return "ctl";
    if (/cable\s*one/.test(clean)) return clean.includes("show low") ? "cable one show low" : "cable one";
    if (/cox/.test(clean)) return "cox";
    if (/\bcatv\b/.test(clean)) return "catv";
    if (/\btelco\b/.test(clean)) return "telco";
    if (/fiber/.test(clean)) return "fiber";
    return clean;
  }

  function stableReviewHash(value) {
    let hash = 2166136261;
    const source = String(value || "");
    for (let index = 0; index < source.length; index += 1) {
      hash ^= source.charCodeAt(index);
      hash = Math.imul(hash, 16777619);
    }
    return (hash >>> 0).toString(36);
  }

  function checkIdentity(data) {
    const details = (Array.isArray(data.details) ? data.details : []).map(detail => [
      detail.spanIndex,
      detail.spanId,
      detail.type,
      detail.linkedCollectionTitle
    ].map(text).join("|")).join(";");
    const identity = [
      data.poleId,
      data.phase,
      data.section,
      data.code,
      data.title,
      data.expected,
      data.actual,
      details
    ].map(normalizedText).join("||");
    return `review-${stableReviewHash(identity)}`;
  }

  function normalizeSpanType(value) {
    const clean = normalizedText(value).replace(/[^a-z]/g, "");
    if (clean === "forespan") return "FORE";
    if (clean === "backspan") return "BACK";
    if (clean === "other" || clean === "otherspan") return "OTHER";
    return clean.toUpperCase();
  }

  function check(data) {
    const ignoreKey = checkIdentity(data);
    const ignored = Boolean(S().getState().excelReviewIgnoredChecks?.[ignoreKey]);
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
      details: Array.isArray(data.details) ? data.details : [],
      ignoreKey,
      ignored
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
    const accepted = ["lowestpowerdisplay", "lowpowerattachmentdisplay"];
    const entries = Object.entries(row || {});
    for (const normalizedName of accepted) {
      const match = entries.find(([header]) => I().normalizeHeaderName(header) === normalizedName);
      if (match && text(match[1])) return match[1];
    }
    return "";
  }

  function isMidAmProject() {
    const settings = S().getState().settings || {};
    return text(settings.projectProfile).toUpperCase() === "METRONET"
      && text(settings.proposedOwner || "MidAm").toUpperCase() === "MIDAM";
  }

  function collectionPoleId(row) {
    return text(pick(row, ["Id", "Pole ID", "PoleId", "PoleName", "Structure Number", "Pole"]));
  }

  function collectionSequence(row) {
    return text(pick(row, ["Sequence", "Seq"]));
  }

  // Excel can expose a formatted MidAm Sequence such as 058 as the number 58.
  // Normalize the numeric part back to three digits and retain one optional
  // route suffix letter before comparing it with the pole Id.
  function normalizeMidAmSequence(value) {
    const raw = text(value).toUpperCase().replace(/\s+/g, "");
    const match = raw.match(/^(\d{1,3})([A-Z])?$/);
    return match ? `${match[1].padStart(3, "0")}${match[2] || ""}` : "";
  }

  function midAmIdSequence(value) {
    return text(value).toUpperCase().split(/\s+/).filter(Boolean)[0] || "";
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
    const byCanonicalPole = new Map();
    entries.forEach(entry => {
      if (entry.poleId) {
        byPole.set(normalizedText(entry.poleId), entry.poleId);
        const canonical = S().canonicalPoleIdentity(entry.poleId);
        if (canonical && !byCanonicalPole.has(canonical)) byCanonicalPole.set(canonical, entry.poleId);
      }
      if (entry.collectionId && entry.poleId) byCollection.set(normalizedText(entry.collectionId), entry.poleId);
    });
    return { byPole, byCollection, byCanonicalPole };
  }

  function resolvePole(value, maps) {
    const candidate = normalizedText(value);
    if (!candidate) return "";
    if (maps.byPole.has(candidate)) return maps.byPole.get(candidate);
    if (maps.byCollection.has(candidate)) return maps.byCollection.get(candidate);
    const canonical = S().canonicalPoleIdentity(value);
    if (maps.byCanonicalPole.has(canonical)) return maps.byCanonicalPole.get(canonical);
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
    const applicable = checks.filter(item => item.phase === phase && item.applicable !== false && !item.ignored);
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

    const midAm = isMidAmProject();
    const normalizedSequence = midAm ? normalizeMidAmSequence(entry.sequence) : entry.sequence;
    const idSequence = midAm ? midAmIdSequence(entry.poleId) : "";
    if (!entry.sequence) {
      add(result, {
        phase: "HOA", section: "Collection", code: "MISSING_SEQUENCE", status: "ERROR",
        title: "Sequence", message: "Sequence is empty.", expected: "Sequence matching the start of Id", actual: "Empty"
      });
    } else if (midAm && !normalizedSequence) {
      add(result, {
        phase: "HOA", section: "Collection", code: "INVALID_MIDAM_SEQUENCE", status: "ERROR",
        title: "Sequence", message: `Sequence ${entry.sequence} must contain three digits and may end with one letter.`,
        expected: "000 or 000A", actual: entry.sequence
      });
    }

    if (midAm && entry.poleId && !/^\d{3}[A-Z]?$/.test(idSequence)) {
      add(result, {
        phase: "HOA", section: "Collection", code: "INVALID_MIDAM_ID_SEQUENCE", status: "ERROR",
        title: "Id / Sequence", message: `The first block of Id ${entry.poleId} must contain three digits and may end with one letter.`,
        expected: "000 or 000A", actual: idSequence || "Empty"
      });
    } else if (midAm && entry.poleId && normalizedSequence && idSequence !== normalizedSequence) {
      add(result, {
        phase: "HOA", section: "Collection", code: "SEQUENCE_ID_MISMATCH", status: "ERROR",
        title: "Sequence", message: `Sequence must equal ${idSequence}, derived from Id ${entry.poleId}.`,
        expected: idSequence, actual: normalizedSequence
      });
    } else if (!midAm && entry.poleId && !normalizedText(entry.poleId).startsWith(normalizedText(normalizedSequence))) {
      add(result, {
        phase: "HOA", section: "Collection", code: "SEQUENCE_ID_MISMATCH", status: "ERROR",
        title: "Sequence", message: `Sequence ${entry.sequence} does not match the start of Id ${entry.poleId}.`,
        expected: `${entry.sequence}...`, actual: entry.poleId
      });
    }

    if (midAm) {
      const owner = text(pick(entry.row, ["Owner"], { contains: false }));
      const normalizedOwner = normalizedText(owner).replace(/\s*>\s*/g, ">");
      if (!owner) {
        add(result, {
          phase: "HOA", section: "Collection", code: "MISSING_MIDAM_COLLECTION_OWNER", status: "ERROR",
          title: "Owner", message: "Collection Owner is empty. MidAm requires UTILITY > MidAm.",
          expected: "UTILITY > MidAm", actual: "Empty"
        });
      } else if (normalizedOwner !== "utility>midam") {
        add(result, {
          phase: "HOA", section: "Collection", code: "UNEXPECTED_MIDAM_COLLECTION_OWNER", status: "WARNING",
          title: "Owner", message: `Collection Owner ${owner} differs from the MidAm project owner.`,
          expected: "UTILITY > MidAm", actual: owner
        });
      }
    }

    const lowPower = exactDisplayLowPower(entry.row);
    if (!text(lowPower)) {
      add(result, {
        phase: "HOA", section: "Collection", code: "MISSING_LOW_POWER", status: "ERROR",
        title: "Low Power", message: "Low Power Attachment.display is empty.", expected: "Non-empty value", actual: "Empty"
      });
    }

    const year = pick(entry.row, ["Year Installed"], { contains: false });
    if (!isMidAmProject() && !text(year)) {
      add(result, {
        phase: "HOA", section: "Collection", code: "MISSING_YEAR_INSTALLED", status: "WARNING",
        title: "Year Installed", message: "Year Installed is missing. Review the required project loading manually.",
        expected: "Non-empty value", actual: "Empty"
      });
    }
  }

  function missingAnchorFields(row) {
    const firstValue = names => {
      for (const name of names) {
        const value = pick(row, [name], { contains: false });
        if (text(value)) return value;
      }
      return "";
    };
    const required = [
      ["Collection Id", pick(row, ["collectionId", "Collection ID"], { contains: false })],
      ["Id", pick(row, ["Id", "Pole ID"], { contains: false })],
      ["Anchor Index", pick(row, ["Anchor Index"], { contains: false })],
      ["Anchor Id", pick(row, ["Anchor Id", "Anchor ID"], { contains: false })],
      ["Type", pick(row, ["Type"], { contains: false })],
      ["Lead Length", firstValue(["Lead Length.display", "Lead Length"])],
      ["Lead Length provider", pick(row, ["Lead Length.provider", "Lead Length Provider"], { contains: false })],
      ["Lead Length bearing", firstValue(["Lead Length.bearing.display", "Lead Length.bearing"])],
      ["Lead Length pitch", firstValue(["Lead Length.pitch.display", "Lead Length.pitch"])],
      ["Owner", pick(row, ["Owner"], { contains: false })],
      ["Guys", pick(row, ["Guys"], { contains: false })]
    ];
    return required.filter(([, value]) => !text(value)).map(([label]) => label);
  }

  function addAnchorChecks(result, poleId) {
    const wanted = normalizedText(poleId);
    rows("anchors").forEach((row, index) => {
      const rowPoleId = text(pick(row, ["Id", "Pole ID"], { contains: false }));
      if (!rowPoleId || normalizedText(rowPoleId) !== wanted) return;
      const missing = missingAnchorFields(row);
      if (!missing.length) return;
      add(result, {
        phase: "HOA", section: "Anchor", code: "ANCHOR_MISSING_REQUIRED_DATA", status: "ERROR",
        title: "Required Fields", message: `Anchor row ${index + 2} has empty required fields: ${missing.join(", ")}.`,
        expected: "All required Anchor fields populated", actual: `Missing: ${missing.join(", ")}`
      });
    });
  }

  function addUnassignedAnchorChecks(globalChecks) {
    rows("anchors").forEach((row, index) => {
      const poleId = text(pick(row, ["Id", "Pole ID"], { contains: false }));
      if (poleId) return;
      const missing = missingAnchorFields(row);
      globalChecks.push(check({
        phase: "HOA", section: "Anchor", code: "ANCHOR_MISSING_REQUIRED_DATA", status: "ERROR",
        title: "Required Fields", message: `Anchor row ${index + 2} has no Id and cannot be assigned to a pole.`,
        expected: "All required Anchor fields populated", actual: `Missing: ${missing.join(", ") || "Id"}`
      }));
    });
  }

  function addSpanCountChecks(result, poleSpans) {
    const foreSpans = poleSpans.filter(span => span.type === "FORE");
    if (foreSpans.length !== 1) {
      const listed = foreSpans.map(span => span.spanIndex || span.spanId || `row ${span.sourceRow}`).join(", ");
      add(result, {
        phase: "HOA", section: "Span", code: foreSpans.length ? "MULTIPLE_FORE_SPANS" : "MISSING_FORE_SPAN",
        status: "ERROR", title: "Fore Span",
        message: `Expected exactly one Fore Span. Found ${foreSpans.length}${listed ? `: ${listed}` : ""}.`,
        expected: "1", actual: String(foreSpans.length), details: foreSpans.map(span => span.details)
      });
    }

    const backSpans = poleSpans.filter(span => span.type === "BACK");
    if (backSpans.length === 1) return;
    if (backSpans.length === 0) {
      add(result, {
        phase: "HOA", section: "Span", code: "MISSING_BACK_SPAN", status: "WARNING", title: "Back Span",
        message: "No Back Span was found. Zero is allowed, but the pole connection should be reviewed.",
        expected: "0 or 1; 1 preferred", actual: "0"
      });
      return;
    }
    const listed = backSpans.map(span => span.spanIndex || span.spanId || `row ${span.sourceRow}`).join(", ");
    add(result, {
      phase: "HOA", section: "Span", code: "MULTIPLE_BACK_SPANS", status: "ERROR", title: "Back Span",
      message: `Expected zero or one Back Span. Found ${backSpans.length}: ${listed}.`,
      expected: "0 or 1", actual: String(backSpans.length), details: backSpans.map(span => span.details)
    });
  }

  function addLinkedCollectionChecks(result, poleSpans, maps) {
    poleSpans.forEach(span => {
      if (!span.linkedTitle) {
        if (span.type === "OTHER") return;
        add(result, {
          phase: "HOA", section: "Span", code: "MISSING_LINKED_COLLECTION_TITLE", status: "WARNING", level: "low",
          title: "Linked Collection", message: "Linked Collection.Title is empty. Review the span connection.",
          expected: "Existing Collection Id", actual: "Empty", details: [span.details]
        });
        return;
      }
      // A populated linked title can legitimately point to a pole maintained
      // in another job. Only an empty Fore/Back link is reviewable here.
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

  function addMidAmChecks(result, poleId) {
    const settings = S().getState().settings || {};
    if (text(settings.projectProfile).toUpperCase() !== "METRONET"
      || text(settings.proposedOwner || "MidAm").toUpperCase() !== "MIDAM") return;

    rowsForPole("spanWires", poleId).forEach((row, index) => {
      const ownerRaw = text(pick(row, ["Owner", "owner"]));
      const owner = normalizedText(ownerRaw);
      const sizeRaw = text(pick(row, ["Size", "Size.display", "Wire Size"]));
      const size = normalizedText(sizeRaw);
      const insulatorRaw = text(pick(row, ["Insulator"]));
      const insulator = normalizedInsulator(insulatorRaw);
      const descriptor = `${ownerRaw || "No owner"} / ${sizeRaw || `row ${index + 2}`}`;
      const power = I().isPowerWire ? I().isPowerWire(row) : /^utility\s*>/i.test(ownerRaw);
      const communication = I().isCommunicationWire ? I().isCommunicationWire(row) : !power;

      if (communication && !MIDAM_COMM_INSULATORS.includes(insulator)) {
        add(result, {
          phase: "HOA", section: "Span.Wire", code: "INVALID_MIDAM_COMM_INSULATOR", status: "ERROR",
          title: "Communication Insulator", message: `Invalid MidAm communication insulator for ${descriptor}.`,
          expected: MIDAM_COMM_INSULATORS.join(", "), actual: insulatorRaw || "Empty"
        });
      }

      if (!power) return;
      const powerType = /primary/i.test(sizeRaw) ? "PRIMARY" : /secondary/i.test(sizeRaw) ? "SECONDARY" : /neutral/i.test(sizeRaw) ? "NEUTRAL" : "";
      if (!powerType) return;
      if (owner !== "utility > midam") {
        add(result, {
          phase: "HOA", section: "Span.Wire", code: "INVALID_MIDAM_POWER_OWNER", status: "ERROR",
          title: `${powerType} Owner`, message: `${powerType} owner must be UTILITY > MidAm.`,
          expected: "UTILITY > MidAm", actual: ownerRaw || "Empty"
        });
      }
      if (size !== MIDAM_POWER_SIZES[powerType]) {
        add(result, {
          phase: "HOA", section: "Span.Wire", code: `INVALID_MIDAM_${powerType}_SIZE`, status: "ERROR",
          title: `${powerType} Size`, message: `Invalid MidAm ${powerType.toLowerCase()} size.`,
          expected: MIDAM_POWER_SIZES[powerType], actual: sizeRaw || "Empty"
        });
      }
      if (!MIDAM_POWER_INSULATORS.includes(insulator)) {
        add(result, {
          phase: "HOA", section: "Span.Wire", code: `INVALID_MIDAM_${powerType}_INSULATOR`, status: "ERROR",
          title: `${powerType} Insulator`, message: `Invalid MidAm ${powerType.toLowerCase()} insulator for ${descriptor}.`,
          expected: MIDAM_POWER_INSULATORS.join(", "), actual: insulatorRaw || "Empty"
        });
      }
    });

    rowsForPole("anchorGuys", poleId).forEach(row => {
      const ownerRaw = text(pick(row, ["Owner", "owner"]));
      const owner = normalizedText(ownerRaw);
      const size = text(pick(row, ["Size"]));
      const utility = owner === "utility > midam";
      const communication = owner.startsWith("communication >");
      const valid = utility
        ? /(?:ehs\s*)?1\s*\/\s*2|0\.500/i.test(size)
        : communication ? /(?:ehs\s*)?3\s*\/\s*8|0\.375/i.test(size) : true;
      if (valid) return;
      add(result, {
        phase: "HOA", section: "Anchor.Guys", code: utility ? "INVALID_MIDAM_POWER_GUY_SIZE" : "INVALID_MIDAM_COMM_GUY_SIZE", status: "ERROR",
        title: utility ? "Power Guy Size" : "Communication Guy Size",
        message: `${utility ? "Power" : "Communication"} ANC/guy has an invalid MidAm size.`,
        expected: utility ? "1/2\" (0.500\")" : "3/8\" (0.375\")", actual: size || "Empty"
      });
    });

    rowsForPole("equipment", poleId)
      .filter(row => /street\s*light|streetlight/i.test(text(pick(row, ["Type"]))))
      .forEach(row => {
        const owner = text(pick(row, ["Owner"]));
        if (normalizedText(owner) !== "utility > midam") {
          add(result, {
            phase: "HOA", section: "Equipment", code: "INVALID_MIDAM_STREETLIGHT_OWNER", status: "ERROR",
            title: "Streetlight Owner", message: "MidAm streetlights must use UTILITY > MidAm.",
            expected: "UTILITY > MidAm", actual: owner || "Empty"
          });
        }
        const bottom = text(pick(row, ["Bottom Height.display", "Bottom Height Display"]));
        const dripLoop = text(pick(row, ["Drip Loop Height.display", "Drip Loop Height Display"]));
        if (!bottom || !dripLoop) {
          add(result, {
            phase: "HOA", section: "Equipment", code: "MISSING_MIDAM_STREETLIGHT_HEIGHT", status: "ERROR",
            title: "Streetlight Clearance Data",
            message: "Streetlight bottom and uncovered drip-loop heights are required for MidAm pole clearance.",
            expected: "Bottom Height.display and Drip Loop Height.display", actual: `Bottom ${bottom || "Empty"}; Drip loop ${dripLoop || "Empty"}`
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

  function makeReadyRowHasFinalData(row) {
    return [
      pick(row, ["Attachment Size"], { contains: true }),
      pick(row, ["Attachment Type"], { contains: true }),
      pick(row, ["Attachment Height.display", "Attachment Height Display"], { contains: true }),
      pick(row, ["Proposed Mid Span.display", "Proposed Midspan.display", "Proposed Mid Span Display"], { contains: true }),
      pick(row, ["Make Ready Notes", "MR Notes", "Notes"], { contains: true }),
      pick(row, ["Comm Transfers"], { contains: true })
    ].some(value => text(value));
  }

  function commTransferRowHasFinalData(row) {
    return [
      pick(row, ["Owner"]),
      pick(row, ["Height.display", "Height Display", "Height"], { contains: true })
    ].some(value => text(value));
  }

  // Final review is a workbook stage, not a per-pole inference. A raw HOA
  // workbook can produce derived Calculator MR, but that must not make an
  // absent Make Ready deliverable look like a final-review error.
  function workbookHasFinalReviewData() {
    return rows("makeReady").some(makeReadyRowHasFinalData)
      || rows("commTransfers").some(commTransferRowHasFinalData);
  }

  function excelFinalRowsForPole(poleId) {
    return rowsForPole("makeReady", poleId).filter(makeReadyRowHasFinalData);
  }

  function excelTransferRowsForPole(poleId) {
    return rowsForPole("commTransfers", poleId).filter(commTransferRowHasFinalData);
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

  function instructionTextKey(value) {
    return normalizedText(text(value).replace(/[’‘]/g, "'").replace(/[“”]/g, "\""))
      .replace(/\s+/g, " ")
      .replace(/[.!]+$/, "")
      .trim();
  }

  function uniqueInstructions(lines) {
    const seen = new Set();
    return lines.filter(line => {
      const key = instructionTextKey(line);
      if (!key || seen.has(key)) return false;
      seen.add(key);
      return true;
    });
  }

  function isModelOnlyInstruction(value) {
    const clean = normalizedText(value);
    return /\b(?:pl|pla)\s+new\s+anc\b/.test(clean)
      || /\bnew\s+comm\s+anc\b/.test(clean)
      || /\bsplit\b.*\bpwr\s+anc\b/.test(clean)
      || /\bproposed\s+slack\s+span\b/.test(clean);
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
      if (owner === "ctl") return /century\s*link|centurylink|\bctl\b|\btelco\b/.test(clean);
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

  function instructionMismatchScore(expected, actual) {
    if (expected.action !== actual.action) return Infinity;
    // "other" has no reliable action identity and could pair unrelated prose.
    if (expected.action === "other") return Infinity;
    if (expected.owners.length && actual.owners.length
      && !expected.owners.some(owner => actual.owners.includes(owner))) return Infinity;
    const expectedHeight = expected.heights[0];
    const actualHeight = actual.heights[0];
    const heightDistance = expectedHeight === undefined || actualHeight === undefined
      ? 0
      : Math.abs(expectedHeight - actualHeight);
    const directionPenalty = expected.direction && actual.direction && expected.direction !== actual.direction ? 120 : 0;
    return heightDistance + directionPenalty;
  }

  function addMakeReadyNotesComparison(result, poleId, makeReadyRows) {
    const generated = S().getState().mr.find(item => normalizedText(item.poleId) === normalizedText(poleId));
    const expectedLines = uniqueInstructions(splitInstructions(generated?.text || ""));
    const actualLines = uniqueInstructions(makeReadyRows.flatMap(row => splitInstructions(pick(row, ["Make Ready Notes", "MR Notes", "Notes"], { contains: true }))));
    const used = new Set();

    expectedLines.forEach(line => {
      const expected = instructionSignature(line, poleId);
      const exactKey = instructionTextKey(line);
      let matchIndex = actualLines.findIndex((actualLine, index) => !used.has(index) && instructionTextKey(actualLine) === exactKey);
      if (matchIndex < 0) {
        matchIndex = actualLines.findIndex((actualLine, index) => !used.has(index) && instructionMatches(expected, instructionSignature(actualLine, poleId)));
      }
      if (matchIndex >= 0) {
        used.add(matchIndex);
        return;
      }
      // Pair a near instruction with the same action/owner before reporting a
      // mismatch. This avoids showing one missing error plus one additional
      // warning for the same UG, riser, movement, or transfer instruction.
      const mismatchCandidates = actualLines
        .map((actualLine, index) => ({
          actualLine,
          index,
          score: used.has(index) ? Infinity : instructionMismatchScore(expected, instructionSignature(actualLine, poleId))
        }))
        .filter(candidate => Number.isFinite(candidate.score))
        .sort((a, b) => a.score - b.score);
      if (mismatchCandidates.length) {
        const mismatch = mismatchCandidates[0];
        used.add(mismatch.index);
        add(result, {
          phase: "FINAL", section: "Make Ready", code: "MR_INSTRUCTION_MISMATCH", status: "ERROR",
          title: "Make Ready Notes", message: "Calculator and Excel instructions differ.",
          expected: line, actual: mismatch.actualLine
        });
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
      // Anchor placement and split-pole directives are proposed by the model,
      // outside the calculator's current deterministic inputs. They are valid
      // supplemental MR and are not reported as unmatched instructions.
      if (isModelOnlyInstruction(line)) return;
      add(result, {
        phase: "FINAL", section: "Make Ready", code: "ADDITIONAL_MR_INSTRUCTION", status: "WARNING",
        title: "Make Ready Notes", message: `Excel instruction has no Calculator equivalent: ${line}`,
        expected: "Calculator instruction", actual: line
      });
    });
  }

  function validUndergroundMakeReady(value) {
    const notes = text(value);
    if (/underground|\bgoing\s+ug\b|\bgo\s+ug\b|\bug\s+transfer\b/i.test(notes)) return true;
    const unableToAttach = notes.match(/\bunable\s+to\s+attach\s+due\s+to\s+([^\r\n.]+)/i);
    if (!unableToAttach) return false;
    const reason = normalizedText(unableToAttach[1]).replace(/[()[\]]/g, "").trim();
    return Boolean(reason && !/^(reason|reasoning|specify reason|insert reason)$/.test(reason));
  }

  function expectedTransfers(poleId) {
    const seen = new Set();
    return S().getSpanCommsForPole(poleId).filter(row => row.transferToNewPole && text(row.existingHOAChange || row.existingHOA)).reduce((items, row) => {
      const owner = normalizedOwner(row.rawOwner || row.ownerBase || row.owner);
      const heightDisplay = row.existingHOAChange || row.existingHOA;
      const height = H().parseHeight(heightDisplay);
      const key = `${owner}|${height}`;
      if (!owner || height === null || seen.has(key)) return items;
      seen.add(key);
      items.push({ owner, ownerDisplay: text(row.rawOwner || row.ownerBase || row.owner), height, heightDisplay });
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
      if (!validUndergroundMakeReady(notes)) {
        add(result, {
          phase: "FINAL", section: "Make Ready", code: "MISSING_UG_INSTRUCTION", status: "ERROR",
          title: "Underground Make Ready", message: "Calculator is UG, but Excel does not contain one valid underground resolution instruction.",
          expected: "UG instruction or Unable to attach due to a specific reason", actual: notes || "Empty"
        });
      } else {
        add(result, {
          phase: "FINAL", section: "Make Ready", code: "VALID_UG_INSTRUCTION", status: "PASS",
          title: "Underground Make Ready", message: "Excel contains a valid underground resolution instruction.",
          expected: "One valid UG resolution", actual: notes
        });
      }
      // UG replacement text describes alternative reasons, not six mandatory
      // instructions. The semantic check above therefore produces at most one
      // problem instead of comparing every generated template line literally.
      addTransferComparisons(result, entry.poleId);
      return;
    }

    addProposedComparisons(result, entry.poleId, makeReadyRows);
    addMakeReadyNotesComparison(result, entry.poleId, makeReadyRows);
    addTransferComparisons(result, entry.poleId);
  }

  function finalizeResult(result) {
    result.hoaStatus = statusFromChecks(result.checks, "HOA");
    result.finalStatus = result.finalApplicable === false
      ? "NOT_APPLICABLE"
      : statusFromChecks(result.checks, "FINAL");
    const statuses = [result.hoaStatus, result.finalStatus]
      .filter(status => !["NOT_READY", "NOT_APPLICABLE"].includes(status));
    result.overallStatus = statuses.includes("ERROR") ? "ERROR" : statuses.includes("WARNING") ? "WARNING" : "PASS";
    const applicableProblems = result.checks.filter(item => item.applicable !== false && !item.ignored && ["ERROR", "WARNING"].includes(item.status));
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
    // Review order follows the job sequence at all times. Severity remains
    // visible in badges and borders but never rearranges the physical route.
    return results.sort((a, b) => naturalCompare(a.sequence || a.poleId, b.sequence || b.poleId));
  }

  function buildSummary(results, globalChecks) {
    const activeGlobalChecks = globalChecks.filter(item => !item.ignored);
    return {
      errors: results.filter(result => result.overallStatus === "ERROR").length + activeGlobalChecks.filter(item => item.status === "ERROR").length,
      warnings: results.filter(result => result.overallStatus === "WARNING").length + activeGlobalChecks.filter(item => item.status === "WARNING").length,
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
    const finalReviewApplicable = workbookHasFinalReviewData();

    if (!hasHeader("collection", ["Id"])) {
      current.globalChecks.push(check({
        phase: "HOA", section: "Collection", code: "MISSING_ID_COLUMN", status: "ERROR",
        title: "Collection Id Column", message: "Collection is missing the Id column.", expected: "Id column", actual: "Missing"
      }));
    }
    addUnassignedAnchorChecks(current.globalChecks);

    current.results = entries.map(entry => {
      const result = {
        poleId: entry.displayPoleId,
        sourcePoleId: entry.poleId,
        sequence: entry.sequence,
        sourceRow: entry.sourceRow,
        hoaStatus: "PASS",
        finalStatus: finalReviewApplicable ? "NOT_READY" : "NOT_APPLICABLE",
        finalApplicable: finalReviewApplicable,
        overallStatus: "PASS",
        checks: []
      };
      const poleSpans = spans.filter(span => normalizedText(span.fromPole) === normalizedText(entry.poleId));
      addCollectionChecks(result, entry);
      addAnchorChecks(result, entry.poleId);
      addSpanCountChecks(result, poleSpans);
      addLinkedCollectionChecks(result, poleSpans, maps);
      addReciprocalChecks(result, poleSpans, spans);
      addEnvironmentChecks(result, poleSpans, spans, environmentPairsSeen);
      if (entry.poleId) addIntecWireChecks(result, entry.poleId);
      if (entry.poleId) addMidAmChecks(result, entry.poleId);
      if (finalReviewApplicable) addFinalChecks(result, entry);
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

  function setCheckIgnored(ignoreKey, ignored) {
    const key = text(ignoreKey);
    if (!key) return current;
    const state = S().getState();
    state.excelReviewIgnoredChecks = state.excelReviewIgnoredChecks || {};
    if (ignored) state.excelReviewIgnoredChecks[key] = true;
    else delete state.excelReviewIgnoredChecks[key];
    return runReview();
  }

  /** @namespace ExcelReview */
  global.ExcelReview = {
    runReview,
    reviewPole,
    getResults,
    getSummary,
    getReviewState,
    setCheckIgnored,
    clearResults
  };
})(window);
