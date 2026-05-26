(function (global) {
  "use strict";

  // ExcelImport translates workbook sheets into the normalized graph stored in
  // AppStore. Flexible header matching keeps imports resilient to export drift.
  const S = () => global.AppStore;
  const H = () => global.HeightUtils;

  function rowsToObjects(rows) {
    if (!Array.isArray(rows) || rows.length === 0) return [];
    const headerRow = rows.find(row => Array.isArray(row) && row.some(v => v !== undefined && v !== null && String(v).trim() !== ""));
    if (!headerRow) return [];
    const headerIndex = rows.indexOf(headerRow);
    const headers = headerRow.map(h => String(h ?? "").trim());

    return rows.slice(headerIndex + 1)
      .filter(row => Array.isArray(row) && row.some(v => v !== undefined && v !== null && String(v).trim() !== ""))
      .map(row => {
        const obj = {};
        headers.forEach((header, index) => {
          if (header) obj[header] = row[index] ?? "";
        });
        return obj;
      });
  }

  function normalizeHeaderName(name) {
    return String(name || "")
      .toLowerCase()
      .normalize("NFD")
      .replace(/[\u0300-\u036f]/g, "")
      .replace(/[._\-\s()\[\]/]/g, "")
      .replace(/[^a-z0-9]/g, "");
  }

  function isBlank(value) {
    return value === undefined || value === null || String(value).trim() === "";
  }

  function findKey(row, names, options = {}) {
    if (!row) return "";
    const keys = Object.keys(row);
    const wanted = names.map(name => ({ raw: name, normalized: normalizeHeaderName(name) })).filter(item => item.normalized);

    for (const item of wanted) {
      const exact = keys.find(k => normalizeHeaderName(k) === item.normalized);
      if (exact) return exact;
    }

    if (options.contains !== false) {
      for (const item of wanted) {
        const partial = keys.find(k => {
          const normalizedKey = normalizeHeaderName(k);
          return normalizedKey.includes(item.normalized) || item.normalized.includes(normalizedKey);
        });
        if (partial) return partial;
      }
    }

    return "";
  }

  function pick(row, names, options = {}) {
    const key = findKey(row, names, options);
    return key ? row[key] : "";
  }

  function findSheet(workbook, preferredNames) {
    const normalized = preferredNames.map(normalizeHeaderName);
    const name = workbook.SheetNames.find(sheetName => normalized.includes(normalizeHeaderName(sheetName)));
    return name ? workbook.Sheets[name] : null;
  }

  function cleanHeight(value) {
    if (isBlank(value)) return "";
    const parsed = H().parseHeight(value);
    return parsed === null ? String(value).trim() : H().formatHeight(parsed);
  }

  function heightFromMeters(value) {
    if (isBlank(value)) return "";
    const n = Number(value);
    if (!Number.isFinite(n)) return cleanHeight(value);
    return H().formatHeight(Math.round(n * 3.280839895 * 12));
  }

  function pickDisplayValue(row, displayNames) {
    if (!row) return "";
    const keys = Object.keys(row);
    const wanted = displayNames.map(normalizeHeaderName).filter(Boolean);

    const exactDisplay = keys.find(k => wanted.includes(normalizeHeaderName(k)) && normalizeHeaderName(k).includes("display"));
    if (exactDisplay && !isBlank(row[exactDisplay])) return row[exactDisplay];

    const partialDisplay = keys.find(k => {
      const normalizedKey = normalizeHeaderName(k);
      if (!normalizedKey.includes("display")) return false;
      return wanted.some(w => {
        const base = w.replace(/display/g, "");
        return normalizedKey.includes(base) || base.includes(normalizedKey.replace(/display/g, ""));
      });
    });
    return partialDisplay && !isBlank(row[partialDisplay]) ? row[partialDisplay] : "";
  }

  function heightFromRow(row, displayNames, decimalNames) {
    const display = pickDisplayValue(row, displayNames);
    if (!isBlank(display)) return cleanHeight(display);
    const decimal = pick(row, decimalNames, { contains: true });
    return heightFromMeters(decimal);
  }

  function parsePoleHeightFromType(typeValue) {
    const text = String(typeValue || "");
    const parts = text.split(">").map(p => p.trim()).filter(Boolean);
    const last = parts.length ? parts[parts.length - 1] : text;
    const match = last.match(/\d+(?:\.\d+)?/);
    if (!match) return "";
    const feet = Number(match[0]);
    if (!Number.isFinite(feet)) return "";
    return H().formatHeight(Math.round(feet * 12));
  }

  const ANSI_POLE_CLASSES = ["H6", "H5", "H4", "H3", "H2", "H1", "1", "2", "3", "4", "5", "6", "7", "9", "10"];
  const ANSI_CLASS_TABLE = {
    20: [null, null, null, null, null, null, 31, 29, 27, 25, 23, 21, 19.6, 17.6, 14],
    25: [null, null, null, null, null, null, 33.5, 31.5, 29.5, 27.5, 25.5, 23, 21.5, 19.5, 15],
    30: [null, null, null, null, null, null, 36.5, 34, 32, 29.5, 27.5, 25, 23.5, 20.5, null],
    35: [null, null, null, null, null, null, 39, 36.5, 34, 31.5, 29, 27, 25, null, null],
    40: [null, null, 51, 48.5, 46, 43.5, 41, 38.5, 36, 33.5, 31, 28.5, null, null, null],
    45: [58.5, 56, 53.5, 51, 48.5, 45.5, 43, 40.5, 37.5, 35, 32.5, 30, null, null, null],
    50: [61, 58.5, 55.5, 53, 50.5, 47.5, 45, 42, 39, 36.5, 34, null, null, null, null],
    55: [63.5, 60.5, 58, 55, 52, 49.5, 46.5, 43.5, 40.5, 38, null, null, null, null, null],
    60: [65.5, 62.5, 59.5, 57, 54, 51, 48, 45, 42, 39, null, null, null, null, null],
    65: [67.5, 64.5, 61.5, 58.5, 55.5, 52.5, 49.5, 46.5, 43.5, 40.5, null, null, null, null, null],
    70: [69, 66.5, 63.5, 60.5, 57, 54, 51, 48, 45, 41.5, null, null, null, null, null],
    75: [71, 68, 65, 62, 59, 55.5, 52.5, 49, 46, null, null, null, null, null, null],
    80: [72.5, 69.5, 66.5, 63.5, 60, 57, 54, 50.5, 47, null, null, null, null, null, null],
    85: [74.5, 71.5, 68, 65, 61.5, 58.5, 55, 51.5, 48, null, null, null, null, null, null],
    90: [76, 73, 69.5, 66.5, 63, 59.5, 56, 53, 49, null, null, null, null, null, null],
    95: [77.5, 74.5, 71, 67.5, 64.5, 61, 57, 54, null, null, null, null, null, null, null],
    100: [79, 76, 72.5, 69, 65.5, 62, 58.5, 55, null, null, null, null, null, null, null],
    105: [80.5, 77, 74, 70.5, 67, 63, 59.5, 56, null, null, null, null, null, null, null],
    110: [82, 78.5, 75, 71.5, 68, 64.5, 60.5, 57, null, null, null, null, null, null, null],
    115: [83.5, 80, 76.5, 72.5, 69, 65.5, 61.5, 58, null, null, null, null, null, null, null],
    120: [85, 81, 77.5, 74, 70, 66.5, 62.5, 59, null, null, null, null, null, null, null],
    125: [86, 82.5, 78.5, 75, 71, 67.5, 63.5, 59.5, null, null, null, null, null, null, null]
  };
  const ANSI_APPROX_GROUNDLINE_DISTANCE = {
    20: 4,
    25: 5,
    30: 5.5,
    35: 6,
    40: 6,
    45: 6.5,
    50: 7,
    55: 7.5,
    60: 8,
    65: 8.5,
    70: 9,
    75: 9.5,
    80: 10,
    85: 10.5,
    90: 11,
    95: 11,
    100: 11,
    105: 12,
    110: 12,
    115: 12,
    120: 12,
    125: 12
  };

  function parseNumber(value) {
    if (isBlank(value)) return null;
    const match = String(value).replace(/,/g, "").match(/-?\d+(?:\.\d+)?/);
    if (!match) return null;
    const n = Number(match[0]);
    return Number.isFinite(n) ? n : null;
  }

  function parseImportedType(typeValue) {
    const parts = String(typeValue || "").split(">").map(part => part.trim()).filter(Boolean);
    const height = parts.length >= 3 ? parseNumber(parts[2]) : parseNumber(parts[parts.length - 1]);
    return {
      species: parts[0] || "",
      classValue: parts.length >= 2 ? parts[1] : "",
      height: height === null ? "" : height
    };
  }

  function roundPoleLengthFromTip(tipHeight) {
    const tipInches = H().parseHeight(tipHeight);
    if (tipInches === null) return "";
    const totalFeet = (tipInches / 12) * 1.10 + 2;
    // IKE estimates can land just above a standard length because embedment is
    // inferred. Keep tiny overages on the lower length, but move up once the
    // estimate is meaningfully past that boundary.
    const lower = Math.floor(totalFeet / 5) * 5;
    const overLower = totalFeet - lower;
    return overLower <= 1 ? lower : lower + 5;
  }

  function classFromAnsiTable(height, circumference) {
    const row = ANSI_CLASS_TABLE[height];
    if (!row || circumference === null) return "";
    let match = null;
    row.forEach((minimum, index) => {
      if (minimum === null || minimum === undefined) return;
      if (minimum <= circumference && (!match || minimum > match.minimum)) {
        match = { classValue: ANSI_POLE_CLASSES[index], minimum };
      }
    });
    return match || null;
  }

  function recalculatePoleClassCheck(data = {}) {
    const importedType = String(data.importedType || "").trim();
    const parsedType = parseImportedType(importedType);
    const manualDiameter = parseNumber(data.manualDiameter);
    const importedDiameter = parseNumber(data.importedDiameter);
    const importedCircumference = parseNumber(data.importedCircumference ?? data.circumference);
    const diameterForCalc = manualDiameter;
    const circumference = diameterForCalc === null
      ? (importedCircumference ?? (importedDiameter === null ? null : Number((importedDiameter * 3.14).toFixed(2))))
      : Number((diameterForCalc * 3.14).toFixed(2));
    const calculatedHeight = roundPoleLengthFromTip(data.tip);
    const tableMatch = classFromAnsiTable(calculatedHeight, circumference);
    const calculatedClass = tableMatch?.classValue || "";
    const species = parsedType.species || "Pole";
    const expectedType = calculatedClass && calculatedHeight ? `${species} > ${calculatedClass} > ${calculatedHeight}` : "";
    const importedClass = String(parsedType.classValue || "").toUpperCase();
    const expectedClass = String(calculatedClass || "").toUpperCase();
    const issues = [];

    if (!data.tip) issues.push("Missing Tip");
    if (circumference === null) issues.push("Missing Circumference");
    if (!importedType) issues.push("Missing Type");
    if (calculatedHeight && !ANSI_CLASS_TABLE[calculatedHeight]) issues.push("No ANSI height row");
    if (data.tip && circumference !== null && !tableMatch) issues.push("No table match");
    if (tableMatch && importedClass && importedClass !== expectedClass) issues.push("Class mismatch");
    if (calculatedHeight && parsedType.height && Number(parsedType.height) !== Number(calculatedHeight)) issues.push("Height mismatch");

    return {
      ...data,
      manualDiameter: String(data.manualDiameter || ""),
      circumference: circumference === null ? "" : String(circumference),
      calculatedHeight: calculatedHeight || "",
      calculatedClass,
      expectedType,
      status: issues.length ? issues.join(", ") : "OK",
      tableMinimumCircumference: tableMatch?.minimum || "",
      circumferenceSource: diameterForCalc === null ? (importedCircumference === null ? "Imported diameter" : "Collection") : "Manual diameter"
    };
  }

  function buildPoleClassChecks(collectionRows) {
    return collectionRows.map(row => {
      const poleId = String(pick(row, ["Id", "Pole ID", "PoleId", "PoleName", "Structure Number", "Pole"])).trim();
      if (!poleId) return null;
      const importedType = String(pick(row, ["Type", "Pole Type"])).trim();
      const tipHeight = heightFromRow(row, ["Tip.display", "Tip Display"], ["Tip"]);
      const circumferenceRaw = pick(row, ["Circumference", "Circumfer", "Circumference (\")", "Ground Circumference"], { contains: true });
      const diameterRaw = pick(row, ["Diameter", "Groundline Diameter"], { contains: true });

      return recalculatePoleClassCheck({
        poleId,
        tip: tipHeight,
        importedCircumference: parseNumber(circumferenceRaw) === null ? "" : String(parseNumber(circumferenceRaw)),
        importedDiameter: parseNumber(diameterRaw) === null ? "" : String(parseNumber(diameterRaw)),
        manualDiameter: "",
        importedType,
        source: "Collection"
      });
    }).filter(Boolean);
  }

  function normalizeOwner(rawOwner, size) {
    const text = `${rawOwner || ""} ${size || ""}`.toLowerCase();
    if (/century\s*link|centurylink|\bctl\b|telco/.test(text)) return "CTL";
    if (/cable\s*one|cox|catv/.test(text)) return "CATV";
    if (/vexus/.test(text)) return "Vexus";
    if (/wecom/.test(text)) return "Wecom";
    if (/fiber/.test(text)) return "FIBER";
    if (/communication/.test(text)) {
      const owner = String(rawOwner || "").replace(/^COMMUNICATION\s*>\s*/i, "").trim();
      return owner || "COMM";
    }
    return "";
  }

  function parseAttachmentSize(value) {
    const raw = String(value || "").trim();
    if (!raw) {
      return {
        attachmentSizeRaw: "",
        attachmentMessenger: "",
        attachmentFiber: "",
        attachmentDirection: "",
        attachmentDirectionTokens: []
      };
    }

    const directionMatch = raw.match(/\(([^)]+)\)\s*$/);
    const attachmentDirection = directionMatch ? directionMatch[1].trim() : "";
    const withoutDirection = raw.replace(/\s*\([^)]+\)\s*$/, "").trim();
    const parts = withoutDirection.split(/\s+/).filter(Boolean);
    const attachmentMessenger = parts.shift() || "";
    const attachmentFiber = parts.join(" ");
    const attachmentDirectionTokens = attachmentDirection
      .split(/[\/,;\s]+/)
      .map(token => token.trim().toUpperCase())
      .filter(Boolean);

    return {
      attachmentSizeRaw: raw,
      attachmentMessenger,
      attachmentFiber,
      attachmentDirection,
      attachmentDirectionTokens
    };
  }

  function importMakeReadyReferences(makeReadyRows) {
    return makeReadyRows.map(row => {
      const poleId = String(pick(row, ["Id", "Pole ID", "PoleId", "Pole", "Structure Number"])).trim();
      const attachmentSize = parseAttachmentSize(pick(row, ["Attachment Size"], { contains: true }));
      if (!poleId && !attachmentSize.attachmentSizeRaw) return null;
      return S().createMakeReadyReference({
        poleId,
        collectionId: String(pick(row, ["collectionId", "Collection ID"])).trim(),
        makeReadyIndex: pick(row, ["Make Ready Index", "MR Index"], { contains: true }),
        makeReadyId: pick(row, ["Make Ready Id", "Make Ready ID", "MR ID"], { contains: true }),
        ...attachmentSize,
        attachmentType: pick(row, ["Attachment Type"], { contains: true }),
        attachmentHeight: heightFromRow(row, ["Attachment Height.display", "Attachment Height Display"], ["Attachment Height"]),
        proposedMidspan: heightFromRow(row, ["Proposed Mid Span.display", "Proposed Midspan.display", "Proposed Mid Span Display"], ["Proposed Mid Span", "Proposed Midspan"]),
        makeReadyNotes: pick(row, ["Make Ready Notes", "MR Notes", "Notes"], { contains: true }),
        raw: row
      });
    }).filter(Boolean);
  }

  function fallbackOwner(rawOwner, size, wireId) {
    const raw = String(rawOwner || "").replace(/^COMMUNICATION\s*>\s*/i, "").trim();
    if (raw) return raw;
    const text = String(size || "").trim();
    if (text) return text.split(">")[0].trim() || text;
    return wireId ? `UNKNOWN-${wireId}` : "UNKNOWN_COMM";
  }

  function isCommunicationWire(row) {
    const rawOwner = String(pick(row, ["Owner", "owner"])).trim();
    const size = String(pick(row, ["Size", "Size.display", "Wire Size"])).trim();
    const text = `${rawOwner} ${size}`.toLowerCase();
    if (/^utility\s*>/i.test(rawOwner)) return false;
    if (/primary|neutral|secondary|transformer|service\s+drop|street\s*light/i.test(size)) return false;
    return /communication|catv|telco|fiber|cable\s*one|century\s*link|centurylink|cox|vexus|wecom/i.test(text);
  }

  function isPowerWire(row) {
    const rawOwner = String(pick(row, ["Owner", "owner"])).trim();
    const size = String(pick(row, ["Size", "Size.display", "Wire Size"])).trim();
    const text = `${rawOwner} ${size}`.toLowerCase();
    return /^utility\s*>/i.test(rawOwner) || /primary|neutral|secondary|transformer/i.test(text);
  }

  function directionFromBearingDisplay(value) {
    const text = String(value || "").trim();
    if (!text) return { direction: "", degrees: "" };
    const n = Number(text.replace(/°/g, "").trim());
    if (!Number.isFinite(n)) return { direction: text, degrees: "" };
    const normalized = (((n % 360) + 360) % 360);
    const dirs8 = ["N", "NE", "E", "SE", "S", "SW", "W", "NW"];
    const index = Math.round(normalized / 45) % 8;
    return { direction: dirs8[index], degrees: Number(normalized.toFixed(2)) };
  }

  function normalizeEnvironment(value) {
    const raw = String(value || "").trim();
    if (!raw) return "NONE";
    const normalized = raw
      .toUpperCase()
      .normalize("NFD")
      .replace(/[\u0300-\u036f]/g, "")
      .replace(/[^A-Z0-9]+/g, "_")
      .replace(/^_+|_+$/g, "");
    const aliases = {
      NONE: "NONE",
      NO_DATA: "NONE",
      STREET: "STREET",
      VEHICULAR_TRAFFIC_STREET_ALONG_STREET_DRIVEWAY_ALLEY: "STREET",
      ALONG_STREET: "STREET",
      DRIVEWAY: "RESIDENTIAL_DRIVEWAY",
      HIGHWAY: "HIGHWAY",
      STATE_HIGHWAYS_AND_INTERSTATES: "HIGHWAY",
      PEDESTRIAN: "PEDESTRIAN",
      NO_VEHICULAR_TRAFFIC_PEDESTRIAN_ONLY_BACKYARD: "PEDESTRIAN",
      BACKYARD: "PEDESTRIAN",
      PARALLEL_TO_STREET: "PARALLEL_TO_STREET",
      OBSTRUCTED_PARALLEL_TO_STREET: "OBSTRUCTED_PARALLEL_TO_STREET",
      UNLIKELY_PARALLEL_TO_STREET: "UNLIKELY_PARALLEL_TO_STREET",
      RESIDENTIAL_DRIVEWAY: "RESIDENTIAL_DRIVEWAY",
      COMMERCIAL_DRIVEWAY: "COMMERCIAL_DRIVEWAY",
      PARKING_LOT: "PARKING_LOT",
      ALLEY: "ALLEY",
      RAILROAD: "RAILROAD",
      RURAL: "RURAL",
      FARM: "FARM",
      CULTIVATED_FARM_FIELDS_FOREST_INDUSTRIAL_SITES: "FARM",
      WATER_WITH_SAILBOATS: "WATER_WITH_SAILBOATS",
      WATER_SAILBOATS: "WATER_WITH_SAILBOATS",
      WATER_WITHOUT_SAILBOATS: "WATER_WITHOUT_SAILBOATS",
      WATER_NO_SAILBOATS: "WATER_WITHOUT_SAILBOATS",
      TROLLEY: "TROLLEY"
    };
    return aliases[normalized] || normalized;
  }

  function clearanceForEnvironment(environment) {
    const option = (S().ENVIRONMENT_OPTIONS || []).find(item => item.value === environment);
    return option ? option.clearance : "";
  }

  function importAppStateSheet(workbook) {
    const sheet = findSheet(workbook, ["AppState"]);
    if (!sheet) return null;
    const cell = sheet.flat().find(value => typeof value === "string" && value.trim().startsWith("{"));
    if (!cell) return null;
    return S().setState(JSON.parse(cell));
  }

  function importExportedTables(workbook) {
    const polesSheet = findSheet(workbook, ["Poles"]);
    const spansSheet = findSheet(workbook, ["Spans"]);
    const spanCommsSheet = findSheet(workbook, ["SpanComms"]);
    const spanSidesSheet = findSheet(workbook, ["SpanSides"]);
    const spanPowerSheet = findSheet(workbook, ["SpanPower"]);
    const makeReadyRefsSheet = findSheet(workbook, ["MakeReadyRefs", "Make Ready References"]);
    if (!polesSheet && !spansSheet && !spanCommsSheet && !spanSidesSheet && !spanPowerSheet && !makeReadyRefsSheet) return null;

    const state = S().resetState();
    state.importedFileName = "Reimported exported Excel";
    state.autoCreateSpanComms = false;

    rowsToObjects(polesSheet || []).forEach(row => {
      const poleId = String(pick(row, ["poleId", "Pole ID", "Id"])).trim();
      if (!poleId) return;
      S().upsertPole(S().createPole({
        poleId,
        collectionId: pick(row, ["collectionId", "Collection ID"]),
        sequence: pick(row, ["sequence", "Sequence"]),
        poleHeight: pick(row, ["poleHeight", "Height", "Altura"]),
        tipHeight: pick(row, ["tipHeight", "Tip Height"]),
        lowPower: pick(row, ["lowPower", "Low Power", "Low Power Attachment"], { contains: true }),
        maxCommHeight: pick(row, ["maxCommHeight", "Max Height", "Altura Max"]),
        topComm: pick(row, ["topComm", "Top Comm"]),
        lowComm: pick(row, ["lowComm", "Low Comm"]),
        owner: pick(row, ["owner", "Owner"]),
        poleType: pick(row, ["poleType", "Pole Type", "Type"]),
        isGenerated: String(pick(row, ["isGenerated", "Generated"])).toLowerCase() === "yes",
        notes: pick(row, ["notes", "Notas"])
      }));
    });

    rowsToObjects(spansSheet || []).forEach(row => {
      const spanId = String(pick(row, ["spanId", "Span ID"])).trim();
      if (!spanId) return;
      S().upsertSpan(S().createSpan(
        spanId,
        String(pick(row, ["fromPole", "From Pole", "Pole"])).trim(),
        String(pick(row, ["toPole", "To Pole", "Other Pole"])).trim(),
        pick(row, ["direction", "Direction", "Dir"]),
        pick(row, ["notes", "Notas"]),
        {
          type: pick(row, ["type", "Type"]),
          spanIndex: pick(row, ["spanIndex", "Span Index"]),
          length: pick(row, ["length", "Span Length"]),
          lengthDisplay: pick(row, ["lengthDisplay", "Span Length.display"], { contains: true }),
          environment: pick(row, ["environment", "Environment"]) || "NONE",
          environmentClearance: pick(row, ["environmentClearance", "Environment Clearance", "Env Clearance"]),
          midspanLowPower: pick(row, ["midspanLowPower", "Midspan Low Power"]),
          midspanMaxCommHeight: pick(row, ["midspanMaxCommHeight", "Max Midspan Comm", "Midspan Max Comm"]),
          bearingDegrees: pick(row, ["bearingDegrees", "Bearing Degrees"]),
          rawType: pick(row, ["rawType", "Raw Type"]),
          sourceCollectionId: pick(row, ["sourceCollectionId"]),
          linkedCollectionId: pick(row, ["linkedCollectionId"]),
          isGeneratedOtherPole: String(pick(row, ["isGeneratedOtherPole"])).toLowerCase() === "yes"
        }
      ));
    });

    rowsToObjects(spanSidesSheet || []).forEach(row => {
      S().upsertSpanSide({
        spanId: String(pick(row, ["spanId", "Span ID"])).trim(),
        poleId: String(pick(row, ["poleId", "Pole ID"])).trim(),
        proposedHOA: pick(row, ["proposedHOA", "Proposed HOA", "Proposed"]),
        proposedHOAChange: pick(row, ["proposedHOAChange", "Cambio Proposed", "Proposed Change"]),
        nextPoleProposedAuto: /^(true|yes|1|si|sí)$/i.test(String(pick(row, ["nextPoleProposedAuto", "Next Pole Proposed Auto"]))),
        proposedMidspan: pick(row, ["proposedMidspan", "Proposed Midspan"]),
        ocalcMS: pick(row, ["ocalcMS", "O-CALC MS", "O-Calc MS"]),
        msProposed: pick(row, ["msProposed", "MS Proposed"]),
        finalMidspan: pick(row, ["finalMidspan", "Midspan final ajustado", "Final Midspan"]),
        clearanceMSReason: pick(row, ["clearanceMSReason", "Clearance MS Reason"]),
        endDrop: pick(row, ["endDrop", "End Drop"]),
        clearanceReference: pick(row, ["clearanceReference", "Clearance Reference"]),
        maxCommHeight: pick(row, ["maxCommHeight", "Max Height"]),
        notes: pick(row, ["notes", "Notas"])
      });
    });

    rowsToObjects(findSheet(workbook, ["Comms"]) || []).forEach(row => {
      const poleId = String(pick(row, ["poleId", "Pole ID", "Id"])).trim();
      const owner = String(pick(row, ["owner", "Owner", "Comm"])).trim();
      if (poleId && owner) S().upsertComm(poleId, owner, pick(row, ["existingHOA", "Existing HOA", "Altura actual"]), pick(row, ["notes", "Notas"]), {
        ownerBase: pick(row, ["ownerBase", "Owner Base"]),
        existingHOAChange: pick(row, ["existingHOAChange", "Existing HOA Change"]),
        serviceDrop: /^(true|yes|1|si|sí)$/i.test(String(pick(row, ["serviceDrop", "Service Drop"]))),
        rawOwner: pick(row, ["rawOwner", "Raw Owner"]),
        unknownOwner: String(pick(row, ["unknownOwner", "Unknown Owner"])).toLowerCase() === "yes",
        size: pick(row, ["size", "Size"]),
        wireId: pick(row, ["wireId", "Wire ID"])
      });
    });

    rowsToObjects(spanCommsSheet || []).forEach(row => {
      S().upsertSpanComm({
        spanId: String(pick(row, ["spanId", "Span ID"])).trim(),
        poleId: String(pick(row, ["poleId", "Pole ID"])).trim(),
        owner: String(pick(row, ["owner", "Owner", "Comm"])).trim(),
        ownerBase: pick(row, ["ownerBase", "Owner Base"]),
        existingHOA: pick(row, ["existingHOA", "Existing HOA"]),
        existingHOAChange: pick(row, ["existingHOAChange", "Existing HOA Change"]),
        serviceDrop: /^(true|yes|1|si|sí)$/i.test(String(pick(row, ["serviceDrop", "Service Drop"]))),
        difference: pick(row, ["difference", "Difference"]),
        remotePoleId: pick(row, ["remotePoleId", "Remote Pole"]),
        remoteHOA: pick(row, ["remoteHOA", "Remote HOA"]),
        ocalcMS: pick(row, ["ocalcMS", "O-Calc MS"]),
        midspan: pick(row, ["midspan", "Imported Midspan", "Midspan"]),
        calculatedMidspan: pick(row, ["calculatedMidspan", "Calculated Midspan"]),
        msProposed: pick(row, ["msProposed", "MS Proposed"]),
        finalMidspan: pick(row, ["finalMidspan", "Final Midspan", "Midspan final ajustado"]),
        mr: pick(row, ["mr", "MR", "Make Ready"]),
        notes: pick(row, ["notes", "Notas"]),
        rawOwner: pick(row, ["rawOwner", "Raw Owner"]),
        unknownOwner: String(pick(row, ["unknownOwner", "Unknown Owner"])).toLowerCase() === "yes",
        size: pick(row, ["size", "Size"]),
        construction: pick(row, ["construction", "Construction"]),
        insulator: pick(row, ["insulator", "Insulator"]),
        wireId: pick(row, ["wireId", "Wire ID"]),
        wireIndex: pick(row, ["wireIndex", "Wire Index"])
      });
    });

    rowsToObjects(spanPowerSheet || []).forEach(row => {
      S().addSpanPower({
        spanId: String(pick(row, ["spanId", "Span ID"])).trim(),
        poleId: String(pick(row, ["poleId", "Pole ID"])).trim(),
        label: pick(row, ["label", "Type", "Tipo"]),
        attachmentHeight: pick(row, ["attachmentHeight", "Attachment Height"]),
        midspan: pick(row, ["midspan", "Midspan"]),
        size: pick(row, ["size", "Size"]),
        owner: pick(row, ["owner", "Owner"]),
        wireId: pick(row, ["wireId", "Wire ID"])
      });
    });

    state.makeReadyReferences = rowsToObjects(makeReadyRefsSheet || []).map(row => S().createMakeReadyReference({
      poleId: pick(row, ["poleId", "Pole ID", "Id"]),
      collectionId: pick(row, ["collectionId", "Collection ID"]),
      makeReadyIndex: pick(row, ["makeReadyIndex", "Make Ready Index"]),
      makeReadyId: pick(row, ["makeReadyId", "Make Ready ID"]),
      attachmentSizeRaw: pick(row, ["attachmentSizeRaw", "Attachment Size"]),
      attachmentMessenger: pick(row, ["attachmentMessenger", "Messenger"]),
      attachmentFiber: pick(row, ["attachmentFiber", "Fiber"]),
      attachmentDirection: pick(row, ["attachmentDirection", "Direction"]),
      attachmentDirectionTokens: String(pick(row, ["attachmentDirectionTokens", "Direction Tokens"])).split(/[|,]/).map(item => item.trim()).filter(Boolean),
      attachmentType: pick(row, ["attachmentType", "Attachment Type"]),
      attachmentHeight: pick(row, ["attachmentHeight", "Attachment Height"]),
      proposedMidspan: pick(row, ["proposedMidspan", "Proposed Midspan"]),
      makeReadyNotes: pick(row, ["makeReadyNotes", "Make Ready Notes"]),
      raw: row
    }));

    return S().normalizeState(state);
  }

  function buildCollectionIndex(collectionRows) {
    const byCollectionId = {};
    const byPoleId = {};

    collectionRows.forEach(row => {
      const collectionId = String(pick(row, ["collectionId", "Collection ID"])).trim();
      const poleId = String(pick(row, ["Id", "Pole ID", "PoleId", "PoleName", "Structure Number", "Pole"])).trim();
      const sequence = String(pick(row, ["Sequence", "Seq"])).trim();
      if (collectionId && poleId) byCollectionId[collectionId] = poleId;
      if (poleId) byPoleId[poleId] = { collectionId, sequence };
    });

    return { byCollectionId, byPoleId };
  }

  function importPolesFromCollection(collectionRows) {
    collectionRows.forEach(row => {
      const poleId = String(pick(row, ["Id", "Pole ID", "PoleId", "PoleName", "Structure Number", "Pole"])).trim();
      if (!poleId) return;

      const poleType = String(pick(row, ["Type", "Pole Type"])).trim();
      const lowPower = heightFromRow(
        row,
        ["Low Power Attachment.display", "Low Power Attachment Display", "Low Power Attachment", "Lowest Power.display", "Low Power.display"],
        ["Low Power Attachment", "Lowest Power", "Low Power"]
      );
      const tipHeight = heightFromRow(row, ["Tip.display", "Tip Display"], ["Tip"]);
      const poleHeight = parsePoleHeightFromType(poleType) || heightFromRow(row, ["Pole Height.display", "Height.display", "Length.display"], ["Pole Height", "Height", "Length"]);

      S().upsertPole(S().createPole({
        poleId,
        collectionId: String(pick(row, ["collectionId", "Collection ID"])).trim(),
        sequence: String(pick(row, ["Sequence", "Seq"])).trim(),
        poleHeight,
        tipHeight,
        lowPower,
        owner: String(pick(row, ["Owner"])).trim(),
        poleType,
        notes: "",
        metadata: {
          rawNote: String(pick(row, ["Note", "Notes", "Notas"])).trim(),
          status: pick(row, ["Status"]),
          vehicleAccessible: pick(row, ["Vehicle Accessible", "Vehicle Accessible (Y/N)"], { contains: true }),
          latitude: pick(row, ["Location.latitude"], { contains: true }),
          longitude: pick(row, ["Location.longitude"], { contains: true })
        }
      }));
    });
  }

  function buildSpanRecords(spanRows, collectionIndex) {
    // Span rows define the graph edges between poles. Unknown remote poles are
    // still created so the user can finish the missing data manually.
    return spanRows.map((row, index) => {
      const rawSpanId = String(pick(row, ["Span Id", "Span ID", "spanId", "Wire Span ID"])).trim() || `SPAN-${index + 1}`;
      const currentPole = String(pick(row, ["Id", "Pole ID", "Pole", "From Pole"])).trim()
        || collectionIndex.byCollectionId[String(pick(row, ["collectionId"])).trim()] || "";
      const linkedCollectionId = String(pick(row, ["Linked Collection.ID", "Linked Collection ID"], { contains: true })).trim();
      const linkedTitle = String(pick(row, ["Linked Collection.Title", "Linked Collection Title", "Other Pole", "To Pole", "Remote Pole"], { contains: true })).trim();
      const linkedPole = linkedTitle || collectionIndex.byCollectionId[linkedCollectionId] || `Unknown-${rawSpanId}`;
      const type = String(pick(row, ["Type"])).trim();
      const length = pick(row, ["Span Length"]);
      const lengthDisplay = heightFromRow(row, ["Span Length.display", "Span Length Display"], ["Span Length"]);
      const dir = directionFromBearingDisplay(pick(row, ["Span Length.bearing.display", "bearing.display"], { contains: true }));
      const environment = normalizeEnvironment(pick(row, ["Environment"]));

      return {
        row,
        index,
        rawSpanId,
        currentPole,
        linkedPole,
        linkedTitle,
        type,
        length,
        lengthDisplay,
        environment,
        environmentClearance: clearanceForEnvironment(environment),
        direction: dir.direction || String(pick(row, ["Direction", "Dir", "Bearing"])).trim(),
        bearingDegrees: dir.degrees,
        spanIndex: pick(row, ["Span Index"]),
        sourceCollectionId: String(pick(row, ["collectionId"])).trim(),
        linkedCollectionId,
        note: "",
        isGeneratedOtherPole: !linkedTitle && !collectionIndex.byCollectionId[linkedCollectionId]
      };
    }).filter(record => record.currentPole);
  }

  function importSpans(spanRecords) {
    const rawSpanToSpanId = {};
    spanRecords.forEach(record => {
      const fromPole = record.currentPole || `Unknown-FROM-${record.rawSpanId}`;
      const toPole = record.linkedPole || `Unknown-${record.rawSpanId}`;

      if (!S().getPole(fromPole)) S().upsertPole(S().createPole({ poleId: fromPole, isGenerated: /^Unknown-/i.test(fromPole) }));
      if (!S().getPole(toPole)) S().upsertPole(S().createPole({ poleId: toPole, isGenerated: record.isGeneratedOtherPole || /^Unknown-/i.test(toPole) }));

      S().upsertSpan(S().createSpan(
        record.rawSpanId,
        fromPole,
        toPole,
        record.direction,
        record.note,
        {
          type: record.type,
          rawType: record.type,
          spanIndex: record.spanIndex,
          length: record.length,
          lengthDisplay: record.lengthDisplay,
          environment: record.environment,
          environmentClearance: record.environmentClearance,
          bearingDegrees: record.bearingDegrees,
          rawSpanId: record.rawSpanId,
          linkedCollectionId: record.linkedCollectionId,
          sourceCollectionId: record.sourceCollectionId,
          isGeneratedOtherPole: record.isGeneratedOtherPole
        }
      ));
      rawSpanToSpanId[record.rawSpanId] = record.rawSpanId;
    });
    return rawSpanToSpanId;
  }

  function importSpanWires(wireRows, rawSpanToSpanId) {
    wireRows.forEach(row => {
      const rawSpanId = String(pick(row, ["Span Id", "Span ID", "spanId", "Wire Span ID"])).trim();
      const spanId = rawSpanToSpanId[rawSpanId] || rawSpanId;
      if (!spanId) return;

      let span = S().getSpan(spanId);
      const poleId = String(pick(row, ["Id", "Pole ID", "Pole", "CollectionId", "Structure Number"])).trim();
      if (!poleId) return;

      if (!span) {
        const unknownPole = `Unknown-${spanId}`;
        S().upsertSpan(S().createSpan(spanId, poleId, unknownPole, "", "", { rawSpanId, isGeneratedOtherPole: true }));
        span = S().getSpan(spanId);
      }
      if (!S().getPole(poleId)) S().upsertPole(S().createPole(poleId));

      const rawOwner = String(pick(row, ["Owner", "owner"])).trim();
      const size = String(pick(row, ["Size", "Size.display", "Wire Size"])).trim();
      const construction = String(pick(row, ["Construction"])).trim();
      const insulator = String(pick(row, ["Insulator"])).trim();
      const wireId = String(pick(row, ["Wire Id", "Wire ID", "wireId"])).trim();
      const wireIndex = String(pick(row, ["Wire Index"])).trim();
      const attachmentHeight = heightFromRow(row, ["Attachment Height.display", "Attachment Height Display", "Existing HOA", "HOA"], ["Attachment Height", "Height", "Wire Height"]);
      const midspan = heightFromRow(row, ["Mid Span Height.display", "Midspan.display", "Midspan"], ["Mid Span Height", "Midspan"]);

      if (isPowerWire(row)) {
        S().addSpanPower({
          spanId,
          poleId,
          label: /primary/i.test(size) ? "Primary" : /secondary/i.test(size) ? "Secondary" : /neutral/i.test(size) ? "Neutral" : "Power",
          attachmentHeight,
          midspan,
          size,
          owner: rawOwner,
          wireId
        });
        return;
      }

      if (!isCommunicationWire(row) && !rawOwner) return;
      const owner = rawOwner || (wireId ? `UNKNOWN-${wireId}` : `UNKNOWN-${spanId}-${poleId}`);
      const ownerBase = owner;
      const unknownOwner = !rawOwner;
      const serviceDrop = /communication\s*drops?|service\s*drop|\bdrop\b/i.test(size);
      S().upsertComm(poleId, owner, attachmentHeight, "", { rawOwner, size, wireId, ownerBase, unknownOwner });
      S().upsertSpanComm({
        spanId,
        poleId,
        owner,
        ownerBase,
        existingHOA: attachmentHeight,
        existingHOAChange: "",
        serviceDrop,
        remotePoleId: span ? S().getOtherPoleId(span, poleId) : "",
        ocalcMS: midspan,
        midspan,
        calculatedMidspan: "",
        mr: "",
        notes: "",
        rawOwner,
        unknownOwner,
        size,
        construction,
        insulator,
        wireId,
        wireIndex
      });
    });
  }

  function importOriginalWorkbook(workbook, fileName) {
    const state = S().resetState();
    state.importedFileName = fileName || "Excel original";
    state.importedAt = new Date().toISOString();
    state.autoCreateSpanComms = false;

    const collectionRows = rowsToObjects(findSheet(workbook, ["Collection", "Poles", "Postes"]) || []);
    const spanRows = rowsToObjects(findSheet(workbook, ["Span", "Spans"]) || []);
    const wireRows = rowsToObjects(findSheet(workbook, ["Span.Wire", "Span Wire", "Wires", "Comms"]) || []);
    const makeReadyRows = rowsToObjects(findSheet(workbook, ["Make Ready", "MakeReady", "MR"]) || []);

    if (!collectionRows.length && !spanRows.length && !wireRows.length) {
      throw new Error("Could not find readable Collection, Span or Span.Wire sheets.");
    }

    const collectionIndex = buildCollectionIndex(collectionRows);
    importPolesFromCollection(collectionRows);
    state.poleClassChecks = buildPoleClassChecks(collectionRows);

    const spanRecords = buildSpanRecords(spanRows, collectionIndex);
    const rawSpanToSpanId = importSpans(spanRecords);
    importSpanWires(wireRows, rawSpanToSpanId);
    state.makeReadyReferences = importMakeReadyReferences(makeReadyRows);

    S().normalizeState(state);
    global.Calculations.recalculateAll();

    if (!Object.keys(S().getState().poles).length) {
      throw new Error("Could not load any pole from the file.");
    }

    return S().getState();
  }

  async function importJsonFile(file) {
    const payload = JSON.parse(await file.text());
    const nextState = payload && payload.state ? payload.state : payload;
    if (!nextState || typeof nextState !== "object" || !nextState.poles || !nextState.spans) {
      throw new Error("The JSON does not contain valid calculator data.");
    }
    const restored = S().setState({
      ...nextState,
      importedFileName: file.name || nextState.importedFileName || "Imported data",
      importedAt: new Date().toISOString()
    });
    global.Calculations.recalculateAll();
    return restored;
  }

  async function importExcelFile(file) {
    const ext = file.name.split(".").pop().toLowerCase();
    if (ext === "csv") {
      const text = await file.text();
      const rows = text.split(/\r?\n/).filter(Boolean).map(line => line.split(",").map(v => v.replace(/^"|"$/g, "")));
      const workbook = { SheetNames: ["CSV"], Sheets: { CSV: rows } };
      return importOriginalWorkbook(workbook, file.name);
    }
    const workbook = await global.MiniXLSX.readArrayBuffer(await file.arrayBuffer());
    return importAppStateSheet(workbook) || importExportedTables(workbook) || importOriginalWorkbook(workbook, file.name);
  }

  global.ExcelImport = {
    importExcelFile,
    importJsonFile,
    importDataFile: file => {
      const ext = file.name.split(".").pop().toLowerCase();
      return ext === "json" || file.type === "application/json" ? importJsonFile(file) : importExcelFile(file);
    },
    rowsToObjects,
    pick,
    findSheet,
    normalizeHeaderName,
    directionFromBearingDisplay,
    recalculatePoleClassCheck,
    ANSI_POLE_CLASSES,
    ANSI_CLASS_TABLE,
    ANSI_APPROX_GROUNDLINE_DISTANCE
  };
})(window);
