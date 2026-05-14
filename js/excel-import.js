(function (global) {
  "use strict";

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
    if (!polesSheet && !spansSheet && !spanCommsSheet && !spanSidesSheet) return null;

    const state = S().resetState();
    state.importedFileName = "Excel exportado reimportado";
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
        proposedMidspan: pick(row, ["proposedMidspan", "Proposed Midspan", "O-Calc MS"]),
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
        difference: pick(row, ["difference", "Difference"]),
        remotePoleId: pick(row, ["remotePoleId", "Remote Pole"]),
        remoteHOA: pick(row, ["remoteHOA", "Remote HOA"]),
        ocalcMS: pick(row, ["ocalcMS", "O-Calc MS"]),
        midspan: pick(row, ["midspan", "Imported Midspan", "Midspan"]),
        calculatedMidspan: pick(row, ["calculatedMidspan", "Calculated Midspan"]),
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
        environment: "NONE",
        environmentClearance: "",
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

      const rawOwner = String(pick(row, ["Owner", "owner", "Company", "Communication Owner"])).trim();
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

      if (!isCommunicationWire(row) && !rawOwner && !size) return;
      const normalizedOwner = normalizeOwner(rawOwner, size);
      const owner = normalizedOwner || fallbackOwner(rawOwner, size, wireId);
      const ownerBase = normalizedOwner || owner;
      const unknownOwner = !normalizedOwner;
      const notes = [
        size ? `Size: ${size}` : "",
        rawOwner ? `Raw owner: ${rawOwner}` : "",
        unknownOwner ? "Owner no normalizado" : "",
        construction ? `Construction: ${construction}` : "",
        insulator ? `Insulator: ${insulator}` : "",
        wireId ? `WireId: ${wireId}` : ""
      ].filter(Boolean).join(" | ");

      S().upsertComm(poleId, owner, attachmentHeight, "", { rawOwner, size, wireId, ownerBase, unknownOwner });
      S().upsertSpanComm({
        spanId,
        poleId,
        owner,
        ownerBase,
        existingHOA: attachmentHeight,
        existingHOAChange: "",
        remotePoleId: span ? S().getOtherPoleId(span, poleId) : "",
        ocalcMS: midspan,
        midspan,
        calculatedMidspan: "",
        mr: "",
        notes,
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

    if (!collectionRows.length && !spanRows.length && !wireRows.length) {
      throw new Error("No encontre hojas Collection, Span o Span.Wire con encabezados legibles.");
    }

    const collectionIndex = buildCollectionIndex(collectionRows);
    importPolesFromCollection(collectionRows);

    const spanRecords = buildSpanRecords(spanRows, collectionIndex);
    const rawSpanToSpanId = importSpans(spanRecords);
    importSpanWires(wireRows, rawSpanToSpanId);

    S().normalizeState(state);
    global.Calculations.recalculateAll();

    if (!Object.keys(S().getState().poles).length) {
      throw new Error("No se pudo cargar ningun poste desde el archivo.");
    }

    return S().getState();
  }

  async function importJsonFile(file) {
    const payload = JSON.parse(await file.text());
    const nextState = payload && payload.state ? payload.state : payload;
    if (!nextState || typeof nextState !== "object" || !nextState.poles || !nextState.spans) {
      throw new Error("El JSON no contiene datos validos de la calculadora.");
    }
    const restored = S().setState({
      ...nextState,
      importedFileName: file.name || nextState.importedFileName || "Datos importados",
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
    directionFromBearingDisplay
  };
})(window);
