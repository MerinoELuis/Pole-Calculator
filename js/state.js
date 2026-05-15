(function (global) {
  "use strict";

  // AppStore is the single source of truth for the calculator. UI modules read
  // from this state, and calculation modules write derived values back into it.
  const CURRENT_VERSION = "1.3.0";
  const STORAGE_KEY = "poleCalculatorAppState.v2";

  const DEFAULT_CLEARANCE_TO_POWER = "40\"";
  const DEFAULT_COMM_CLEARANCE = "12\"";
  const DEFAULT_BOLT_CLEARANCE = "4\"";
  const ENVIRONMENT_OPTIONS = [
    { value: "NONE", label: "None", clearance: "" },
    { value: "STREET", label: "Street", clearance: "15'6\"" },
    { value: "HIGHWAY", label: "Highway", clearance: "18'" },
    { value: "PEDESTRIAN", label: "Pedestrian", clearance: "10'" },
    { value: "PARALLEL_TO_STREET", label: "Parallel to street", clearance: "15'6\"" },
    { value: "OBSTRUCTED_PARALLEL_TO_STREET", label: "Obstructed parallel to street", clearance: "15'6\"" },
    { value: "UNLIKELY_PARALLEL_TO_STREET", label: "Unlikely parallel to street", clearance: "15'6\"" },
    { value: "RESIDENTIAL_DRIVEWAY", label: "Residential driveway", clearance: "15'6\"" },
    { value: "COMMERCIAL_DRIVEWAY", label: "Commercial driveway", clearance: "15'6\"" },
    { value: "PARKING_LOT", label: "Parking lot", clearance: "15'6\"" },
    { value: "ALLEY", label: "Alley", clearance: "15'6\"" },
    { value: "RAILROAD", label: "Railroad", clearance: "25'" },
    { value: "RURAL", label: "Rural", clearance: "10'" },
    { value: "FARM", label: "Farm", clearance: "15'6\"" },
    { value: "WATER_WITH_SAILBOATS", label: "Water with sailboats", clearance: "Variable" },
    { value: "WATER_WITHOUT_SAILBOATS", label: "Water without sailboats", clearance: "15'" },
    { value: "TROLLEY", label: "Trolley", clearance: "25'" }
  ];

  // Keep startup empty. Imported Excel/JSON or local saves are the only things
  // that should populate the workspace in normal use.
  const emptyState = () => ({
    version: CURRENT_VERSION,
    importedFileName: "",
    importedAt: new Date().toISOString(),
    selectedPoleId: "",
    selectedSpanId: "",
    autoCreateSpanComms: false,
    settings: {
      clearanceToPower: DEFAULT_CLEARANCE_TO_POWER,
      polePowerCommsClearance: "40\"",
      commClearance: DEFAULT_COMM_CLEARANCE,
      boltClearance: DEFAULT_BOLT_CLEARANCE,
      midspanPowerCommClearance: "30\"",
      midspanCommCommClearance: "4\"",
      position: "TOP_COMM",
      mrCase: "LOWER",
      proposedOwner: "Wecom",
      environmentClearance: "15'6\"",
      sagPer100Ft: "1'"
    },
    poles: {},
    spans: {},
    spanSides: {},
    spanComms: {},
    spanPower: {},
    movements: [],
    mr: [],
    warnings: [],
    ui: {
      search: "",
      filter: "all"
    }
  });

  let state = emptyState();

  function trim(value) {
    return String(value ?? "").trim();
  }

  function keyForSpanSide(spanId, poleId) {
    return [spanId, poleId].map(v => trim(v)).join("__");
  }

  function keyForSpanComm(spanId, poleId, owner, wireId) {
    return [spanId, poleId, owner, wireId || ""].map(v => trim(v)).join("__");
  }

  function createPole(poleId, poleHeight = "", notes = "", extra = {}) {
    // Constructors normalize all incoming data so later code can assume each
    // entity has the same shape, whether it came from Excel, JSON, or the UI.
    if (typeof poleId === "object" && poleId !== null) {
      const data = poleId;
      return {
        poleId: trim(data.poleId || data.id || ""),
        collectionId: trim(data.collectionId || ""),
        sequence: trim(data.sequence || ""),
        poleHeight: trim(data.poleHeight || ""),
        tipHeight: trim(data.tipHeight || ""),
        lowPower: trim(data.lowPower || ""),
        maxCommHeight: trim(data.maxCommHeight || ""),
        topComm: trim(data.topComm || ""),
        lowComm: trim(data.lowComm || ""),
        owner: trim(data.owner || ""),
        poleType: trim(data.poleType || ""),
        isGenerated: Boolean(data.isGenerated),
        ugActive: Boolean(data.ugActive),
        ugReason: trim(data.ugReason || ""),
        pcoActive: Boolean(data.pcoActive),
        pcoScope: trim(data.pcoScope || ""),
        pcoType: trim(data.pcoType || ""),
        pcoDetail: trim(data.pcoDetail || ""),
        notes: trim(data.notes || ""),
        comms: Array.isArray(data.comms) ? data.comms : [],
        metadata: data.metadata || {}
      };
    }

    return {
      poleId: trim(poleId),
      collectionId: trim(extra.collectionId || ""),
      sequence: trim(extra.sequence || ""),
      poleHeight: trim(poleHeight),
      tipHeight: trim(extra.tipHeight || ""),
      lowPower: trim(extra.lowPower || ""),
      maxCommHeight: trim(extra.maxCommHeight || ""),
      topComm: trim(extra.topComm || ""),
      lowComm: trim(extra.lowComm || ""),
      owner: trim(extra.owner || ""),
      poleType: trim(extra.poleType || ""),
      isGenerated: Boolean(extra.isGenerated),
      ugActive: Boolean(extra.ugActive),
      ugReason: trim(extra.ugReason || ""),
      pcoActive: Boolean(extra.pcoActive),
      pcoScope: trim(extra.pcoScope || ""),
      pcoType: trim(extra.pcoType || ""),
      pcoDetail: trim(extra.pcoDetail || ""),
      notes: trim(notes),
      comms: [],
      metadata: extra.metadata || {}
    };
  }

  function createComm(owner, existingHOA = "", notes = "", extra = {}) {
    return {
      owner: trim(owner),
      ownerBase: trim(extra.ownerBase || owner),
      existingHOA: trim(existingHOA),
      existingHOAChange: trim(extra.existingHOAChange || ""),
      notes: trim(notes),
      rawOwner: trim(extra.rawOwner || ""),
      unknownOwner: Boolean(extra.unknownOwner),
      size: trim(extra.size || ""),
      wireId: trim(extra.wireId || "")
    };
  }

  function createSpan(spanId, fromPole, toPole = "", direction = "", notes = "", extra = {}) {
    const id = trim(spanId);
    const fallbackToPole = trim(toPole) || `Unknown-${id || Math.random().toString(36).slice(2, 8)}`;
    return {
      spanId: id,
      fromPole: trim(fromPole),
      toPole: fallbackToPole,
      direction: trim(direction),
      bearingDegrees: extra.bearingDegrees === null || extra.bearingDegrees === undefined ? "" : extra.bearingDegrees,
      notes: trim(notes),
      type: trim(extra.type || ""),
      spanIndex: trim(extra.spanIndex || ""),
      length: trim(extra.length || ""),
      lengthDisplay: trim(extra.lengthDisplay || ""),
      environment: trim(extra.environment || "NONE"),
      environmentClearance: trim(extra.environmentClearance || ""),
      midspanLowPower: trim(extra.midspanLowPower || ""),
      midspanMaxCommHeight: trim(extra.midspanMaxCommHeight || ""),
      rawSpanIds: Array.isArray(extra.rawSpanIds) ? extra.rawSpanIds : (extra.rawSpanId ? [extra.rawSpanId] : []),
      rawType: trim(extra.rawType || ""),
      linkedCollectionId: trim(extra.linkedCollectionId || ""),
      sourceCollectionId: trim(extra.sourceCollectionId || ""),
      isGeneratedOtherPole: Boolean(extra.isGeneratedOtherPole)
    };
  }

  function createSpanSide(data = {}) {
    return {
      spanId: trim(data.spanId || ""),
      poleId: trim(data.poleId || ""),
      proposedHOA: trim(data.proposedHOA || ""),
      proposedHOAChange: trim(data.proposedHOAChange || ""),
      nextPoleProposedAuto: Boolean(data.nextPoleProposedAuto),
      proposedMidspan: trim(data.proposedMidspan || ""),
      ocalcMS: trim(data.ocalcMS || data["O-CALC MS"] || ""),
      msProposed: trim(data.msProposed || data.proposedMidspan || ""),
      finalMidspan: trim(data.finalMidspan || ""),
      clearanceMSStatus: trim(data.clearanceMSStatus || ""),
      clearanceMSMessage: trim(data.clearanceMSMessage || ""),
      clearanceMSReason: trim(data.clearanceMSReason || ""),
      clearanceMSIssue: Boolean(data.clearanceMSIssue),
      proposedFlaggingStatus: trim(data.proposedFlaggingStatus || ""),
      proposedFlaggingMessage: trim(data.proposedFlaggingMessage || ""),
      pendingMidspanFinal: trim(data.pendingMidspanFinal || ""),
      clearanceFixReadyAt: Number(data.clearanceFixReadyAt || 0),
      endDrop: trim(data.endDrop || ""),
      clearanceReference: trim(data.clearanceReference || "LOW_POWER"),
      maxCommHeight: trim(data.maxCommHeight || ""),
      topComm: trim(data.topComm || ""),
      lowComm: trim(data.lowComm || ""),
      notes: trim(data.notes || ""),
      lockedEndDrop: Boolean(data.lockedEndDrop),
      updatedAt: data.updatedAt || ""
    };
  }

  function createSpanComm(data = {}) {
    return {
      spanId: trim(data.spanId || ""),
      poleId: trim(data.poleId || ""),
      owner: trim(data.owner || ""),
      ownerBase: trim(data.ownerBase || data.owner || ""),
      existingHOA: trim(data.existingHOA || ""),
      existingHOAChange: trim(data.existingHOAChange || ""),
      serviceDrop: Boolean(data.serviceDrop),
      difference: trim(data.difference || ""),
      remotePoleId: trim(data.remotePoleId || ""),
      remoteHOA: trim(data.remoteHOA || ""),
      ocalcMS: trim(data.ocalcMS || ""),
      midspan: trim(data.midspan || ""),
      calculatedMidspan: trim(data.calculatedMidspan || ""),
      msProposed: trim(data.msProposed || data.calculatedMidspan || data.midspan || ""),
      finalMidspan: trim(data.finalMidspan || ""),
      clearanceMSStatus: trim(data.clearanceMSStatus || ""),
      clearanceMSMessage: trim(data.clearanceMSMessage || ""),
      clearanceMSIssue: Boolean(data.clearanceMSIssue),
      pendingMidspanFinal: trim(data.pendingMidspanFinal || ""),
      clearanceFixReadyAt: Number(data.clearanceFixReadyAt || 0),
      flaggingStatus: trim(data.flaggingStatus || ""),
      flaggingMessage: trim(data.flaggingMessage || ""),
      mr: trim(data.mr || ""),
      notes: trim(data.notes || ""),
      rawOwner: trim(data.rawOwner || ""),
      unknownOwner: Boolean(data.unknownOwner),
      size: trim(data.size || ""),
      construction: trim(data.construction || ""),
      insulator: trim(data.insulator || ""),
      wireId: trim(data.wireId || ""),
      wireIndex: trim(data.wireIndex || ""),
      updatedAt: data.updatedAt || ""
    };
  }

  function createSpanPower(data = {}) {
    return {
      spanId: trim(data.spanId || ""),
      poleId: trim(data.poleId || ""),
      label: trim(data.label || ""),
      attachmentHeight: trim(data.attachmentHeight || ""),
      midspan: trim(data.midspan || ""),
      size: trim(data.size || ""),
      owner: trim(data.owner || ""),
      wireId: trim(data.wireId || "")
    };
  }

  function setState(nextState) {
    state = normalizeState(nextState || emptyState());
    return state;
  }

  function getState() { return state; }

  function resetState() {
    state = emptyState();
    return state;
  }

  function getPole(poleId) { return state.poles[poleId] || null; }
  function getSpan(spanId) { return state.spans[spanId] || null; }

  function getSpanSide(spanId, poleId) {
    return state.spanSides[keyForSpanSide(spanId, poleId)] || null;
  }

  function getSpanComm(spanId, poleId, owner, wireId = "") {
    if (wireId) return state.spanComms[keyForSpanComm(spanId, poleId, owner, wireId)] || null;
    const exact = state.spanComms[keyForSpanComm(spanId, poleId, owner, "")];
    if (exact) return exact;
    return Object.values(state.spanComms).find(sc => sc.spanId === spanId && sc.poleId === poleId && sc.owner === owner) || null;
  }

  function upsertPole(pole) {
    if (!pole || !pole.poleId) return null;
    const existing = state.poles[pole.poleId] || createPole(pole.poleId);
    state.poles[pole.poleId] = {
      ...existing,
      ...pole,
      comms: Array.isArray(pole.comms) ? pole.comms : existing.comms || [],
      metadata: { ...(existing.metadata || {}), ...(pole.metadata || {}) }
    };
    return state.poles[pole.poleId];
  }

  function updatePoleField(poleId, field, value) {
    const pole = state.poles[poleId];
    if (!pole) return null;
    if (!["poleHeight", "lowPower", "maxCommHeight", "topComm", "lowComm", "notes", "sequence"].includes(field)) return pole;
    pole[field] = trim(value);
    return pole;
  }

  function updateSetting(field, value) {
    if (!state.settings || !Object.prototype.hasOwnProperty.call(emptyState().settings, field)) return null;
    state.settings[field] = trim(value);
    if (field === "polePowerCommsClearance") state.settings.clearanceToPower = trim(value);
    if (field === "clearanceToPower") state.settings.polePowerCommsClearance = trim(value);
    return state.settings;
  }

  function updateSpanField(spanId, field, value) {
    const span = state.spans[spanId];
    if (!span) return null;
    if (!["environment", "environmentClearance", "midspanLowPower", "midspanMaxCommHeight", "notes"].includes(field)) return span;
    span[field] = trim(value);
    return span;
  }

  function upsertComm(poleId, owner, existingHOA = "", notes = "", extra = {}) {
    const pole = state.poles[poleId] || upsertPole(createPole(poleId));
    const ownerKey = trim(owner);
    const wireKey = trim(extra.wireId || "");
    let idx = pole.comms.findIndex(c => c.owner === ownerKey && (!wireKey || c.wireId === wireKey));
    if (idx < 0) idx = pole.comms.findIndex(c => c.owner === ownerKey && !c.wireId && !wireKey);

    if (idx >= 0) {
      pole.comms[idx] = {
        ...pole.comms[idx],
        owner: ownerKey,
        ownerBase: trim(extra.ownerBase || pole.comms[idx].ownerBase || ownerKey),
        existingHOA: trim(existingHOA) || pole.comms[idx].existingHOA,
        existingHOAChange: trim(extra.existingHOAChange || pole.comms[idx].existingHOAChange || ""),
        notes: trim(notes) || pole.comms[idx].notes,
        rawOwner: trim(extra.rawOwner || pole.comms[idx].rawOwner || ""),
        unknownOwner: Boolean(extra.unknownOwner || pole.comms[idx].unknownOwner),
        size: trim(extra.size || pole.comms[idx].size || ""),
        wireId: trim(extra.wireId || pole.comms[idx].wireId || "")
      };
      return pole.comms[idx];
    }
    const comm = createComm(ownerKey, existingHOA, notes, extra);
    pole.comms.push(comm);
    return comm;
  }

  function upsertSpan(span) {
    if (!span || !span.spanId) return null;
    const existing = state.spans[span.spanId] || createSpan(span.spanId, span.fromPole, span.toPole);
    const rawSpanIds = Array.from(new Set([...(existing.rawSpanIds || []), ...(span.rawSpanIds || [])].filter(Boolean)));
    state.spans[span.spanId] = { ...existing, ...span, rawSpanIds };
    ensureSpanSideForSpan(state.spans[span.spanId]);
    return state.spans[span.spanId];
  }

  function upsertSpanSide(data) {
    if (!data || !data.spanId || !data.poleId) return null;
    const key = keyForSpanSide(data.spanId, data.poleId);
    const existing = state.spanSides[key] || createSpanSide(data);
    state.spanSides[key] = createSpanSide({ ...existing, ...data, updatedAt: new Date().toISOString() });
    return state.spanSides[key];
  }

  function upsertSpanComm(data) {
    if (!data || !data.spanId || !data.poleId || !data.owner) return null;
    const key = keyForSpanComm(data.spanId, data.poleId, data.owner, data.wireId || "");
    const existing = state.spanComms[key] || createSpanComm(data);
    state.spanComms[key] = createSpanComm({ ...existing, ...data, updatedAt: new Date().toISOString() });
    return state.spanComms[key];
  }

  function removeSpanComm(spanId, poleId, owner, wireId = "") {
    const key = keyForSpanComm(spanId, poleId, owner, wireId || "");
    delete state.spanComms[key];
    const pole = state.poles[poleId];
    const stillUsed = Object.values(state.spanComms).some(sc =>
      sc.poleId === poleId && sc.owner === owner && (!wireId || sc.wireId === wireId)
    );
    if (pole && !stillUsed) {
      pole.comms = (pole.comms || []).filter(comm => !(comm.owner === owner && (!wireId || comm.wireId === wireId)));
    }
  }

  function addSpanPower(data) {
    const clean = createSpanPower(data);
    if (!clean.spanId || !clean.poleId) return null;
    const key = [clean.spanId, clean.poleId, clean.wireId || clean.label || Object.keys(state.spanPower).length].join("__");
    state.spanPower[key] = clean;
    return clean;
  }

  function updateSpanPowerField(powerKey, field, value) {
    if (!["attachmentHeight", "midspan", "size", "label"].includes(field)) return null;
    const row = state.spanPower[trim(powerKey)];
    if (!row) return null;
    row[field] = trim(value);
    return row;
  }

  function getConnectedSpans(poleId) {
    return Object.values(state.spans).filter(span => span.fromPole === poleId || span.toPole === poleId);
  }

  function getOtherPoleId(span, poleId) {
    if (!span) return "";
    if (span.fromPole === poleId) return span.toPole;
    if (span.toPole === poleId) return span.fromPole;
    return "";
  }

  function getSpanSidesForPole(poleId) {
    return Object.values(state.spanSides).filter(side => side.poleId === poleId);
  }

  function getSpanSidesForSpan(spanId) {
    return Object.values(state.spanSides).filter(side => side.spanId === spanId);
  }

  function getSpanCommsForPole(poleId) {
    return Object.values(state.spanComms).filter(sc => sc.poleId === poleId);
  }

  function getSpanCommsForSpan(spanId) {
    return Object.values(state.spanComms).filter(sc => sc.spanId === spanId);
  }

  function getSpanPowerForPole(poleId) {
    return Object.entries(state.spanPower)
      .filter(([, row]) => row.poleId === poleId)
      .map(([key, row]) => ({ ...row, key }));
  }

  function getSpanPowerForSpan(spanId) {
    return Object.values(state.spanPower).filter(row => row.spanId === spanId);
  }

  function poleHasChanges(poleId) {
    const sideChange = getSpanSidesForPole(poleId).some(side => side.proposedHOA || side.proposedMidspan || side.ocalcMS || side.msProposed || side.finalMidspan || side.endDrop || side.notes);
    const commChange = getSpanCommsForPole(poleId).some(sc => sc.existingHOAChange || sc.notes || sc.mr);
    return sideChange || commChange;
  }

  function ensureUnknownPoles() {
    Object.values(state.spans).forEach(span => {
      if (span.fromPole && !state.poles[span.fromPole]) upsertPole(createPole({ poleId: span.fromPole, isGenerated: /^Unknown-/i.test(span.fromPole) }));
      if (span.toPole && !state.poles[span.toPole]) upsertPole(createPole({ poleId: span.toPole, isGenerated: /^Unknown-/i.test(span.toPole) }));
    });
  }

  function ensureSpanSideForSpan(span) {
    if (!span || !span.spanId) return;
    if (span.fromPole) upsertSpanSide({ spanId: span.spanId, poleId: span.fromPole });
    if (span.toPole) upsertSpanSide({ spanId: span.spanId, poleId: span.toPole });
  }

  function ensureSpanSides() {
    Object.values(state.spans).forEach(ensureSpanSideForSpan);
  }

  function ensureSpanComms() {
    Object.values(state.poles).forEach(pole => {
      const spans = getConnectedSpans(pole.poleId);
      pole.comms.forEach(comm => {
        spans.forEach(span => {
          const existing = getSpanComm(span.spanId, pole.poleId, comm.owner, comm.wireId || "");
          if (!existing) {
            upsertSpanComm({
              spanId: span.spanId,
              poleId: pole.poleId,
              owner: comm.owner,
              ownerBase: comm.ownerBase || comm.owner,
              existingHOA: comm.existingHOA,
              existingHOAChange: comm.existingHOAChange || "",
              rawOwner: comm.rawOwner || "",
              unknownOwner: Boolean(comm.unknownOwner),
              size: comm.size || "",
              wireId: comm.wireId || ""
            });
          } else if (!existing.existingHOA && comm.existingHOA) {
            upsertSpanComm({ ...existing, existingHOA: comm.existingHOA });
          }
        });
      });
    });
  }

  function ensureEndpointComms() {
    // Si un extremo del span no trae filas propias, se crean comms placeholder
    // con los mismos owners/wireIds que llegan desde el otro extremo. El usuario
    // completará las alturas; así el poste nunca queda visualmente vacío.
    Object.values(state.spans).forEach(span => {
      [span.fromPole, span.toPole].filter(Boolean).forEach(poleId => {
        const rowsAtPole = getSpanCommsForPole(poleId);
        if (rowsAtPole.length) return;
        const otherPoleId = getOtherPoleId(span, poleId);
        getSpanCommsForPole(otherPoleId)
          .filter(row => row.spanId === span.spanId)
          .forEach(row => {
            upsertComm(poleId, row.owner, "", "", {
              ownerBase: row.ownerBase || row.owner,
              rawOwner: row.rawOwner || row.owner,
              wireId: row.wireId || "",
              size: row.size || ""
            });
            upsertSpanComm({
              spanId: span.spanId,
              poleId,
              owner: row.owner,
              ownerBase: row.ownerBase || row.owner,
              existingHOA: "",
              existingHOAChange: "",
              serviceDrop: Boolean(row.serviceDrop),
              rawOwner: row.rawOwner || row.owner,
              size: row.size || "",
              wireId: row.wireId || ""
            });
          });
      });
    });
  }

  function normalizeState(raw) {
    const next = { ...emptyState(), ...raw };
    next.settings = { ...emptyState().settings, ...(raw && raw.settings ? raw.settings : {}) };
    next.poles = next.poles || {};
    next.spans = next.spans || {};
    next.spanSides = next.spanSides || {};
    next.spanComms = next.spanComms || {};
    next.spanPower = next.spanPower || {};
    next.movements = Array.isArray(next.movements) ? next.movements : [];
    next.mr = Array.isArray(next.mr) ? next.mr : [];
    next.warnings = Array.isArray(next.warnings) ? next.warnings : [];
    next.ui = { search: "", filter: "all", ...(next.ui || {}) };

    Object.keys(next.poles).forEach(id => {
      const pole = next.poles[id];
      next.poles[id] = {
        ...createPole(id),
        ...pole,
        poleId: pole.poleId || id,
        comms: Array.isArray(pole.comms) ? pole.comms.map(c => createComm(c.owner, c.existingHOA, c.notes, c)) : []
      };
    });

    Object.keys(next.spans).forEach(id => {
      const span = next.spans[id];
      next.spans[id] = { ...createSpan(id, span.fromPole, span.toPole), ...span, spanId: span.spanId || id };
    });

    const normalizedSides = {};
    Object.values(next.spanSides).forEach(side => {
      const clean = createSpanSide(side);
      if (clean.spanId && clean.poleId) normalizedSides[keyForSpanSide(clean.spanId, clean.poleId)] = clean;
    });
    next.spanSides = normalizedSides;

    const normalizedSpanComms = {};
    Object.values(next.spanComms).forEach(sc => {
      const clean = createSpanComm(sc);
      if (clean.spanId && clean.poleId && clean.owner) normalizedSpanComms[keyForSpanComm(clean.spanId, clean.poleId, clean.owner, clean.wireId || "")] = clean;
    });
    next.spanComms = normalizedSpanComms;

    const normalizedPower = {};
    Object.values(next.spanPower || {}).forEach(row => {
      const clean = createSpanPower(row);
      if (clean.spanId && clean.poleId) normalizedPower[[clean.spanId, clean.poleId, clean.wireId || clean.label || Object.keys(normalizedPower).length].join("__")] = clean;
    });
    next.spanPower = normalizedPower;

    state = next;
    ensureUnknownPoles();
    ensureSpanSides();
    if (state.autoCreateSpanComms) ensureSpanComms();
    ensureEndpointComms();
    return state;
  }

  function saveToLocal() {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
  }

  function loadFromLocal() {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (!raw) return null;
    return setState(JSON.parse(raw));
  }

  global.AppStore = {
    CURRENT_VERSION,
    STORAGE_KEY,
    getState,
    setState,
    resetState,
    saveToLocal,
    loadFromLocal,
    createPole,
    createComm,
    createSpan,
    createSpanSide,
    createSpanComm,
    createSpanPower,
    keyForSpanSide,
    keyForSpanComm,
    getPole,
    getSpan,
    getSpanSide,
    getSpanComm,
    upsertPole,
    updatePoleField,
    updateSetting,
    updateSpanField,
    updateSpanPowerField,
    upsertComm,
    upsertSpan,
    upsertSpanSide,
    upsertSpanComm,
    removeSpanComm,
    addSpanPower,
    getConnectedSpans,
    getOtherPoleId,
    getSpanSidesForPole,
    getSpanSidesForSpan,
    getSpanCommsForPole,
    getSpanCommsForSpan,
    getSpanPowerForPole,
    getSpanPowerForSpan,
    poleHasChanges,
    ensureSpanSides,
    ensureSpanComms,
    ensureEndpointComms,
    normalizeState
  };
  global.AppStore.ENVIRONMENT_OPTIONS = ENVIRONMENT_OPTIONS;
})(window);
