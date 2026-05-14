(function (global) {
  "use strict";

  const CURRENT_VERSION = "1.2.0";
  const STORAGE_KEY = "poleCalculatorAppState.v2";

  const DEFAULT_CLEARANCE_TO_POWER = "40\"";
  const DEFAULT_COMM_CLEARANCE = "12\"";
  const DEFAULT_BOLT_CLEARANCE = "4\"";

  const emptyState = () => ({
    version: CURRENT_VERSION,
    importedFileName: "Datos demo",
    importedAt: new Date().toISOString(),
    selectedPoleId: "P72",
    selectedSpanId: "",
    autoCreateSpanComms: false,
    settings: {
      clearanceToPower: DEFAULT_CLEARANCE_TO_POWER,
      polePowerCommsClearance: "40\"",
      commClearance: DEFAULT_COMM_CLEARANCE,
      boltClearance: DEFAULT_BOLT_CLEARANCE,
      midspanPowerCommClearance: "30\"",
      midspanCommCommClearance: "4'",
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
      proposedMidspan: trim(data.proposedMidspan || ""),
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
      difference: trim(data.difference || ""),
      remotePoleId: trim(data.remotePoleId || ""),
      remoteHOA: trim(data.remoteHOA || ""),
      ocalcMS: trim(data.ocalcMS || ""),
      midspan: trim(data.midspan || ""),
      calculatedMidspan: trim(data.calculatedMidspan || ""),
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

  function addSpanPower(data) {
    const clean = createSpanPower(data);
    if (!clean.spanId || !clean.poleId) return null;
    const key = [clean.spanId, clean.poleId, clean.wireId || clean.label || Object.keys(state.spanPower).length].join("__");
    state.spanPower[key] = clean;
    return clean;
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
    return Object.values(state.spanPower).filter(row => row.poleId === poleId);
  }

  function getSpanPowerForSpan(spanId) {
    return Object.values(state.spanPower).filter(row => row.spanId === spanId);
  }

  function poleHasChanges(poleId) {
    const sideChange = getSpanSidesForPole(poleId).some(side => side.proposedHOA || side.proposedMidspan || side.endDrop || side.notes);
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
    return state;
  }

  function loadSampleData() {
    state = emptyState();
    state.importedFileName = "Datos de ejemplo";

    [
      createPole("P72", "35'", "", { lowPower: "30'8\"", sequence: "P72" }),
      createPole("P73", "40'", "", { lowPower: "22'5\"", sequence: "P73" }),
      createPole("P74", "35'", "", { lowPower: "23'", sequence: "P74" }),
      createPole("P81", "35'", "", { lowPower: "20'9\"", sequence: "P81" }),
      createPole("P79", "40'", "", { lowPower: "25'", sequence: "P79" }),
      createPole({ poleId: "Unknown-S-P79", isGenerated: true })
    ].forEach(upsertPole);

    [
      createSpan("S-P72-P73", "P72", "P73", "E", "Principal", { lengthDisplay: "125'", bearingDegrees: 90 }),
      createSpan("S-P73-P74", "P73", "P74", "E", "Principal", { lengthDisplay: "110'", bearingDegrees: 90 }),
      createSpan("S-P73-P81", "P73", "P81", "S", "Ramal", { lengthDisplay: "95'", bearingDegrees: 180 }),
      createSpan("S-P79-UNKNOWN", "P79", "Unknown-S-P79", "N", "Other pole editable", { lengthDisplay: "80'", bearingDegrees: 0, isGeneratedOtherPole: true })
    ].forEach(upsertSpan);

    [
      ["P72", "CTL", "24'"], ["P72", "CATV", "22'6\""], ["P72", "FIBER", "21'10\""],
      ["P73", "CTL", "25'"], ["P73", "CATV", "23'1\""], ["P73", "FIBER", "22'"],
      ["P74", "CTL", "24'8\""], ["P74", "CATV", "22'10\""],
      ["P81", "CATV", "23'4\""], ["P81", "FIBER", "21'8\""],
      ["P79", "CTL", "24'2\""], ["P79", "FIBER", "22'4\""],
      ["Unknown-S-P79", "CTL", "23'8\""], ["Unknown-S-P79", "FIBER", "22'1\""]
    ].forEach(([poleId, owner, hoa]) => upsertComm(poleId, owner, hoa, "", { ownerBase: owner }));

    const sampleSpanComms = [
      { spanId: "S-P72-P73", poleId: "P72", owner: "CTL", ownerBase: "CTL", existingHOA: "24'", ocalcMS: "21'8\"", midspan: "21'8\"" },
      { spanId: "S-P72-P73", poleId: "P73", owner: "CTL", ownerBase: "CTL", existingHOA: "25'", ocalcMS: "21'8\"", midspan: "21'8\"" },
      { spanId: "S-P72-P73", poleId: "P72", owner: "CATV", ownerBase: "CATV", existingHOA: "22'6\"", ocalcMS: "20'2\"", midspan: "20'2\"" },
      { spanId: "S-P72-P73", poleId: "P73", owner: "CATV", ownerBase: "CATV", existingHOA: "23'1\"", ocalcMS: "20'2\"", midspan: "20'2\"" },
      { spanId: "S-P73-P74", poleId: "P73", owner: "CTL", ownerBase: "CTL", existingHOA: "25'", ocalcMS: "21'6\"", midspan: "21'6\"" },
      { spanId: "S-P73-P74", poleId: "P74", owner: "CTL", ownerBase: "CTL", existingHOA: "24'8\"", ocalcMS: "21'6\"", midspan: "21'6\"" },
      { spanId: "S-P73-P81", poleId: "P73", owner: "CATV", ownerBase: "CATV", existingHOA: "23'1\"", ocalcMS: "20'9\"", midspan: "20'9\"" },
      { spanId: "S-P73-P81", poleId: "P81", owner: "CATV", ownerBase: "CATV", existingHOA: "23'4\"", ocalcMS: "20'9\"", midspan: "20'9\"" },
      { spanId: "S-P79-UNKNOWN", poleId: "P79", owner: "CTL", ownerBase: "CTL", existingHOA: "24'2\"", ocalcMS: "22'2\"", midspan: "22'2\"" },
      { spanId: "S-P79-UNKNOWN", poleId: "Unknown-S-P79", owner: "CTL", ownerBase: "CTL", existingHOA: "23'8\"", ocalcMS: "22'2\"", midspan: "22'2\"" }
    ];
    sampleSpanComms.forEach(data => upsertSpanComm(data));

    upsertSpanSide({ spanId: "S-P72-P73", poleId: "P72", proposedHOA: "24'1\"", proposedMidspan: "23'7\"" });
    upsertSpanSide({ spanId: "S-P72-P73", poleId: "P73", proposedHOA: "24'1\"", proposedMidspan: "23'7\"" });
    upsertSpanSide({ spanId: "S-P73-P74", poleId: "P73", proposedHOA: "19'1\"", proposedMidspan: "19'2\"" });

    return normalizeState(state);
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
    loadSampleData,
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
    upsertComm,
    upsertSpan,
    upsertSpanSide,
    upsertSpanComm,
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
    normalizeState
  };
})(window);
