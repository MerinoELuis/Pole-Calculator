(function (global) {
  "use strict";

  // AppStore is the single source of truth for the calculator. UI modules read
  // from this state, and calculation modules write derived values back into it.
  const CURRENT_VERSION = "1.4.1";
  const STORAGE_KEY = "poleCalculatorAppState.v2";

  const DEFAULT_CLEARANCE_TO_POWER = "40\"";
  const DEFAULT_COMM_CLEARANCE = "12\"";
  const DEFAULT_BOLT_CLEARANCE = "4\"";

  /**
   * @typedef {Object} Pole
   * @property {string} poleId Stable visible pole identifier.
   * @property {string} lowPower Editable lowest power attachment on the pole.
   * @property {string} maxCommHeight Derived maximum communication height.
   * @property {string} standaloneProposedHOA Proposed height for a terminal pole.
   */

  /**
   * @typedef {Object} Span
   * @property {string} spanId Stable imported or generated edge identifier.
   * @property {string} fromPole Directed source pole ID.
   * @property {string} toPole Directed destination pole ID.
   * @property {string} type Imported Fore Span, Back Span, or Other relation.
   * @property {string} sourceSpanId Physical source used by an additional Proposed row.
   */

  /**
   * @typedef {Object} SpanSide
   * @property {string} spanId Related span identifier.
   * @property {string} poleId Pole that owns this Proposed row.
   * @property {string} proposedHOA Editable Proposed attachment height.
   * @property {string} proposedHOAChange Editable or automatic Next Pole Proposed.
   * @property {string} finalMidspan Derived Adjusted Final MS.
   */

  /**
   * @typedef {Object} SpanComm
   * @property {string} spanId Related span identifier.
   * @property {string} poleId Pole endpoint that owns this row.
   * @property {string} owner Imported communication owner.
   * @property {string} wireId Imported wire identity used for endpoint matching.
   * @property {string} existingHOA Imported baseline attachment height.
   * @property {string} existingHOAChange Editable or automatic new attachment height.
   * @property {string} midspan Imported baseline midspan for this exact row.
   * @property {string} calculatedMidspan Derived midspan after endpoint movements.
   * @property {boolean} isEndpointPlaceholder True only for a synthetic owner row copied to an otherwise empty endpoint.
   */

  /**
   * @typedef {Object} AppState
   * @property {string} version Saved-state format version.
   * @property {Object.<string, Pole>} poles Poles keyed by poleId.
   * @property {Object.<string, Span>} spans Spans keyed by spanId.
   * @property {Object.<string, SpanSide>} spanSides Proposed rows keyed by span and pole.
   * @property {Object.<string, SpanComm>} spanComms Comm rows keyed by span, pole, owner, and wire.
   * @property {Object} settings Project profile and editable calculation settings.
   */
  const ENVIRONMENT_OPTIONS = [
    { value: "NONE", label: "None", clearance: "15'6\"" },
    { value: "STREET", label: "Street", clearance: "15'6\"" },
    { value: "HIGHWAY", label: "Highway", clearance: "18'" },
    { value: "PEDESTRIAN", label: "Pedestrian", clearance: "9'6\"" },
    { value: "PARALLEL_TO_STREET", label: "Parallel to street", clearance: "15'6\"" },
    { value: "OBSTRUCTED_PARALLEL_TO_STREET", label: "Obstructed parallel to street", clearance: "15'6\"" },
    { value: "UNLIKELY_PARALLEL_TO_STREET", label: "Unlikely parallel to street", clearance: "15'6\"" },
    { value: "RESIDENTIAL_DRIVEWAY", label: "Residential driveway", clearance: "15'6\"" },
    { value: "COMMERCIAL_DRIVEWAY", label: "Commercial driveway", clearance: "15'6\"" },
    { value: "PARKING_LOT", label: "Parking lot", clearance: "15'6\"" },
    { value: "ALLEY", label: "Alley", clearance: "15'6\"" },
    { value: "RAILROAD", label: "Railroad", clearance: "25'" },
    { value: "RURAL", label: "Rural", clearance: "9'6\"" },
    { value: "FARM", label: "Farm", clearance: "15'6\"" },
    { value: "WATER_WITH_SAILBOATS", label: "Water with sailboats", clearance: "Variable" },
    { value: "WATER_WITHOUT_SAILBOATS", label: "Water without sailboats", clearance: "15'" },
    { value: "TROLLEY", label: "Trolley", clearance: "25'" }
  ];

  // Keep startup empty. Imported Excel/JSON or local saves are the only things
  // that should populate the workspace in normal use.
  /** @returns {AppState} A new empty state with current defaults. */
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
      projectProfile: "INTEC",
      position: "TOP_COMM",
      mrCase: "LOWER",
      proposedOwner: "Wecom",
      calculateBackspanMidspan: false,
      attachmentMessengerSize: "",
      fiberSizes: {},
      borrowMidspanFromPhysicalSpan: false,
      proposeForeSpanWithoutMidspan: false,
      allowLowPowerMidspanAdjustment: true,
      showServiceDrop: true,
      showResagServiceDrop: true,
      hideProposedOwner: false,
      environmentClearances: {},
      streetlightBracketCommClearance: "",
      streetlightDripLoopCommClearance: "",
      powerGuyCommClearance: "",
      streetlightGroundingRequired: false,
      environmentClearance: "15'6\"",
      sagPer100Ft: "1'"
    },
    poles: {},
    spans: {},
    spanSides: {},
    spanComms: {},
    spanPower: {},
    makeReadyReferences: [],
    // Raw review source is preserved separately from normalized calculator
    // entities. Empty IDs and duplicate source rows are meaningful to Excel
    // Review even though they cannot be graph keys.
    excelReviewSource: {
      collection: { headers: [], rows: [] },
      spans: { headers: [], rows: [] },
      spanWires: { headers: [], rows: [] },
      equipment: { headers: [], rows: [] },
      anchors: { headers: [], rows: [] },
      anchorGuys: { headers: [], rows: [] },
      makeReady: { headers: [], rows: [] },
      commTransfers: { headers: [], rows: [] }
    },
    // Keys are stable identities for review findings the user has explicitly
    // accepted for this job. Results are still regenerated from the workbook.
    excelReviewIgnoredChecks: {},
    poleClassChecks: [],
    movements: [],
    mr: [],
    warnings: [],
    // Canonical identities intentionally removed by the user. Update Data uses
    // this list to avoid restoring the same pole under a slightly different
    // display name such as "P01" versus "P01 STEEL".
    deletedPoleIds: [],
    ui: {
      search: "",
      filter: "all",
      // Per-pole presentation preference. Collapsing the comm table never
      // removes engineering data and is preserved in the saved job JSON.
      hiddenCommPoleIds: [],
      hiddenPoleIds: []
    }
  });

  let state = emptyState();

  function trim(value) {
    return String(value ?? "").trim();
  }

  // STEEL describes construction and UG/PCO describe engineering state; none
  // of them changes the physical pole identity. Collection keeps the preferred
  // display name while import/update matching uses this stable comparison key.
  function canonicalPoleIdentity(value) {
    const parts = trim(value).replace(/\s+/g, " ").split(" ").filter(Boolean);
    while (parts.length > 1 && /^(STEEL|UG|PCO)$/i.test(parts[parts.length - 1])) parts.pop();
    return parts.join(" ").toUpperCase();
  }

  function defaultEnvironmentClearance(environment) {
    const env = trim(environment || "NONE") || "NONE";
    const profileValue = state.settings?.environmentClearances?.[env];
    if (profileValue) return trim(profileValue);
    const option = ENVIRONMENT_OPTIONS.find(item => item.value === env);
    return option ? option.clearance : "";
  }

  function keyForSpanSide(spanId, poleId) {
    return [spanId, poleId].map(v => trim(v)).join("__");
  }

  function keyForSpanComm(spanId, poleId, owner, wireId) {
    return [spanId, poleId, owner, wireId || ""].map(v => trim(v)).join("__");
  }

  /**
   * Normalizes a pole from an imported object or individual values.
   * @param {string|Object} poleId Pole ID or pole-shaped source object.
   * @param {string} [poleHeight]
   * @param {string} [notes]
   * @param {Object} [extra]
   * @returns {Pole}
   */
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
        standaloneProposedHOA: trim(data.standaloneProposedHOA || ""),
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
      standaloneProposedHOA: trim(extra.standaloneProposedHOA || ""),
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

  /**
   * Creates a normalized graph edge and a stable Unknown endpoint when needed.
   * @param {string} spanId
   * @param {string} fromPole
   * @param {string} [toPole]
   * @param {string} [direction]
   * @param {string} [notes]
   * @param {Object} [extra]
   * @returns {Span}
   */
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
      environmentClearance: trim(extra.environmentClearance || defaultEnvironmentClearance(extra.environment || "NONE")),
      midspanLowPower: trim(extra.midspanLowPower || ""),
      midspanMaxCommHeight: trim(extra.midspanMaxCommHeight || ""),
      rawSpanIds: Array.isArray(extra.rawSpanIds) ? extra.rawSpanIds : (extra.rawSpanId ? [extra.rawSpanId] : []),
      rawType: trim(extra.rawType || ""),
      linkedCollectionId: trim(extra.linkedCollectionId || ""),
      linkedCollectionTitle: trim(extra.linkedCollectionTitle || ""),
      sourceCollectionId: trim(extra.sourceCollectionId || ""),
      sourceSpanId: trim(extra.sourceSpanId || ""),
      isManualProposed: Boolean(extra.isManualProposed),
      isGeneratedOtherPole: Boolean(extra.isGeneratedOtherPole)
    };
  }

  /**
   * @param {Object} [data]
   * @returns {SpanSide}
   */
  function createSpanSide(data = {}) {
    return {
      spanId: trim(data.spanId || ""),
      poleId: trim(data.poleId || ""),
      isManualProposed: Boolean(data.isManualProposed),
      isAdditionalProposed: Boolean(data.isAdditionalProposed),
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

  /**
   * @param {Object} [data]
   * @returns {SpanComm}
   */
  function createSpanComm(data = {}) {
    return {
      spanId: trim(data.spanId || ""),
      poleId: trim(data.poleId || ""),
      owner: trim(data.owner || ""),
      ownerBase: trim(data.ownerBase || data.owner || ""),
      existingHOA: trim(data.existingHOA || ""),
      existingHOAChange: trim(data.existingHOAChange || ""),
      serviceDrop: Boolean(data.serviceDrop),
      downGuy: Boolean(data.downGuy),
      transferToNewPole: Boolean(data.transferToNewPole),
      resagServiceDrop: Boolean(data.resagServiceDrop),
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
      autoCalcStatus: trim(data.autoCalcStatus || ""),
      autoCalcMessage: trim(data.autoCalcMessage || ""),
      mr: trim(data.mr || ""),
      notes: trim(data.notes || ""),
      rawOwner: trim(data.rawOwner || ""),
      unknownOwner: Boolean(data.unknownOwner),
      size: trim(data.size || ""),
      construction: trim(data.construction || ""),
      insulator: trim(data.insulator || ""),
      wireId: trim(data.wireId || ""),
      wireIndex: trim(data.wireIndex || ""),
      isEndpointPlaceholder: Boolean(data.isEndpointPlaceholder),
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

  function createMakeReadyReference(data = {}) {
    return {
      poleId: trim(data.poleId || data.id || ""),
      collectionId: trim(data.collectionId || ""),
      makeReadyIndex: trim(data.makeReadyIndex || data.index || ""),
      makeReadyId: trim(data.makeReadyId || ""),
      attachmentSizeRaw: trim(data.attachmentSizeRaw || data.attachmentSize || ""),
      attachmentMessenger: trim(data.attachmentMessenger || ""),
      attachmentFiber: trim(data.attachmentFiber || ""),
      attachmentDirection: trim(data.attachmentDirection || ""),
      attachmentDirectionTokens: Array.isArray(data.attachmentDirectionTokens) ? data.attachmentDirectionTokens.map(trim).filter(Boolean) : [],
      attachmentType: trim(data.attachmentType || ""),
      attachmentHeight: trim(data.attachmentHeight || ""),
      proposedMidspan: trim(data.proposedMidspan || ""),
      makeReadyNotes: trim(data.makeReadyNotes || data.notes || ""),
      commTransfers: trim(data.commTransfers || ""),
      raw: data.raw || {}
    };
  }

  function normalizeReviewSheet(sheet) {
    const source = sheet && typeof sheet === "object" ? sheet : {};
    return {
      headers: Array.isArray(source.headers) ? source.headers.map(trim) : [],
      rows: Array.isArray(source.rows)
        ? source.rows.map(row => (row && typeof row === "object" ? { ...row } : {}))
        : []
    };
  }

  function normalizeExcelReviewSource(source) {
    const raw = source && typeof source === "object" ? source : {};
    return {
      collection: normalizeReviewSheet(raw.collection),
      spans: normalizeReviewSheet(raw.spans),
      spanWires: normalizeReviewSheet(raw.spanWires),
      equipment: normalizeReviewSheet(raw.equipment),
      anchors: normalizeReviewSheet(raw.anchors),
      anchorGuys: normalizeReviewSheet(raw.anchorGuys),
      makeReady: normalizeReviewSheet(raw.makeReady),
      commTransfers: normalizeReviewSheet(raw.commTransfers)
    };
  }

  /**
   * @param {Object} nextState
   * @returns {AppState} Normalized current state.
   */
  function setState(nextState) {
    state = normalizeState(nextState || emptyState());
    return state;
  }

  /** @returns {AppState} Mutable application state owned by AppStore. */
  function getState() { return state; }

  /** @returns {AppState} A new empty state. */
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
    if (!["poleHeight", "lowPower", "maxCommHeight", "topComm", "lowComm", "standaloneProposedHOA", "notes", "sequence"].includes(field)) return pole;
    pole[field] = trim(value);
    return pole;
  }

  function updateSetting(field, value) {
    if (!state.settings || !Object.prototype.hasOwnProperty.call(emptyState().settings, field)) return null;
    if (field === "projectProfile" && global.ProjectProfiles) {
      state.settings = global.ProjectProfiles.applyProfileSettings(state.settings, value);
      Object.values(state.spans || {}).forEach(span => {
        span.environmentClearance = defaultEnvironmentClearance(span.environment || "NONE");
      });
      return state.settings;
    }
    state.settings[field] = trim(value);
    if (field === "polePowerCommsClearance") state.settings.clearanceToPower = trim(value);
    if (field === "clearanceToPower") state.settings.polePowerCommsClearance = trim(value);
    return state.settings;
  }

  function applyProjectProfile(profileId) {
    if (!global.ProjectProfiles) return state.settings;
    state.settings = global.ProjectProfiles.applyProfileSettings(state.settings || emptyState().settings, profileId);
    return state.settings;
  }

  function updateSpanField(spanId, field, value) {
    const span = state.spans[spanId];
    if (!span) return null;
    if (!["environment", "environmentClearance", "midspanLowPower", "midspanMaxCommHeight", "notes"].includes(field)) return span;
    span[field] = trim(value);
    if (field === "environment" && !span.environmentClearance) {
      span.environmentClearance = defaultEnvironmentClearance(value);
    }
    return span;
  }

  function upsertComm(poleId, owner, existingHOA = "", notes = "", extra = {}) {
    const pole = state.poles[poleId] || upsertPole(createPole(poleId));
    // A manual/imported comm means the pole is no longer intentionally empty.
    // Clear the guard used after a user deletes every comm from this pole.
    pole.metadata = { ...(pole.metadata || {}), suppressEndpointComms: false };
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

  function removeSpanSide(spanId, poleId) {
    delete state.spanSides[keyForSpanSide(spanId, poleId)];
  }

  function removeManualSpan(spanId) {
    const span = state.spans[trim(spanId)];
    if (!span || (!span.isManualProposed && !/^(manual|additional)-proposed-/i.test(trim(spanId)))) return false;
    delete state.spans[trim(spanId)];
    Object.keys(state.spanSides).forEach(key => {
      if (state.spanSides[key].spanId === spanId) delete state.spanSides[key];
    });
    return true;
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

  /**
   * Removes every comm relationship owned by one pole without touching its
   * spans, Proposed rows, power wires, equipment, or the opposite endpoint.
   * The suppression marker prevents normalizeState() from recreating empty
   * endpoint reference rows after the job is saved and loaded again.
   * @param {string} poleId
   * @returns {Array<SpanComm>} Removed rows, used by callers to refresh spans.
   */
  function removeSpanCommsForPole(poleId) {
    const id = trim(poleId);
    const rows = getSpanCommsForPole(id).slice();
    rows.forEach(row => removeSpanComm(row.spanId, row.poleId, row.owner, row.wireId || ""));
    const pole = state.poles[id];
    if (pole) {
      pole.comms = [];
      pole.metadata = { ...(pole.metadata || {}), suppressEndpointComms: true };
    }
    return rows;
  }

  /**
   * Deletes one pole and every graph entity that depends on it. The canonical
   * identity is retained as a tombstone so Update Data cannot recreate a pole
   * that the user deliberately removed from this job.
   * @param {string} poleId
   * @param {{remember?: boolean}} [options]
   * @returns {{poleId:string, spanIds:string[], commCount:number}}
   */
  function removePole(poleId, options = {}) {
    const id = trim(poleId);
    if (!state.poles[id]) return { poleId: id, spanIds: [], commCount: 0 };

    const spanIds = new Set(getConnectedSpans(id).map(span => span.spanId));
    // Additional Proposed rows inherit a physical source span and must leave
    // with it even if their generated endpoints do not directly match id.
    let foundDependent = true;
    while (foundDependent) {
      foundDependent = false;
      Object.values(state.spans).forEach(span => {
        if (spanIds.has(span.spanId) || !span.sourceSpanId || !spanIds.has(span.sourceSpanId)) return;
        spanIds.add(span.spanId);
        foundDependent = true;
      });
    }

    const commRows = Object.values(state.spanComms).filter(row => row.poleId === id || spanIds.has(row.spanId));
    commRows.forEach(row => removeSpanComm(row.spanId, row.poleId, row.owner, row.wireId || ""));
    Object.keys(state.spanSides).forEach(key => {
      const row = state.spanSides[key];
      if (row.poleId === id || spanIds.has(row.spanId)) delete state.spanSides[key];
    });
    Object.keys(state.spanPower).forEach(key => {
      const row = state.spanPower[key];
      if (row.poleId === id || spanIds.has(row.spanId)) delete state.spanPower[key];
    });
    spanIds.forEach(spanId => delete state.spans[spanId]);
    delete state.poles[id];

    state.makeReadyReferences = state.makeReadyReferences.filter(row => row.poleId !== id);
    state.poleClassChecks = state.poleClassChecks.filter(row => row.poleId !== id);
    state.movements = state.movements.filter(row => row.poleId !== id);
    state.mr = state.mr.filter(row => row.poleId !== id);
    state.warnings = state.warnings.filter(row => row.poleId !== id && !spanIds.has(row.spanId));
    if (state.selectedPoleId === id) state.selectedPoleId = "";

    state.ui.hiddenPoleIds = (state.ui.hiddenPoleIds || []).filter(value => value !== id);
    state.ui.hiddenCommPoleIds = (state.ui.hiddenCommPoleIds || []).filter(value => value !== id);
    if (options.remember !== false) {
      const canonical = canonicalPoleIdentity(id);
      state.deletedPoleIds = Array.from(new Set([...(state.deletedPoleIds || []), canonical].filter(Boolean)));
    }
    return { poleId: id, spanIds: Array.from(spanIds), commCount: commRows.length };
  }

  /**
   * Stores the user action attached to one normalized Power Equipment row.
   * The imported physical heights remain unchanged; actionHeight is the
   * proposed redress/raise height used by calculations and Make Ready.
   */
  function updatePowerEquipmentField(poleId, equipmentIndex, field, value) {
    if (!["actionActive", "actionHeight"].includes(field)) return null;
    const pole = state.poles[poleId];
    const rows = pole?.metadata?.powerEquipment;
    const index = Number(equipmentIndex);
    if (!pole || !Array.isArray(rows) || !Number.isInteger(index) || !rows[index]) return null;
    const nextRows = rows.map((row, rowIndex) => rowIndex === index
      ? {
          ...row,
          [field]: field === "actionActive" ? Boolean(value) : trim(value)
        }
      : row);
    pole.metadata = { ...(pole.metadata || {}), powerEquipment: nextRows };
    return nextRows[index];
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
    const equipmentChange = (state.poles[poleId]?.metadata?.powerEquipment || [])
      .some(row => Boolean(row.actionActive || trim(row.actionHeight || "")));
    return Boolean(state.poles[poleId]?.standaloneProposedHOA) || sideChange || commChange || equipmentChange;
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

  function isSyntheticProposedSpan(span) {
    return Boolean(span?.sourceSpanId || /^(manual|additional)-proposed-/i.test(trim(span?.spanId)));
  }

  function proposedSideHasSlot(side) {
    return Boolean(side && (
      side.isManualProposed || side.proposedHOA || side.proposedHOAChange ||
      side.ocalcMS || side.proposedMidspan || side.notes
    ));
  }

  function physicalSpanCandidateFor(synthetic) {
    return Object.values(state.spans)
      .filter(candidate => candidate.spanId !== synthetic.spanId)
      .filter(candidate => !isSyntheticProposedSpan(candidate))
      .filter(candidate => candidate.fromPole === synthetic.fromPole && candidate.toPole === synthetic.toPole)
      .sort((a, b) => {
        const score = span =>
          getSpanPowerForSpan(span.spanId).length * 100 +
          getSpanCommsForSpan(span.spanId).length * 20 +
          (span.lengthDisplay ? 5 : 0) +
          (span.bearingDegrees !== "" && span.bearingDegrees !== null ? 2 : 0);
        return score(b) - score(a);
      })[0] || null;
  }

  function inheritPhysicalSpanData(span, source) {
    return {
      ...span,
      sourceSpanId: source.spanId,
      direction: source.direction || span.direction || "",
      bearingDegrees: source.bearingDegrees ?? span.bearingDegrees ?? "",
      type: source.type || span.type || "Other",
      rawType: source.rawType || span.rawType || "",
      spanIndex: source.spanIndex || span.spanIndex || "",
      length: source.length || span.length || "",
      lengthDisplay: source.lengthDisplay || span.lengthDisplay || "",
      environment: source.environment || span.environment || "NONE",
      environmentClearance: source.environmentClearance || span.environmentClearance || "",
      midspanLowPower: source.midspanLowPower || span.midspanLowPower || "",
      midspanMaxCommHeight: source.midspanMaxCommHeight || span.midspanMaxCommHeight || "",
      rawSpanIds: Array.from(new Set([...(source.rawSpanIds || []), ...(span.rawSpanIds || [])].filter(Boolean)))
    };
  }

  function reconcileSyntheticProposedSpans() {
    // Older versions created an additional span even for the first Proposed on
    // an imported connection. Move that first proposal to the physical span;
    // genuine later proposals keep a sourceSpanId to inherit physical data.
    Object.values(state.spans)
      .filter(isSyntheticProposedSpan)
      .sort((a, b) => trim(a.spanId).localeCompare(trim(b.spanId), undefined, { numeric: true }))
      .forEach(synthetic => {
        const source = synthetic.sourceSpanId ? state.spans[synthetic.sourceSpanId] : physicalSpanCandidateFor(synthetic);
        if (!source || source.spanId === synthetic.spanId) return;
        const syntheticSide = state.spanSides[keyForSpanSide(synthetic.spanId, synthetic.fromPole)];
        const sourceSide = state.spanSides[keyForSpanSide(source.spanId, synthetic.fromPole)] || createSpanSide({
          spanId: source.spanId,
          poleId: synthetic.fromPole
        });

        if (proposedSideHasSlot(syntheticSide) && !proposedSideHasSlot(sourceSide)) {
          state.spanSides[keyForSpanSide(source.spanId, synthetic.fromPole)] = createSpanSide({
            ...sourceSide,
            ...syntheticSide,
            spanId: source.spanId,
            poleId: synthetic.fromPole,
            isManualProposed: true,
            isAdditionalProposed: false
          });
          delete state.spans[synthetic.spanId];
          Object.keys(state.spanSides).forEach(key => {
            if (state.spanSides[key].spanId === synthetic.spanId) delete state.spanSides[key];
          });
          return;
        }

        state.spans[synthetic.spanId] = inheritPhysicalSpanData(synthetic, source);
      });
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
    // If one end of the span has no own rows, create placeholder comms with the
    // same owners/wireIds arriving from the other end. The user can complete the
    // heights later, and the pole never appears visually empty.
    Object.values(state.spans).forEach(span => {
      [span.fromPole, span.toPole].filter(Boolean).forEach(poleId => {
        if (state.poles[poleId]?.metadata?.suppressEndpointComms) return;
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
              downGuy: Boolean(row.downGuy),
              transferToNewPole: false,
              resagServiceDrop: false,
              rawOwner: row.rawOwner || row.owner,
              size: row.size || "",
              wireId: row.wireId || "",
              // This is not a Span.Wire measurement. It only keeps the arriving
              // owner visible until that endpoint receives its own field data.
              isEndpointPlaceholder: true
            });
          });
      });
    });
  }

  function normalizeState(raw) {
    const next = { ...emptyState(), ...raw };
    const rawSettings = raw && raw.settings ? raw.settings : {};
    const profileDefaults = global.ProjectProfiles
      ? (global.ProjectProfiles.getProfile(rawSettings.projectProfile || "INTEC")?.settings || {})
      : {};
    // Apply project defaults before saved values. Existing user edits win,
    // while older JSON files automatically receive newly introduced profile
    // rules such as MidAm streetlight and Back Span behavior.
    next.settings = { ...emptyState().settings, ...profileDefaults, ...rawSettings };
    // Migrate Metronet saves created before the WI selector existed.
    if (String(next.settings.projectProfile || "").toUpperCase() === "METRONET"
      && (!next.settings.proposedOwner || String(next.settings.proposedOwner).toUpperCase() === "METRONET")) {
      next.settings.proposedOwner = "MidAm";
    }
    next.settings.fiberSizes = next.settings.fiberSizes && typeof next.settings.fiberSizes === "object"
      ? { ...next.settings.fiberSizes }
      : {};
    next.poles = next.poles || {};
    next.spans = next.spans || {};
    next.spanSides = next.spanSides || {};
    next.spanComms = next.spanComms || {};
    next.spanPower = next.spanPower || {};
    next.makeReadyReferences = Array.isArray(next.makeReadyReferences)
      ? next.makeReadyReferences.map(createMakeReadyReference)
      : [];
    next.excelReviewSource = normalizeExcelReviewSource(next.excelReviewSource);
    next.excelReviewIgnoredChecks = next.excelReviewIgnoredChecks && typeof next.excelReviewIgnoredChecks === "object"
      ? { ...next.excelReviewIgnoredChecks }
      : {};
    next.poleClassChecks = Array.isArray(next.poleClassChecks) ? next.poleClassChecks : [];
    next.movements = Array.isArray(next.movements) ? next.movements : [];
    next.mr = Array.isArray(next.mr) ? next.mr : [];
    next.warnings = Array.isArray(next.warnings) ? next.warnings : [];
    next.deletedPoleIds = Array.isArray(next.deletedPoleIds)
      ? Array.from(new Set(next.deletedPoleIds.map(canonicalPoleIdentity).filter(Boolean)))
      : [];
    next.ui = { search: "", filter: "all", hiddenCommPoleIds: [], hiddenPoleIds: [], ...(next.ui || {}) };
    next.ui.hiddenCommPoleIds = Array.isArray(next.ui.hiddenCommPoleIds)
      ? Array.from(new Set(next.ui.hiddenCommPoleIds.map(trim).filter(Boolean)))
      : [];
    next.ui.hiddenPoleIds = Array.isArray(next.ui.hiddenPoleIds)
      ? Array.from(new Set(next.ui.hiddenPoleIds.map(trim).filter(id => Boolean(id && next.poles[id]))))
      : [];

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
    reconcileSyntheticProposedSpans();
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

  /**
   * Public state API. See docs/DATA_MODEL.md for entity ownership and keys.
   * @namespace AppStore
   */
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
    createMakeReadyReference,
    canonicalPoleIdentity,
    defaultEnvironmentClearance,
    keyForSpanSide,
    keyForSpanComm,
    getPole,
    getSpan,
    getSpanSide,
    getSpanComm,
    upsertPole,
    updatePoleField,
    updateSetting,
    applyProjectProfile,
    updateSpanField,
    updateSpanPowerField,
    updatePowerEquipmentField,
    upsertComm,
    upsertSpan,
    upsertSpanSide,
    removeSpanSide,
    removeManualSpan,
    upsertSpanComm,
    removeSpanComm,
    removeSpanCommsForPole,
    removePole,
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
