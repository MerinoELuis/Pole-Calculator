(function (global) {
  "use strict";

  // AppStore is the single source of truth for the calculator. UI modules read
  // from this state, and calculation modules write derived values back into it.
  const CURRENT_VERSION = "1.8.223";
  const STORAGE_KEY = "poleCalculatorAppState.v2";

  const DEFAULT_CLEARANCE_TO_POWER = "40\"";
  const DEFAULT_COMM_CLEARANCE = "12\"";
  const DEFAULT_BOLT_CLEARANCE = "4\"";
  const Identity = global.AppStoreIdentity;

  function normalizePoleInsetReason(value) {
    return String(value || "").trim().toUpperCase() === "OVERLOADED"
      ? "OVERLOADED"
      : "FAILING_CLEARANCES";
  }

  // Convert imported/save filenames into the editable job label. Keep this
  // logic in AppStore so Save, AutoProposed and import all use one source.
  function jobNameFromFileName(value) {
    const clean = String(value || "")
      .replace(/\.(?:json|xlsx?|xlsm|csv)$/i, "")
      .replace(/_(?:Pole_Calculator|AutoProposed|Debug)$/i, "")
      .replace(/^excel[_\-\s]*/i, "")
      .replace(/[_\-\s]*20\d{2}[-_]\d{2}[-_]\d{2}(?:\s*\(\d+\))?$/i, "")
      .trim();
    return clean || "Pole Job";
  }

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
   * @property {string} otherHOA Optional per-span override height used for comm movement calculations.
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
    jobName: "",
    importedFileName: "",
    importedAt: new Date().toISOString(),
    selectedPoleId: "",
    selectedSpanId: "",
    autoCreateSpanComms: false,
    settings: {
      clearanceToPower: DEFAULT_CLEARANCE_TO_POWER,
      polePowerCommsClearance: "40\"",
      primaryPowerCommsClearance: "43\"",
      commClearance: DEFAULT_COMM_CLEARANCE,
      boltClearance: DEFAULT_BOLT_CLEARANCE,
      midspanPowerCommClearance: "30\"",
      midspanPrimaryPowerCommClearance: "33\"",
      midspanCommCommClearance: "4\"",
      projectProfile: "INTEC",
      metronetWI: "MIDAM",
      position: "TOP_COMM",
      mrCase: "LOWER",
      proposedOwner: "Wecom",
      calculateBackspanMidspan: false,
      attachmentMessengerSize: "",
      fiberSizes: {},
      borrowMidspanFromPhysicalSpan: false,
      proposeForeSpanWithoutMidspan: false,
      allowLowPowerMidspanAdjustment: true,
      autoMoveCommsToProposed: false,
      showServiceDrop: true,
      showResagServiceDrop: true,
      hideProposedOwner: false,
      environmentClearances: {},
      streetlightBracketCommClearance: "",
      streetlightGroundedBracketCommClearance: "",
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
    // Operator-edited IDs are matched back to imported workbook IDs on Update Data.
    poleIdAliases: {},
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
      hiddenPoleIds: [],
      // Unknown poles begin hidden, but selecting one from the index records a
      // user exception so Save/Load does not hide it again unexpectedly.
      revealedUnknownPoleIds: []
    }
  });

  let state = emptyState();

  function trim(value) {
    return String(value ?? "").trim();
  }

  // Midspan fields are heights, not free-form labels. Older imports could
  // accidentally store a pole Id here when the source workbook had no
  // Midspan column. Drop those invalid values during every state normalization
  // so a loaded debug/export file cannot display fabricated Midspan data.
  function normalizedHeightOrEmpty(value) {
    const raw = trim(value || "");
    if (!raw) return "";
    return global.HeightUtils?.parseHeight(raw) === null ? "" : raw;
  }

  function firstNormalizedHeight(...values) {
    for (const value of values) {
      const normalized = normalizedHeightOrEmpty(value);
      if (normalized) return normalized;
    }
    return "";
  }

  function sourceRowValue(row, candidates) {
    if (!row || typeof row !== "object") return "";
    for (const candidate of candidates) {
      if (Object.prototype.hasOwnProperty.call(row, candidate) && trim(row[candidate])) {
        return trim(row[candidate]);
      }
    }
    return "";
  }

  function sourceHeight(row, candidates) {
    const raw = sourceRowValue(row, candidates);
    if (!raw) return "";
    const inches = global.HeightUtils?.parseHeight(raw);
    return inches === null || inches === undefined
      ? ""
      : (global.HeightUtils?.formatHeight ? global.HeightUtils.formatHeight(inches) : raw);
  }

  function isImportedCommunicationRow(row) {
    const owner = sourceRowValue(row, ["Owner", "owner"]);
    const size = sourceRowValue(row, ["Size", "Size.display", "Wire Size"]);
    const text = `${owner} ${size}`.toLowerCase();
    if (/^utility\s*>/i.test(owner)) return false;
    if (/primary|neutral|secondary|transformer|service\s+drop|street\s*light/i.test(size)) return false;
    return /communication|catv|telco|fiber|cable\s*one|century\s*link|centurylink|cox|vexus|wecom/i.test(text);
  }

  // Keep imported communication identities consistent with excel-import.js.
  // The repair pass reads the original Span.Wire snapshot, so it must not
  // reinsert the category prefix or generic Communications suffix.
  function normalizeImportedCommOwner(value) {
    return String(value || "")
      .replace(/^COMMUNICATION\s*>\s*/i, "")
      .replace(/\s+communications?$/i, "")
      .trim();
  }

  function importedServiceDrop(size) {
    return /communication\s*drops?|service\s*drop|\bdrop\b/i.test(String(size || ""));
  }

  // Older saved states can contain the Collection comm rows and the raw
  // Span.Wire rows while a matching spanComms row was lost during an update.
  // Reconcile only rows that still exist on the pole (so an intentionally
  // deleted cable is not recreated), and always derive Service Drop from the
  // imported wire size rather than from stale saved state.
  function repairImportedSpanComms(next) {
    const sourceRows = next?.excelReviewSource?.spanWires?.rows;
    if (!Array.isArray(sourceRows) || !sourceRows.length) return;

    const poleByCanonical = new Map(Object.keys(next.poles || {}).map(id => [canonicalPoleIdentity(id), id]));
    sourceRows.forEach(row => {
      if (!isImportedCommunicationRow(row)) return;
      const spanId = sourceRowValue(row, ["Span Id", "Span ID", "spanId", "Wire Span ID"]);
      const sourcePoleId = sourceRowValue(row, ["Id", "Pole ID", "Pole", "From Pole"]);
      const owner = normalizeImportedCommOwner(sourceRowValue(row, ["Owner", "owner"]));
      const wireId = sourceRowValue(row, ["Wire Id", "Wire ID", "wireId"]);
      if (!spanId || !sourcePoleId || !owner || !wireId || !next.spans?.[spanId]) return;

      const poleId = next.poles?.[sourcePoleId]
        ? sourcePoleId
        : poleByCanonical.get(canonicalPoleIdentity(sourcePoleId)) || "";
      const pole = next.poles?.[poleId];
      if (!pole) return;
      const poleComm = (pole.comms || []).find(comm => comm.wireId === wireId);
      if (!poleComm) return;

      const key = keyForSpanComm(spanId, poleId, owner, wireId);
      const existing = next.spanComms[key];
      const size = sourceRowValue(row, ["Size", "Size.display", "Wire Size"]);
      const attachmentHeight = sourceHeight(row, [
        "Attachment Height.display",
        "Attachment Height Display",
        "Existing HOA",
        "HOA",
        "Attachment Height"
      ]) || poleComm.existingHOA || "";
      const midspan = sourceHeight(row, [
        "Mid Span Height.display",
        "Midspan.display",
        "Mid Span Height",
        "Midspan"
      ]);
      const span = next.spans[spanId];
      const remotePoleId = span.fromPole === poleId ? span.toPole : span.fromPole;
      const values = {
        spanId,
        poleId,
        owner,
        ownerBase: owner,
        existingHOA: attachmentHeight,
        serviceDrop: importedServiceDrop(size),
        remotePoleId: remotePoleId || "",
        ocalcMS: midspan,
        midspan,
        rawOwner: owner,
        size,
        construction: sourceRowValue(row, ["Construction"]),
        insulator: sourceRowValue(row, ["Insulator"]),
        wireId,
        wireIndex: sourceRowValue(row, ["Wire Index"]),
        isEndpointPlaceholder: false
      };

      if (existing) {
        // Imported facts are authoritative for this classification. Preserve
        // user movement fields while correcting stale source metadata.
        existing.serviceDrop = values.serviceDrop;
        existing.size = values.size || existing.size;
        existing.rawOwner = values.rawOwner || existing.rawOwner;
        existing.ownerBase = values.ownerBase || existing.ownerBase;
        if (!existing.existingHOA && values.existingHOA) existing.existingHOA = values.existingHOA;
        if (!existing.midspan && values.midspan) {
          existing.midspan = values.midspan;
          existing.ocalcMS = values.ocalcMS;
        }
        return;
      }

      next.spanComms[key] = createSpanComm(values);
    });
  }

  // STEEL describes construction and UG/PCO describe engineering state; none
  // of them changes the physical pole identity. Collection keeps the preferred
  // display name while import/update matching uses this stable comparison key.
  function canonicalPoleIdentity(value) {
    return Identity.canonicalPoleIdentity(value);
  }

  function isUnknownPoleId(value) {
    return Identity.isUnknownPoleId(value);
  }

  function isAddressableUnknownPoleId(value) {
    return Identity.isAddressableUnknownPoleId(value);
  }

  function defaultEnvironmentClearance(environment) {
    const env = trim(environment || "NONE") || "NONE";
    const profileValue = state.settings?.environmentClearances?.[env];
    if (profileValue) return trim(profileValue);
    const option = ENVIRONMENT_OPTIONS.find(item => item.value === env);
    return option ? option.clearance : "";
  }

  function keyForSpanSide(spanId, poleId) {
    return Identity.keyForSpanSide(spanId, poleId);
  }

  function keyForSpanComm(spanId, poleId, owner, wireId) {
    return Identity.keyForSpanComm(spanId, poleId, owner, wireId);
  }

  /**
   * Normalizes a pole from an imported object or individual values.
   * @param {string|Object} poleId Pole ID or pole-shaped source object.
   * @param {string} [poleHeight]
   * @param {string} [notes]
   * @param {Object} [extra]
   * @returns {Pole}
   */
  // `poleResolution` is the internal name for the mutually exclusive UG/PCO
  // pole mode. It is deliberately non-enumerable so old JSON exports keep
  // their established `ugActive`/`pcoActive` fields while new code can use a
  // single, explicit value.
  function attachPoleResolutionAlias(record) {
    Object.defineProperty(record, "poleResolution", {
      configurable: true,
      get: () => global.PoleResolution.forPole(record),
      set: value => global.PoleResolution.apply(record, value)
    });
    return record;
  }

  function getPoleResolution(pole) {
    return global.PoleResolution.forPole(pole);
  }

  function createPole(poleId, poleHeight = "", notes = "", extra = {}) {
    // Constructors normalize all incoming data so later code can assume each
    // entity has the same shape, whether it came from Excel, JSON, or the UI.
    if (typeof poleId === "object" && poleId !== null) {
      const data = global.StateCompatibility.normalizePoleInput(poleId);
      return attachPoleResolutionAlias({
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
        ugActive: data.ugActive,
        ugSpanId: trim(data.ugSpanId || "ALL") || "ALL",
        ugSpanMRSpanId: trim(data.ugSpanMRSpanId || ""),
        ugSpanMRText: trim(data.ugSpanMRText || ""),
        ugReason: trim(data.ugReason || ""),
        ugMRText: trim(data.ugMRText || ""),
        ugRiserDirection: trim(data.ugRiserDirection || "").toUpperCase(),
        riserActive: data.riserActive === true ? true : data.riserActive === false ? false : null,
        secondaryDripLoopActive: Boolean(data.secondaryDripLoopActive),
        secondaryDripLoopHoa: trim(data.secondaryDripLoopHoa || ""),
        ancActive: data.ancActive === true ? true : data.ancActive === false ? false : null,
        ancSize: trim(data.ancSize || "8\""),
        ancDistance: trim(data.ancDistance || ""),
        ancDistanceManual: data.ancDistanceManual === true,
        ancDirection: trim(data.ancDirection || "").toUpperCase(),
        ancDirectionManual: data.ancDirectionManual === true || Boolean(String(data.ancDirection || "").trim()),
        dgHoa: trim(data.dgHoa || ""),
        ohgActive: Boolean(data.ohgActive),
        ohgHoa: trim(data.ohgHoa || ""),
        ohgDirection: trim(data.ohgDirection || "").toUpperCase(),
        lessThan12CommClearanceActive: Boolean(data.lessThan12CommClearanceActive),
        redTagActive: Boolean(data.redTagActive),
        redTagMinHeight: trim(data.redTagMinHeight || ""),
        redTagClass: trim(data.redTagClass || ""),
        redTagReason: trim(data.redTagReason || ""),
        pcoActive: data.pcoActive,
        commMovementsActive: data.commMovementsActive !== false,
        pcoMRText: trim(data.pcoMRText || ""),
        poleInsetActive: Boolean(data.poleInsetActive),
        poleInsetReason: normalizePoleInsetReason(data.poleInsetReason),
        poleInsetSpanId: trim(data.poleInsetSpanId || ""),
        poleInsetHeight: trim(data.poleInsetHeight || ""),
        pcoScope: trim(data.pcoScope || ""),
        pcoType: trim(data.pcoType || ""),
        pcoDetail: trim(data.pcoDetail || ""),
        notes: trim(data.notes || ""),
        comms: Array.isArray(data.comms) ? data.comms : [],
        metadata: data.metadata || {}
      });
    }

    extra = global.StateCompatibility.normalizePoleInput(extra);
    return attachPoleResolutionAlias({
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
      ugActive: extra.ugActive,
      ugSpanId: trim(extra.ugSpanId || "ALL") || "ALL",
      ugSpanMRSpanId: trim(extra.ugSpanMRSpanId || ""),
      ugSpanMRText: trim(extra.ugSpanMRText || ""),
      ugReason: trim(extra.ugReason || ""),
      ugMRText: trim(extra.ugMRText || ""),
      ugRiserDirection: trim(extra.ugRiserDirection || "").toUpperCase(),
      riserActive: extra.riserActive === true ? true : extra.riserActive === false ? false : null,
      secondaryDripLoopActive: Boolean(extra.secondaryDripLoopActive),
      secondaryDripLoopHoa: trim(extra.secondaryDripLoopHoa || ""),
      ancActive: extra.ancActive === true ? true : extra.ancActive === false ? false : null,
      ancSize: trim(extra.ancSize || "8\""),
      ancDistance: trim(extra.ancDistance || ""),
      ancDistanceManual: extra.ancDistanceManual === true,
      ancDirection: trim(extra.ancDirection || "").toUpperCase(),
      ancDirectionManual: extra.ancDirectionManual === true || Boolean(String(extra.ancDirection || "").trim()),
      dgHoa: trim(extra.dgHoa || ""),
      ohgActive: Boolean(extra.ohgActive),
      ohgHoa: trim(extra.ohgHoa || ""),
      ohgDirection: trim(extra.ohgDirection || "").toUpperCase(),
      lessThan12CommClearanceActive: Boolean(extra.lessThan12CommClearanceActive),
      redTagActive: Boolean(extra.redTagActive),
      redTagMinHeight: trim(extra.redTagMinHeight || ""),
      redTagClass: trim(extra.redTagClass || ""),
      redTagReason: trim(extra.redTagReason || ""),
      pcoActive: extra.pcoActive,
      commMovementsActive: extra.commMovementsActive !== false,
      pcoMRText: trim(extra.pcoMRText || ""),
      poleInsetActive: Boolean(extra.poleInsetActive),
      poleInsetReason: normalizePoleInsetReason(extra.poleInsetReason),
      poleInsetSpanId: trim(extra.poleInsetSpanId || ""),
      poleInsetHeight: trim(extra.poleInsetHeight || ""),
      pcoScope: trim(extra.pcoScope || ""),
      pcoType: trim(extra.pcoType || ""),
      pcoDetail: trim(extra.pcoDetail || ""),
      notes: trim(notes),
      comms: [],
      metadata: extra.metadata || {}
    });
  }

  // Canonical owner names are exposed internally without changing the saved
  // JSON contract.  The legacy enumerable fields remain the import/export
  // adapter, while these aliases give calculations a clearer vocabulary.
  function attachOwnerAliases(record) {
    Object.defineProperties(record, {
      ownerCanonical: {
        configurable: true,
        get: () => trim(record.ownerBase || record.owner || ""),
        set: value => { record.ownerBase = trim(value); }
      },
      ownerImported: {
        configurable: true,
        get: () => trim(record.rawOwner || record.ownerBase || record.owner || ""),
        set: value => { record.rawOwner = trim(value); }
      },
      ownerDisplay: {
        configurable: true,
        get: () => trim(record.owner || record.ownerBase || record.rawOwner || ""),
        set: value => { record.owner = trim(value); }
      }
    });
    return record;
  }

  // Canonical accessors keep business modules independent from the legacy
  // field names that are still required by saved JSON and workbook exports.
  function getOwnerCanonical(record) {
    return trim(record?.ownerCanonical || record?.ownerBase || record?.owner || "");
  }

  function getOwnerImported(record) {
    return trim(record?.ownerImported || record?.rawOwner || getOwnerCanonical(record) || "");
  }

  function getOwnerDisplay(record) {
    return trim(record?.ownerDisplay || record?.owner || getOwnerCanonical(record) || getOwnerImported(record) || "");
  }

  function createComm(owner, existingHOA = "", notes = "", extra = {}) {
    const ownerDisplay = trim(extra.ownerDisplay || owner);
    const ownerCanonical = trim(extra.ownerCanonical || extra.ownerBase || ownerDisplay);
    const ownerImported = trim(extra.ownerImported || extra.rawOwner || "");
    return attachOwnerAliases({
      owner: ownerDisplay,
      ownerBase: ownerCanonical,
      existingHOA: trim(existingHOA),
      existingHOAChange: trim(extra.existingHOAChange || ""),
      otherHOA: trim(extra.otherHOA || ""),
      pofActive: Boolean(extra.pofActive),
      notes: trim(notes),
      rawOwner: ownerImported,
      unknownOwner: Boolean(extra.unknownOwner),
      size: trim(extra.size || ""),
      wireId: trim(extra.wireId || "")
    });
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
  function attachSpanRelationAliases(record) {
    Object.defineProperties(record, {
      spanRelation: {
        configurable: true,
        get: () => trim(record.type || ""),
        set: value => { record.type = trim(value); }
      },
      importedRelation: {
        configurable: true,
        get: () => trim(record.rawType || record.type || ""),
        set: value => { record.rawType = trim(value); }
      }
    });
    return record;
  }

  function getSpanRelation(span) {
    return trim(span?.spanRelation || span?.type || span?.importedRelation || span?.rawType || "");
  }

  function getImportedSpanRelation(span) {
    return trim(span?.importedRelation || span?.rawType || getSpanRelation(span) || "");
  }

  function createSpan(spanId, fromPole, toPole = "", direction = "", notes = "", extra = {}) {
    const id = trim(spanId);
    const fallbackToPole = trim(toPole) || `Unknown-${id || Math.random().toString(36).slice(2, 8)}`;
    return attachSpanRelationAliases({
      spanId: id,
      fromPole: trim(fromPole),
      toPole: fallbackToPole,
      direction: trim(direction),
      bearingDegrees: extra.bearingDegrees === null || extra.bearingDegrees === undefined ? "" : extra.bearingDegrees,
      notes: trim(notes),
      type: trim(extra.spanRelation || extra.type || ""),
      spanIndex: trim(extra.spanIndex || ""),
      length: trim(extra.length || ""),
      lengthDisplay: trim(extra.lengthDisplay || ""),
      environment: trim(extra.environment || "NONE"),
      environmentClearance: trim(extra.environmentClearance || defaultEnvironmentClearance(extra.environment || "NONE")),
      midspanLowPower: trim(extra.midspanLowPower || ""),
      sourceMidspanLowPower: trim(extra.sourceMidspanLowPower || ""),
      midspanMaxCommHeight: trim(extra.midspanMaxCommHeight || ""),
      // A span-level UG decision is separate from the pole-level UG action.
      // null means the operator has not selected Case 2 for this span.
      ugActive: extra.ugActive === true ? true : extra.ugActive === false ? false : null,
      ugReason: trim(extra.ugReason || ""),
      rawSpanIds: Array.isArray(extra.rawSpanIds) ? extra.rawSpanIds : (extra.rawSpanId ? [extra.rawSpanId] : []),
      rawType: trim(extra.importedRelation || extra.rawType || extra.spanRelation || extra.type || ""),
      linkedCollectionId: trim(extra.linkedCollectionId || ""),
      linkedCollectionTitle: trim(extra.linkedCollectionTitle || ""),
      sourceCollectionId: trim(extra.sourceCollectionId || ""),
      sourceSpanId: trim(extra.sourceSpanId || ""),
      isManualProposed: Boolean(extra.isManualProposed),
      isGeneratedOtherPole: Boolean(extra.isGeneratedOtherPole)
    });
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
      isProposedExcluded: Boolean(data.isProposedExcluded),
      isSlack: Boolean(data.isSlack),
      proposedHOA: trim(data.proposedHOA || ""),
      autoCalcProposedStatus: trim(data.autoCalcProposedStatus || ""),
      autoCalcProposedMode: trim(data.autoCalcProposedMode || ""),
      proposedHOAChange: trim(data.proposedHOAChange || ""),
      nextPoleProposedAuto: Boolean(data.nextPoleProposedAuto),
      proposedMidspan: normalizedHeightOrEmpty(data.proposedMidspan),
      ocalcMS: normalizedHeightOrEmpty(data.ocalcMS || data["O-CALC MS"]),
      msProposed: firstNormalizedHeight(data.msProposed, data.proposedMidspan),
      finalMidspan: normalizedHeightOrEmpty(data.finalMidspan),
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
    const ownerDisplay = trim(data.ownerDisplay || data.owner || "");
    const ownerCanonical = trim(data.ownerCanonical || data.ownerBase || ownerDisplay);
    const ownerImported = trim(data.ownerImported || data.rawOwner || "");
    return attachOwnerAliases({
      spanId: trim(data.spanId || ""),
      poleId: trim(data.poleId || ""),
      owner: ownerDisplay,
      ownerBase: ownerCanonical,
      existingHOA: trim(data.existingHOA || ""),
      existingHOAChange: trim(data.existingHOAChange || ""),
      otherHOA: trim(data.otherHOA || ""),
      otherHOAActive: Boolean(data.otherHOAActive),
      serviceDrop: Boolean(data.serviceDrop),
      downGuy: Boolean(data.downGuy),
      transferToNewPole: Boolean(data.transferToNewPole),
      resagServiceDrop: Boolean(data.resagServiceDrop),
      pofActive: Boolean(data.pofActive),
      difference: trim(data.difference || ""),
      remotePoleId: trim(data.remotePoleId || ""),
      remoteHOA: trim(data.remoteHOA || ""),
      ocalcMS: normalizedHeightOrEmpty(data.ocalcMS),
      midspan: normalizedHeightOrEmpty(data.midspan),
      calculatedMidspan: normalizedHeightOrEmpty(data.calculatedMidspan),
      msProposed: firstNormalizedHeight(data.msProposed, data.calculatedMidspan, data.midspan),
      finalMidspan: normalizedHeightOrEmpty(data.finalMidspan),
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
      rawOwner: ownerImported,
      unknownOwner: Boolean(data.unknownOwner),
      size: trim(data.size || ""),
      construction: trim(data.construction || ""),
      insulator: trim(data.insulator || ""),
      wireId: trim(data.wireId || ""),
      wireIndex: trim(data.wireIndex || ""),
      isEndpointPlaceholder: Boolean(data.isEndpointPlaceholder),
      updatedAt: data.updatedAt || ""
    });
  }

  function createSpanPower(data = {}) {
    return {
      spanId: trim(data.spanId || ""),
      poleId: trim(data.poleId || ""),
      label: trim(data.label || ""),
      attachmentHeight: trim(data.attachmentHeight || ""),
      attachmentHeightChange: normalizedHeightOrEmpty(data.attachmentHeightChange),
      midspan: normalizedHeightOrEmpty(data.midspan),
      calculatedMidspan: normalizedHeightOrEmpty(data.calculatedMidspan),
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
    if (wireId) {
      const direct = state.spanComms[keyForSpanComm(spanId, poleId, owner, wireId)];
      if (direct) return direct;
      const normalizedOwner = normalizeImportedCommOwner(owner);
      return Object.values(state.spanComms).find(sc => sc.spanId === spanId
        && sc.poleId === poleId
        && sc.wireId === wireId
        && normalizeImportedCommOwner(sc.owner) === normalizedOwner) || null;
    }
    const exact = state.spanComms[keyForSpanComm(spanId, poleId, owner, "")];
    if (exact) return exact;
    const normalizedOwner = normalizeImportedCommOwner(owner);
    return Object.values(state.spanComms).find(sc => sc.spanId === spanId
      && sc.poleId === poleId
      && normalizeImportedCommOwner(sc.owner) === normalizedOwner) || null;
  }

  function upsertPole(pole) {
    if (!pole || !pole.poleId) return null;
    const existing = state.poles[pole.poleId] || createPole(pole.poleId);
    state.poles[pole.poleId] = attachPoleResolutionAlias({
      ...existing,
      ...pole,
      comms: Array.isArray(pole.comms)
        ? pole.comms.map(comm => attachOwnerAliases({ ...comm }))
        : (existing.comms || []).map(comm => attachOwnerAliases({ ...comm })),
      metadata: { ...(existing.metadata || {}), ...(pole.metadata || {}) }
    });
    return state.poles[pole.poleId];
  }

  function updatePoleField(poleId, field, value) {
    const pole = state.poles[poleId];
    if (!pole) return null;
    if (!["poleHeight", "lowPower", "maxCommHeight", "topComm", "lowComm", "standaloneProposedHOA", "ugSpanId", "ugSpanMRSpanId", "ugSpanMRText", "ugReason", "ugMRText", "pcoMRText", "poleInsetReason", "poleInsetSpanId", "poleInsetHeight", "ugRiserDirection", "secondaryDripLoopActive", "secondaryDripLoopHoa", "ancSize", "ancDistance", "ancDirection", "dgHoa", "ohgHoa", "ohgDirection", "lessThan12CommClearanceActive", "redTagMinHeight", "redTagClass", "redTagReason", "notes", "sequence"].includes(field)) return pole;
    if (["lessThan12CommClearanceActive", "secondaryDripLoopActive"].includes(field)) {
      pole[field] = value === true || String(value).toLowerCase() === "true";
      return pole;
    }
    pole[field] = trim(value);
    if (field === "ancDistance") pole.ancDistanceManual = Boolean(String(value || "").trim());
    if (field === "ancDirection") pole.ancDirectionManual = Boolean(String(value || "").trim());
    if (field === "ugSpanId") pole[field] = pole[field] || "ALL";
    if (["ugRiserDirection", "ancDirection", "ohgDirection"].includes(field)) pole[field] = pole[field].toUpperCase();
    if (field === "poleInsetReason") pole[field] = normalizePoleInsetReason(value);
    return pole;
  }

  function validatePoleRename(poleId, nextPoleId) {
    const currentId = trim(poleId);
    const nextId = trim(nextPoleId);
    if (!currentId || !state.poles[currentId]) return { ok: false, reason: "Pole not found." };
    if (!nextId) return { ok: false, reason: "Pole ID cannot be empty." };
    if (currentId === nextId) return { ok: true, changed: false, previousPoleId: currentId, poleId: currentId };

    const nextCanonical = canonicalPoleIdentity(nextId);
    const duplicateId = Object.keys(state.poles).find(existingId =>
      existingId !== currentId && canonicalPoleIdentity(existingId) === nextCanonical
    );
    if (duplicateId) return { ok: false, reason: `A pole with ID ${duplicateId} already exists.` };
    const deletedId = (state.deletedPoleIds || []).find(id => canonicalPoleIdentity(id) === nextCanonical);
    if (deletedId) return { ok: false, reason: `Pole ID ${nextId} was deleted from this job.` };
    return { ok: true, changed: true, previousPoleId: currentId, poleId: nextId };
  }

  function renamePole(poleId, nextPoleId) {
    const validation = validatePoleRename(poleId, nextPoleId);
    if (!validation.ok || !validation.changed) return validation;
    const currentId = validation.previousPoleId;
    const nextId = validation.poleId;

    const poleIdAliases = { ...(state.poleIdAliases || {}) };
    Object.keys(poleIdAliases).forEach(sourceId => {
      if (poleIdAliases[sourceId] === currentId) poleIdAliases[sourceId] = nextId;
    });
    poleIdAliases[currentId] = nextId;
    state.poleIdAliases = poleIdAliases;

    const remap = value => value === currentId ? nextId : value;
    const remapRow = row => {
      const next = { ...row };
      ["poleId", "remotePoleId", "fromPole", "toPole", "sourcePoleId", "targetPoleId"].forEach(field => {
        if (typeof next[field] === "string") next[field] = remap(next[field]);
      });
      return next;
    };

    const pole = { ...state.poles[currentId], poleId: nextId };
    delete state.poles[currentId];
    state.poles[nextId] = pole;
    Object.values(state.poles).forEach(item => {
      if (Array.isArray(item.comms)) item.comms = item.comms.map(remapRow);
    });

    Object.values(state.spans || {}).forEach(span => {
      span.fromPole = remap(span.fromPole);
      span.toPole = remap(span.toPole);
    });

    const nextSides = {};
    Object.values(state.spanSides || {}).forEach(source => {
      const row = remapRow(source);
      nextSides[keyForSpanSide(row.spanId, row.poleId)] = row;
    });
    state.spanSides = nextSides;

    const nextComms = {};
    Object.values(state.spanComms || {}).forEach(source => {
      const row = attachOwnerAliases(remapRow(source));
      nextComms[keyForSpanComm(row.spanId, row.poleId, getOwnerCanonical(row), row.wireId || "")] = row;
    });
    state.spanComms = nextComms;

    const nextPower = {};
    Object.values(state.spanPower || {}).forEach((source, index) => {
      const row = remapRow(source);
      const key = [row.spanId, row.poleId, row.wireId || row.label || index].join("__");
      nextPower[key] = row;
    });
    state.spanPower = nextPower;

    ["makeReadyReferences", "poleClassChecks", "movements", "mr", "warnings"].forEach(field => {
      if (Array.isArray(state[field])) state[field] = state[field].map(remapRow);
    });
    state.selectedPoleId = remap(state.selectedPoleId);
    state.ui.hiddenPoleIds = (state.ui.hiddenPoleIds || []).map(remap);
    state.ui.revealedUnknownPoleIds = (state.ui.revealedUnknownPoleIds || []).map(remap);
    return validation;
  }

  function updateSetting(field, value) {
    if (!state.settings || !Object.prototype.hasOwnProperty.call(emptyState().settings, field)) return null;
    if (field === "metronetWI" && global.ProjectProfiles) {
      const wi = String(value || "").trim().toUpperCase();
      const profileId = wi === "CSU" ? "CSU" : wi === "COMED" ? "COMED" : "METRONET";
      state.settings = global.ProjectProfiles.applyProfileSettings(state.settings, profileId);
      Object.values(state.spans || {}).forEach(span => {
        span.environmentClearance = defaultEnvironmentClearance(span.environment || "NONE");
      });
      return state.settings;
    }
    if (field === "projectProfile" && global.ProjectProfiles) {
      state.settings = global.ProjectProfiles.applyProfileSettings(state.settings, value);
      Object.values(state.spans || {}).forEach(span => {
        span.environmentClearance = defaultEnvironmentClearance(span.environment || "NONE");
      });
      return state.settings;
    }
    if (field === "autoMoveCommsToProposed") {
      state.settings[field] = value === true || String(value).toLowerCase() === "true";
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
    if (!["environment", "environmentClearance", "midspanLowPower", "midspanMaxCommHeight", "notes", "ugActive", "ugReason"].includes(field)) return span;
    if (field === "ugActive") {
      span.ugActive = value === true ? true : value === false ? false : null;
    } else {
      span[field] = trim(value);
    }
    if (field === "environment" && !span.environmentClearance) {
      span.environmentClearance = defaultEnvironmentClearance(value);
    }
    return span;
  }

  function upsertComm(poleId, owner, existingHOA = "", notes = "", extra = {}) {
    const pole = state.poles[poleId] || upsertPole(createPole(poleId));
    const ownerKey = trim(extra.ownerCanonical || extra.ownerBase || owner);
    const ownerDisplay = trim(extra.ownerDisplay || owner);
    const wireKey = trim(extra.wireId || "");
    let idx = pole.comms.findIndex(c => getOwnerCanonical(c) === ownerKey && (!wireKey || c.wireId === wireKey));
    if (idx < 0) idx = pole.comms.findIndex(c => getOwnerCanonical(c) === ownerKey && !c.wireId && !wireKey);

    if (idx >= 0) {
      pole.comms[idx] = attachOwnerAliases({
        ...pole.comms[idx],
        owner: ownerDisplay,
        ownerBase: ownerKey,
        existingHOA: trim(existingHOA) || pole.comms[idx].existingHOA,
        existingHOAChange: trim(extra.existingHOAChange || pole.comms[idx].existingHOAChange || ""),
        otherHOA: trim(extra.otherHOA || pole.comms[idx].otherHOA || ""),
        pofActive: Boolean(extra.pofActive || pole.comms[idx].pofActive),
        notes: trim(notes) || pole.comms[idx].notes,
        rawOwner: trim(extra.ownerImported || extra.rawOwner || pole.comms[idx].rawOwner || ""),
        unknownOwner: Boolean(extra.unknownOwner || pole.comms[idx].unknownOwner),
        size: trim(extra.size || pole.comms[idx].size || ""),
        wireId: trim(extra.wireId || pole.comms[idx].wireId || "")
      });
      return pole.comms[idx];
    }
    const comm = createComm(ownerDisplay, existingHOA, notes, { ...extra, ownerCanonical: ownerKey });
    pole.comms.push(comm);
    return comm;
  }

  function upsertSpan(span) {
    if (!span || !span.spanId) return null;
    const existing = state.spans[span.spanId] || createSpan(span.spanId, span.fromPole, span.toPole);
    const rawSpanIds = Array.from(new Set([...(existing.rawSpanIds || []), ...(span.rawSpanIds || [])].filter(Boolean)));
    state.spans[span.spanId] = attachSpanRelationAliases({
      ...existing,
      ...span,
      type: trim(span.spanRelation || span.type || existing.spanRelation || existing.type || ""),
      rawType: trim(span.importedRelation || span.rawType || existing.importedRelation || existing.rawType || ""),
      rawSpanIds
    });
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
    const ownerKey = trim(data?.ownerCanonical || data?.ownerBase || data?.owner || "");
    if (!data || !data.spanId || !data.poleId || !ownerKey) return null;
    const key = keyForSpanComm(data.spanId, data.poleId, ownerKey, data.wireId || "");
    const existing = state.spanComms[key] || createSpanComm(data);
    state.spanComms[key] = createSpanComm({
      ...existing,
      ...data,
      owner: data.ownerDisplay || data.owner || existing.owner,
      ownerCanonical: ownerKey,
      ownerImported: data.ownerImported || data.rawOwner || existing.ownerImported || "",
      updatedAt: new Date().toISOString()
    });
    const pole = state.poles[data.poleId];
    const poleComm = pole?.comms?.find(comm =>
      getOwnerCanonical(comm) === ownerKey && (!data.wireId || !comm.wireId || comm.wireId === data.wireId)
    );
    if (poleComm && Object.prototype.hasOwnProperty.call(data, "pofActive")) {
      poleComm.pofActive = Boolean(state.spanComms[key].pofActive);
    }
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

  function isPrimaryPowerRow(row) {
    return /\bprimary\b/i.test(`${row?.label || ""} ${row?.size || ""} ${row?.type || ""}`);
  }

  function updateSpanPowerField(powerKey, field, value) {
    if (!["attachmentHeight", "attachmentHeightChange", "midspan", "calculatedMidspan", "size", "label"].includes(field)) return null;
    const row = state.spanPower[trim(powerKey)];
    if (!row) return null;
    const isSharedPoleHeightEdit = ["attachmentHeight", "attachmentHeightChange"].includes(field);
    const sourceHeight = isSharedPoleHeightEdit
      ? global.HeightUtils?.parseHeight(row.attachmentHeight || "")
      : null;
    row[field] = trim(value);
    if (isSharedPoleHeightEdit && Number.isFinite(sourceHeight)) {
      // Power wires at the same pole and attachment height move as one stack,
      // even when their types or connected spans differ. Each span still keeps
      // its own Midspan and calculation.
      Object.values(state.spanPower).forEach(candidate => {
        if (candidate === row || candidate.poleId !== row.poleId) return;
        const candidateSourceHeight = global.HeightUtils?.parseHeight(candidate.attachmentHeight || "");
        if (candidateSourceHeight === sourceHeight) candidate[field] = trim(value);
      });
    }
    return row;
  }

  function removeSpanPower(powerKey) {
    const key = trim(powerKey);
    const row = state.spanPower[key];
    if (!row) return null;
    delete state.spanPower[key];
    return row;
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
    state.ui.revealedUnknownPoleIds = (state.ui.revealedUnknownPoleIds || []).filter(value => value !== id);
    if (options.remember !== false) {
      const canonical = canonicalPoleIdentity(id);
      state.deletedPoleIds = Array.from(new Set([...(state.deletedPoleIds || []), canonical].filter(Boolean)));
    }
    return { poleId: id, spanIds: Array.from(spanIds), commCount: commRows.length };
  }

  /**
   * Stores the user action attached to one normalized Power Equipment row.
   * The imported physical heights remain unchanged; actionHeight is the
   * proposed redress/raise height used by calculations and Make Ready. INTEC
   * and Metronet MidAm Streetlights keep Ground and Raise as independent user actions. Risers
   * keep Raise and Secure as independent user actions.
   */
  function updatePowerEquipmentField(poleId, equipmentIndex, field, value) {
    if (!["owner", "attachmentHeight", "actionActive", "actionHeight", "raiseActive", "raiseHeight", "secureActive"].includes(field)) return null;
    const pole = state.poles[poleId];
    const rows = pole?.metadata?.powerEquipment;
    const index = Number(equipmentIndex);
    if (!pole || !Array.isArray(rows) || !Number.isInteger(index) || !rows[index]) return null;
    const nextRows = rows.map((row, rowIndex) => {
      if (rowIndex !== index) return row;
      const nextRow = {
        ...row,
        [field]: ["actionActive", "raiseActive", "secureActive"].includes(field) ? Boolean(value) : trim(value)
      };
      const isMidAm = String(state.settings?.projectProfile || "").toUpperCase() === "METRONET"
        && String(state.settings?.proposedOwner || "").toUpperCase() === "MIDAM";
      const isStreetlight = String(row.category || row.type || "").toUpperCase().includes("STREETLIGHT");
      if (field === "raiseActive" && Boolean(value) && isMidAm && isStreetlight) {
        const source = global.HeightUtils?.parseHeight(row.attachmentHeight || "");
        if (source !== null && source !== undefined) {
          nextRow.raiseHeight = global.HeightUtils.formatHeight(source + 12);
        }
      }
      return nextRow;
    });
    pole.metadata = { ...(pole.metadata || {}), powerEquipment: nextRows };
    return nextRows[index];
  }

  function removePowerEquipment(poleId, equipmentIndex) {
    const pole = state.poles[poleId];
    const rows = pole?.metadata?.powerEquipment;
    const index = Number(equipmentIndex);
    if (!pole || !Array.isArray(rows) || !Number.isInteger(index) || !rows[index]) return null;
    const removed = rows[index];
    pole.metadata = { ...(pole.metadata || {}), powerEquipment: rows.filter((_, rowIndex) => rowIndex !== index) };
    return removed;
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
    const commChange = getSpanCommsForPole(poleId).some(sc => sc.existingHOAChange || sc.otherHOA || sc.notes || sc.mr);
    const equipmentChange = (state.poles[poleId]?.metadata?.powerEquipment || [])
      .some(row => Boolean(row.actionActive || trim(row.actionHeight || "") || row.raiseActive || row.secureActive || trim(row.raiseHeight || "")));
    const powerChange = getSpanPowerForPole(poleId).some(row => Boolean(row.attachmentHeightChange));
    return Boolean(state.poles[poleId]?.standaloneProposedHOA || state.poles[poleId]?.poleInsetHeight || state.poles[poleId]?.poleInsetSpanId || state.poles[poleId]?.ugMRText || state.poles[poleId]?.ugRiserDirection || state.poles[poleId]?.secondaryDripLoopActive || state.poles[poleId]?.secondaryDripLoopHoa || state.poles[poleId]?.ancActive || state.poles[poleId]?.ohgActive)
      || state.poles[poleId]?.riserActive === true
      || state.poles[poleId]?.riserActive === false
      || sideChange || commChange || equipmentChange || powerChange;
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
      sourceMidspanLowPower: source.sourceMidspanLowPower || span.sourceMidspanLowPower || "",
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
              otherHOA: comm.otherHOA || "",
              pofActive: Boolean(comm.pofActive),
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
              otherHOA: "",
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
    raw = global.StateCompatibility.unwrapSavedState(raw) || {};
    const next = { ...emptyState(), ...raw };
    next.jobName = trim(next.jobName) || jobNameFromFileName(next.importedFileName);
    const rawSettings = raw && raw.settings ? raw.settings : {};
    const requestedProfile = String(rawSettings.projectProfile || "").trim().toUpperCase();
    const requestedWI = String(rawSettings.metronetWI || "").trim().toUpperCase();
    const profileKey = requestedProfile === "METRONET" && requestedWI === "COMED"
      ? "COMED" : (rawSettings.projectProfile || "INTEC");
    const profileDefaults = global.ProjectProfiles
      ? (global.ProjectProfiles.getProfileSettings
        ? global.ProjectProfiles.getProfileSettings(profileKey)
        : (global.ProjectProfiles.getProfile(profileKey)?.settings || {}))
      : {};
    // Apply project defaults before saved values. Existing user edits win,
    // while older JSON files automatically receive newly introduced profile
    // rules such as MidAm streetlight and Back Span behavior.
    next.settings = { ...emptyState().settings, ...profileDefaults, ...rawSettings };
    next.settings.autoMoveCommsToProposed = next.settings.autoMoveCommsToProposed === true
      || String(next.settings.autoMoveCommsToProposed).toLowerCase() === "true";
    if (requestedProfile === "CSU") {
      next.settings.projectProfile = "METRONET";
      next.settings.metronetWI = "CSU";
    } else if (requestedProfile === "COMED" || (requestedProfile === "METRONET" && requestedWI === "COMED")) {
      next.settings.projectProfile = "METRONET";
      next.settings.metronetWI = "COMED";
      if (!next.settings.proposedOwner || ["METRONET", "MIDAM"].includes(String(next.settings.proposedOwner).toUpperCase())) {
        next.settings.proposedOwner = "Power";
      }
      // Older ComEd saves inherited generic MetroNet values. Replace only
      // those known legacy defaults; keep any explicitly edited clearance.
      const comedLegacyValues = {
        polePowerCommsClearance: "40\"",
        clearanceToPower: "40\"",
        boltClearance: "6\"",
        streetlightBracketCommClearance: "20\""
      };
      const comedDefaults = global.ProjectProfiles?.getProfileSettings?.("COMED") || {};
      Object.entries(comedLegacyValues).forEach(([field, legacy]) => {
        if (!rawSettings || !Object.prototype.hasOwnProperty.call(rawSettings, field) || rawSettings[field] === legacy) {
          if (Object.prototype.hasOwnProperty.call(comedDefaults, field)) next.settings[field] = comedDefaults[field];
        }
      });
    }
    // Back Span calculation is a profile rule, not an operator preference.
    // Older INTEC saves persisted the former `false` default, which made an
    // imported Back Span midspan visible but prevented it from responding to
    // endpoint HOA movements. Always migrate that setting to the active
    // profile value when the profile defines it.
    if (Object.prototype.hasOwnProperty.call(profileDefaults, "calculateBackspanMidspan")) {
      next.settings.calculateBackspanMidspan = profileDefaults.calculateBackspanMidspan === true;
    }
    // Migrate Metronet saves created before the WI selector existed.
    if (String(next.settings.projectProfile || "").toUpperCase() === "METRONET"
      && String(next.settings.metronetWI || "").toUpperCase() !== "COMED"
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
    next.poleIdAliases = next.poleIdAliases && typeof next.poleIdAliases === "object"
      ? { ...next.poleIdAliases }
      : {};
    next.movements = Array.isArray(next.movements) ? next.movements : [];
    next.mr = Array.isArray(next.mr) ? next.mr : [];
    next.warnings = Array.isArray(next.warnings) ? next.warnings : [];
    next.deletedPoleIds = Array.isArray(next.deletedPoleIds)
      ? Array.from(new Set(next.deletedPoleIds.map(canonicalPoleIdentity).filter(Boolean)))
      : [];
    next.ui = { search: "", filter: "all", hiddenPoleIds: [], revealedUnknownPoleIds: [], ...(next.ui || {}) };
    // Remove presentation keys from saves created while per-section comm
    // visibility existed. Comm tables are now always visible.
    delete next.ui.hiddenCommPoleIds;
    next.ui.hiddenPoleIds = Array.isArray(next.ui.hiddenPoleIds)
      ? Array.from(new Set(next.ui.hiddenPoleIds.map(trim).filter(id => Boolean(id && next.poles[id]))))
      : [];
    next.ui.revealedUnknownPoleIds = Array.isArray(next.ui.revealedUnknownPoleIds)
      ? Array.from(new Set(next.ui.revealedUnknownPoleIds.map(trim).filter(id => Boolean(id && next.poles[id]))))
      : [];

    // A physical span cannot terminate at the same pole where it starts. Old
    // imports with an ambiguous Collection ID could persist that self-link;
    // normalize it to the editable generated endpoint used by the importer.
    Object.keys(next.spans).forEach(id => {
      const span = next.spans[id];
      const fromPole = trim(span?.fromPole || "");
      const toPole = trim(span?.toPole || "");
      if (!fromPole || !toPole || fromPole.toLowerCase() !== toPole.toLowerCase()) return;
      const rawSpanId = trim(span.rawSpanId || span.spanId || id) || id;
      next.spans[id] = {
        ...span,
        toPole: `Unknown-${rawSpanId}`,
        linkedCollectionId: "",
        linkedCollectionTitle: "",
        isGeneratedOtherPole: true
      };
    });

    Object.keys(next.poles).forEach(id => {
      const pole = next.poles[id];
      next.poles[id] = {
        ...createPole(id),
        ...pole,
        poleId: pole.poleId || id,
        // Older saves can contain a selected direction but a false manual
        // marker. The populated field is the user-visible ANC choice.
        ancDirectionManual: pole.ancDirectionManual === true || Boolean(String(pole.ancDirection || "").trim()),
        comms: Array.isArray(pole.comms) ? pole.comms.map(c => createComm(c.owner, c.existingHOA, c.notes, c)) : []
      };
    });

    // Older ComEd imports were incorrectly marked as mandatory MidAm Ground
    // actions because both workflows use the METRONET calculation profile.
    // Clear only that imported marker; a later manual checkbox remains intact.
    const isComedState = String(next.settings.projectProfile || "").toUpperCase() === "METRONET"
      && String(next.settings.metronetWI || "").toUpperCase() === "COMED";
    if (isComedState) {
      Object.keys(next.poles).forEach(id => {
        const pole = next.poles[id];
        const equipment = Array.isArray(pole.metadata?.powerEquipment) ? pole.metadata.powerEquipment : [];
        const cleaned = equipment.map(item => {
          const isStreetlight = String(item.category || item.type || "").toUpperCase().includes("STREETLIGHT");
          return isStreetlight && item.groundingRequired === true
            ? { ...item, groundingRequired: false, actionActive: false }
            : item;
        });
        if (cleaned.some((item, index) => item !== equipment[index])) {
          next.poles[id] = { ...pole, metadata: { ...(pole.metadata || {}), powerEquipment: cleaned } };
        }
      });
    }

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
      const owner = normalizeImportedCommOwner(getOwnerCanonical(clean) || getOwnerImported(clean));
      const displayOwner = getOwnerDisplay(clean) || owner;
      const importedOwner = getOwnerImported(clean) || owner;
      if (!clean.spanId || !clean.poleId || !owner) return;
      const normalized = createSpanComm({
        ...clean,
        owner: displayOwner,
        ownerCanonical: owner,
        ownerDisplay: displayOwner,
        ownerImported: importedOwner
      });
      const key = keyForSpanComm(normalized.spanId, normalized.poleId, getOwnerCanonical(normalized), normalized.wireId || "");
      const previous = normalizedSpanComms[key];
      if (!previous) {
        normalizedSpanComms[key] = normalized;
        return;
      }
      // Older states can contain both the raw-owner and normalized-owner key
      // for the same physical wire. Merge them while preserving any manual
      // movement, transfer, or midspan value already present.
      normalizedSpanComms[key] = createSpanComm({
        ...previous,
        ...normalized,
        existingHOA: normalized.existingHOA || previous.existingHOA,
        existingHOAChange: normalized.existingHOAChange || previous.existingHOAChange,
        otherHOA: normalized.otherHOA || previous.otherHOA,
        remoteHOA: normalized.remoteHOA || previous.remoteHOA,
        midspan: normalized.midspan || previous.midspan,
        ocalcMS: normalized.ocalcMS || previous.ocalcMS,
        calculatedMidspan: normalized.calculatedMidspan || previous.calculatedMidspan,
        msProposed: normalized.msProposed || previous.msProposed,
        finalMidspan: normalized.finalMidspan || previous.finalMidspan,
        mr: normalized.mr || previous.mr,
        notes: normalized.notes || previous.notes,
        serviceDrop: normalized.serviceDrop || previous.serviceDrop,
        downGuy: normalized.downGuy || previous.downGuy,
        transferToNewPole: normalized.transferToNewPole || previous.transferToNewPole,
        resagServiceDrop: normalized.resagServiceDrop || previous.resagServiceDrop,
        pofActive: normalized.pofActive || previous.pofActive
      });
    });
    next.spanComms = normalizedSpanComms;
    repairImportedSpanComms(next);

    const normalizedPower = {};
    Object.values(next.spanPower || {}).forEach(row => {
      const clean = createSpanPower(row);
      if (clean.spanId && clean.poleId) normalizedPower[[clean.spanId, clean.poleId, clean.wireId || clean.label || Object.keys(normalizedPower).length].join("__")] = clean;
    });
    next.spanPower = normalizedPower;

    state = next;
    ensureUnknownPoles();
    const revealedUnknown = new Set(state.ui.revealedUnknownPoleIds || []);
    const hiddenPoleIds = new Set(state.ui.hiddenPoleIds || []);
    // Collection is the authoritative list of poles that belong in the main
    // workspace. Span endpoints that are not present there remain addressable
    // in the index, but stay hidden until the operator explicitly reveals one.
    const collectionRows = state.excelReviewSource?.collection?.rows || [];
    const collectionPoleIds = new Set();
    collectionRows.forEach(row => {
      const key = Object.keys(row || {}).find(candidate => {
        const normalized = String(candidate).toLowerCase().replace(/[^a-z0-9]/g, "");
        return ["id", "poleid", "polename", "structurenumber", "pole"].includes(normalized);
      });
      const value = key ? trim(row[key]) : "";
      if (value) collectionPoleIds.add(canonicalPoleIdentity(value));
    });
    if (collectionPoleIds.size) {
      Object.keys(state.poles).forEach(id => {
        if (collectionPoleIds.has(canonicalPoleIdentity(id))) hiddenPoleIds.delete(id);
        else if (!revealedUnknown.has(id)) hiddenPoleIds.add(id);
      });
    } else {
      Object.keys(state.poles).forEach(id => {
        if (isAddressableUnknownPoleId(id) && !revealedUnknown.has(id)) hiddenPoleIds.add(id);
      });
    }
    state.ui.hiddenPoleIds = Array.from(hiddenPoleIds);
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
    jobNameFromFileName,
    getState,
    setState,
    resetState,
    saveToLocal,
    loadFromLocal,
    createPole,
    getPoleResolution,
    createComm,
    getOwnerCanonical,
    getOwnerImported,
    getOwnerDisplay,
    createSpan,
    getSpanRelation,
    getImportedSpanRelation,
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
    validatePoleRename,
    renamePole,
    updateSetting,
    applyProjectProfile,
    updateSpanField,
    updateSpanPowerField,
    updatePowerEquipmentField,
    removePowerEquipment,
    upsertComm,
    upsertSpan,
    upsertSpanSide,
    removeSpanSide,
    removeManualSpan,
    upsertSpanComm,
    removeSpanComm,
    removePole,
    addSpanPower,
    removeSpanPower,
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
