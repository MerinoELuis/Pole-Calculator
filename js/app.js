(function (global) {
  "use strict";

  // app.js owns the browser UI: rendering tables, binding controls, dispatching
  // user edits to the store, and refreshing affected poles after recalculation.
  const S = global.AppStore;
  const H = global.HeightUtils;

  const els = {};
  let delayedMidspanRenderTimer = null;
  const delayedMidspanRenderPoleIds = new Set();
  const editableInputTimers = new Map();
  const undoHistory = [];
  const MAX_UNDO_STEPS = 100;
  let restoringUndo = false;
  let saveFileHandle = null;
  let lastSavedSerialized = "";
  let hasUnsavedChanges = false;
  const SPAN_COLOR_CLASS_COUNT = 5;
  const FILE_HANDLE_DB = "poleCalculatorFileHandles";
  const FILE_HANDLE_STORE = "handles";
  const SAVE_HANDLE_KEY = "currentSaveFile";
  const JSON_PICKER_ID = "pole-calculator-json";


  function qs(id) { return document.getElementById(id); }

  function escapeHtml(value) {
    return String(value ?? "")
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&#039;");
  }

  function toast(message, type = "info") {
    if (!els.toastHost) return;
    const div = document.createElement("div");
    div.className = `toast ${type}`;
    div.textContent = message;
    els.toastHost.appendChild(div);
    setTimeout(() => div.remove(), 4200);
  }

  function setPoleIndexOpen(open) {
    const isOpen = Boolean(open);
    document.body.classList.toggle("pole-index-open", isOpen);
    els.poleIndexDrawer?.classList.toggle("open", isOpen);
    els.poleIndexDrawer?.setAttribute("aria-hidden", String(!isOpen));
    if (els.poleIndexDrawer) els.poleIndexDrawer.inert = !isOpen;
    els.poleIndexToggle?.setAttribute("aria-expanded", String(isOpen));
    els.poleIndexBackdrop?.classList.toggle("hidden", !isOpen);
    if (isOpen) els.poleSearchInput?.focus();
  }

  function updatePoleIndexToggleVisibility() {
    if (!els.poleIndexToggle || !els.topIndexPanel) return;
    if ((S.getState().ui.activeView || "calculator") !== "calculator") {
      els.poleIndexToggle.classList.remove("visible");
      return;
    }
    const panelIsAboveViewport = els.topIndexPanel.getBoundingClientRect().bottom <= 0;
    els.poleIndexToggle.classList.toggle("visible", panelIsAboveViewport);
  }

  // Native prompt/confirm dialogs ignore the app theme. These small helpers
  // build one reusable in-page dialog so data-entry actions keep the same look
  // and keyboard behavior as the rest of the calculator.
  function closeAppDialog(result = null) {
    const dialog = document.querySelector("[data-app-dialog]");
    if (!dialog) return;
    const resolve = dialog._resolveDialog;
    dialog.remove();
    if (resolve) resolve(result);
  }

  function openAppDialog({ title, description = "", fields = [], confirmLabel = "Accept", danger = false }) {
    closeAppDialog(null);
    return new Promise(resolve => {
      const dialog = document.createElement("div");
      dialog.className = "app-dialog-backdrop";
      dialog.dataset.appDialog = "true";
      dialog._resolveDialog = resolve;
      dialog.innerHTML = `<section class="app-dialog" role="dialog" aria-modal="true" aria-label="${escapeHtml(title)}">
        <div class="app-dialog-header">
          <h3>${escapeHtml(title)}</h3>
          ${description ? `<p>${escapeHtml(description)}</p>` : ""}
        </div>
        <form class="app-dialog-form">
          ${fields.map(field => {
            if (field.type === "checkbox-list") {
              return `<fieldset class="dialog-checkbox-list">
                <legend>${escapeHtml(field.label)}</legend>
                ${field.options.map(option => `<label>
                  <input type="checkbox" name="${escapeHtml(field.name)}" value="${escapeHtml(option.value)}" ${option.checked ? "checked" : ""}>
                  <span>${escapeHtml(option.label)}</span>
                </label>`).join("")}
              </fieldset>`;
            }
            return `<label class="dialog-field">
              <span>${escapeHtml(field.label)}</span>
              <input class="input" name="${escapeHtml(field.name)}" value="${escapeHtml(field.value || "")}" placeholder="${escapeHtml(field.placeholder || "")}">
            </label>`;
          }).join("")}
          <div class="app-dialog-actions">
            <button class="btn" type="button" data-dialog-cancel>Cancel</button>
            <button class="btn ${danger ? "btn-danger-solid" : "btn-primary"}" type="submit">${escapeHtml(confirmLabel)}</button>
          </div>
        </form>
      </section>`;
      document.body.appendChild(dialog);
      const form = dialog.querySelector("form");
      dialog.querySelector("[data-dialog-cancel]").addEventListener("click", () => closeAppDialog(null));
      dialog.addEventListener("click", event => {
        if (event.target === dialog) closeAppDialog(null);
      });
      dialog.addEventListener("keydown", event => {
        if (event.key === "Escape") closeAppDialog(null);
      });
      form.addEventListener("submit", event => {
        event.preventDefault();
        const result = {};
        fields.forEach(field => {
          if (field.type === "checkbox-list") {
            result[field.name] = Array.from(form.querySelectorAll(`[name="${CSS.escape(field.name)}"]:checked`)).map(input => input.value);
            return;
          }
          result[field.name] = form.elements[field.name]?.value || "";
        });
        closeAppDialog(result);
      });
      const firstInput = dialog.querySelector("input");
      if (firstInput) firstInput.focus();
    });
  }

  async function confirmInApp(title, description, confirmLabel = "Delete") {
    const result = await openAppDialog({ title, description, confirmLabel, danger: true });
    return Boolean(result);
  }

  function cloneCurrentState() {
    return JSON.parse(JSON.stringify(S.getState()));
  }

  function recordUndoSnapshot() {
    if (restoringUndo) return;
    const snapshot = cloneCurrentState();
    const serialized = JSON.stringify(snapshot);
    const previous = undoHistory[undoHistory.length - 1];
    if (previous && previous.serialized === serialized) return;
    undoHistory.push({ serialized, snapshot });
    if (undoHistory.length > MAX_UNDO_STEPS) undoHistory.shift();
    markDirty();
  }

  function safeJobFilePart(value) {
    const raw = String(value || "pole_job")
      .replace(/\.[^.]+$/, "")
      // A loaded save may already include the app suffix. Remove it before the
      // date so names such as JOB_2026-07-13_Pole_Calculator are also cleaned.
      .replace(/_(?:Pole_Calculator|AutoProposed|Debug)$/i, "")
      .replace(/^excel[_\-\s]*/i, "")
      .replace(/[_\-\s]*20\d{2}[-_]\d{2}[-_]\d{2}$/i, "")
      .trim();
    return (raw || "pole_job")
      .replace(/[<>:"/\\|?*\u0000-\u001F]+/g, "_")
      .replace(/\s+/g, "_");
  }

  function saveFileName() {
    const state = S.getState();
    return `${safeJobFilePart(state.importedFileName || "pole_job")}_Pole_Calculator.json`;
  }

  function openFileHandleDb() {
    return new Promise((resolve, reject) => {
      if (!window.indexedDB) return resolve(null);
      const request = indexedDB.open(FILE_HANDLE_DB, 1);
      request.onupgradeneeded = () => request.result.createObjectStore(FILE_HANDLE_STORE);
      request.onsuccess = () => resolve(request.result);
      request.onerror = () => reject(request.error);
    });
  }

  async function storeFileHandle(handle) {
    if (!handle) return;
    const db = await openFileHandleDb();
    if (!db) return;
    await new Promise((resolve, reject) => {
      const tx = db.transaction(FILE_HANDLE_STORE, "readwrite");
      tx.objectStore(FILE_HANDLE_STORE).put(handle, SAVE_HANDLE_KEY);
      tx.oncomplete = resolve;
      tx.onerror = () => reject(tx.error);
    });
    db.close();
  }

  async function readStoredFileHandle() {
    const db = await openFileHandleDb();
    if (!db) return null;
    const handle = await new Promise((resolve, reject) => {
      const tx = db.transaction(FILE_HANDLE_STORE, "readonly");
      const request = tx.objectStore(FILE_HANDLE_STORE).get(SAVE_HANDLE_KEY);
      request.onsuccess = () => resolve(request.result || null);
      request.onerror = () => reject(request.error);
    });
    db.close();
    return handle;
  }

  function savePayload() {
    global.Calculations.recalculateAll();
    const state = cloneCurrentState();
    return {
      app: "pole-calculator",
      exportedAt: new Date().toISOString(),
      version: state.version || S.CURRENT_VERSION,
      sourceFile: state.importedFileName || "",
      state
    };
  }

  function serializedSavePayload() {
    return JSON.stringify(savePayload(), null, 2);
  }

  function markDirty() {
    hasUnsavedChanges = true;
    updateSaveButtonState();
  }

  function markClean(serialized = "") {
    lastSavedSerialized = serialized || JSON.stringify(savePayload());
    hasUnsavedChanges = false;
    updateSaveButtonState();
  }

  function updateSaveButtonState() {
    if (!els.saveLocalBtn) return;
    els.saveLocalBtn.textContent = hasUnsavedChanges ? "Save *" : "Save";
    els.saveLocalBtn.title = saveFileHandle
      ? "Save changes to the selected JSON file"
      : "Choose where to save the job JSON file";
  }

  async function writeSaveFile(handle, text) {
    const writable = await handle.createWritable();
    await writable.write(text);
    await writable.close();
  }

  async function saveLocalFile() {
    const text = serializedSavePayload();
    const suggestedName = saveFileName();
    if (window.showSaveFilePicker) {
      // A persisted handle is only a Load location hint. Reusing it here would
      // overwrite the previous job after a reload or a new raw Excel import.
      if (!saveFileHandle) {
        saveFileHandle = await window.showSaveFilePicker({
          id: JSON_PICKER_ID,
          suggestedName,
          types: [{
            description: "Pole Calculator JSON",
            accept: { "application/json": [".json"] }
          }]
        });
      }
      await writeSaveFile(saveFileHandle, text);
      await storeFileHandle(saveFileHandle);
      markClean(text);
      toast(`Saved ${saveFileHandle.name || suggestedName}.`, "success");
      return true;
    }

    global.ProjectExport.downloadJson(suggestedName, JSON.parse(text));
    markClean(text);
    toast("Browser file overwrite is not available; JSON was downloaded.", "warning");
    return true;
  }

  async function loadLocalFile() {
    let file = null;
    if (window.showOpenFilePicker) {
      const pickerOptions = {
        id: JSON_PICKER_ID,
        multiple: false,
        types: [{
          description: "Pole Calculator JSON",
          accept: { "application/json": [".json"] }
        }]
      };
      // File System Access pickers can use the previous file handle as a
      // starting location while still opening the picker every time.
      const storedHandle = await readStoredFileHandle();
      if (storedHandle) pickerOptions.startIn = storedHandle;
      let handles;
      try {
        handles = await window.showOpenFilePicker(pickerOptions);
      } catch (error) {
        // Some browser versions support the picker but reject a file handle in
        // startIn. Retry normally instead of making Load unusable.
        if (!pickerOptions.startIn || error?.name !== "TypeError") throw error;
        delete pickerOptions.startIn;
        handles = await window.showOpenFilePicker(pickerOptions);
      }
      const [handle] = handles;
      saveFileHandle = handle;
      await storeFileHandle(saveFileHandle);
      file = handle ? await handle.getFile() : null;
    } else {
      file = await new Promise(resolve => {
        const input = document.createElement("input");
        input.type = "file";
        input.accept = ".json,application/json";
        input.addEventListener("change", () => resolve(input.files?.[0] || null), { once: true });
        input.click();
      });
    }
    if (!file) return false;
    recordUndoSnapshot();
    await global.ExcelImport.importJsonFile(file);
    global.ExcelReview.clearResults();
    if (S.getState().excelReviewSource?.collection?.rows?.length) global.ExcelReview.runReview();
    render();
    markClean(serializedSavePayload());
    toast("JSON loaded.", "success");
    return true;
  }

  function spanSideHasUserWork(side) {
    return Boolean(side?.proposedHOA || side?.proposedHOAChange || side?.proposedMidspan || side?.ocalcMS || side?.notes || side?.isManualProposed);
  }

  function spanCommHasUserWork(row) {
    return Boolean(
      row?.existingHOAChange ||
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

    const owner = normalizedCommOwner(oldRow.ownerBase || oldRow.owner);
    const logicalCandidates = Object.entries(spanComms).filter(([, row]) =>
      row.spanId === oldRow.spanId &&
      row.poleId === oldRow.poleId &&
      normalizedCommOwner(row.ownerBase || row.owner) === owner
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
    return {
      ...importedRow,
      existingHOAChange: oldRow.existingHOAChange || importedRow.existingHOAChange || "",
      serviceDrop: Boolean(oldRow.serviceDrop || importedRow.serviceDrop),
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
    const ownerKey = normalizedCommOwner(row.ownerBase || row.owner);
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

  function mergeImportedUpdate(previous, imported) {
    const merged = JSON.parse(JSON.stringify(imported));
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
      blankValuesPreserved: 0,
      missingRowsPreserved: 0
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
      }
    });

    Object.entries(previous.poles || {}).forEach(([poleId, oldPole]) => {
      const targetPoleId = remapPoleId(poleId, poleAliases);
      const mappedOldPole = { ...oldPole, poleId: targetPoleId };
      if (!merged.poles[targetPoleId]) {
        merged.poles[targetPoleId] = mappedOldPole;
        reconciliation.missingRowsPreserved += 1;
        return;
      }
      const preservedPole = preserveValuesMissingFromUpdate(merged.poles[targetPoleId], mappedOldPole, reconciliation);
      merged.poles[targetPoleId] = {
        ...preservedPole,
        poleId: targetPoleId,
        ugActive: Boolean(mappedOldPole.ugActive || preservedPole.ugActive),
        pcoActive: Boolean(mappedOldPole.pcoActive || preservedPole.pcoActive),
        standaloneProposedHOA: mappedOldPole.standaloneProposedHOA || preservedPole.standaloneProposedHOA || "",
        notes: mappedOldPole.notes || preservedPole.notes || ""
      };
    });

    merged.makeReadyReferences = mergeMakeReadyReferences(
      (previous.makeReadyReferences || []).map(row => ({ ...row, poleId: remapPoleId(row.poleId, poleAliases) })),
      merged.makeReadyReferences || [],
      reconciliation
    );

    const importedPoleClassById = new Map((merged.poleClassChecks || []).map(row => [row.poleId, row]));
    (previous.poleClassChecks || []).forEach(oldRowSource => {
      const oldRow = { ...oldRowSource, poleId: remapPoleId(oldRowSource.poleId, poleAliases) };
      const importedRow = importedPoleClassById.get(oldRow.poleId);
      if (!importedRow) return;
      const preserved = preserveValuesMissingFromUpdate(importedRow, oldRow, reconciliation);
      importedPoleClassById.set(
        oldRow.poleId,
        global.ExcelImport?.recalculatePoleClassCheck
          ? global.ExcelImport.recalculatePoleClassCheck(preserved)
          : preserved
      );
    });
    merged.poleClassChecks = Array.from(importedPoleClassById.values());

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
    merged.updateDiagnostics = {
      updatedAt: new Date().toISOString(),
      ...reconciliation
    };
    merged.ui = previous.ui || merged.ui || {};
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
      ["Pole", "poles", ["poleHeight", "lowPower", "poleType", "standaloneProposedHOA", "ugActive", "pcoActive"]],
      ["Span", "spans", ["fromPole", "toPole", "type", "direction", "bearingDegrees", "lengthDisplay", "environment"]],
      ["Proposed", "spanSides", ["proposedHOA", "proposedHOAChange", "endDrop", "ocalcMS", "msProposed", "finalMidspan", "clearanceMSStatus", "proposedFlaggingStatus"]],
      ["Comm", "spanComms", ["owner", "existingHOA", "existingHOAChange", "remoteHOA", "midspan", "calculatedMidspan", "finalMidspan", "flaggingStatus", "serviceDrop", "downGuy", "transferToNewPole", "resagServiceDrop", "isEndpointPlaceholder"]],
      ["Power", "spanPower", ["attachmentHeight", "midspan", "owner", "size"]]
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

  function undoLastAction() {
    const previous = undoHistory.pop();
    if (!previous) return toast("No changes to undo.", "info");
    restoringUndo = true;
    clearTimeout(delayedMidspanRenderTimer);
    delayedMidspanRenderTimer = null;
    delayedMidspanRenderPoleIds.clear();
    S.setState(previous.snapshot);
    global.Calculations.recalculateAll();
    if (S.getState().excelReviewSource?.collection?.rows?.length) global.ExcelReview.runReview();
    render();
    restoringUndo = false;
    markDirty();
    toast("Last change undone.", "success");
  }

  function spanLabel(span) {
    if (!span) return "";
    return `${span.fromPole} → ${span.toPole || "Unknown"}`;
  }

  function shortSpanLabel(span) {
    if (!span) return "";
    return `${span.fromPole || "?"} → ${span.toPole || "Unknown"}`;
  }

  function spanLengthDisplay(span) {
    if (!span) return "";
    return span.lengthDisplay || "";
  }

  function poleLink(poleId) {
    if (!poleId) return "";
    return `<button class="link-button" type="button" data-scroll-to-pole="${escapeHtml(poleId)}">${escapeHtml(poleId)}</button>`;
  }

  function displayMidspan(sc) {
    return global.Calculations.displayMidspanForComm(sc);
  }

  function displayDecimalFeetInput(primary, fallback = "") {
    if (primary) return primary;
    const parsed = H.parseHeight(fallback);
    if (parsed === null) return fallback || "";
    return String(Number((parsed / 12).toFixed(3)));
  }

  function connectedSpansSorted(poleId) {
    return S.getConnectedSpans(poleId).sort((a, b) =>
      String(a.spanIndex || a.spanId).localeCompare(String(b.spanIndex || b.spanId), undefined, { numeric: true })
    );
  }

  function spanColorClass(poleId, spanId) {
    const spans = connectedSpansSorted(poleId);
    const index = Math.max(0, spans.findIndex(span => span.spanId === spanId));
    return `span-color-${index % SPAN_COLOR_CLASS_COUNT}`;
  }

  function spanRowClasses(poleId, spanId) {
    return `span-color-row ${spanColorClass(poleId, spanId)}`;
  }

  function spanChip(poleId, spanId) {
    return `<span class="span-color-chip ${spanColorClass(poleId, spanId)}" aria-hidden="true"></span>`;
  }

  function spanColorDot(poleId, spanId) {
    return `<span class="span-color-dot ${spanColorClass(poleId, spanId)}" aria-hidden="true"></span>`;
  }

  function commOwnerLabel(sc) {
    const label = global.Calculations.commOwnerLabel
      ? global.Calculations.commOwnerLabel(sc)
      : String(sc?.rawOwner || sc?.owner || "").trim();
    return label || "No owner";
  }

  function normalizedHeightLabel(value) {
    const parsed = H.parseHeight(value || "");
    return parsed === null ? String(value || "").trim() : H.formatHeight(parsed);
  }

  function commGroupKey(sc) {
    return [
      commOwnerLabel(sc).toLowerCase(),
      normalizedHeightLabel(sc?.existingHOA || "").toLowerCase()
    ].join("|");
  }

  function spanHasRealMidspan(spanId) {
    const hasCommMidspan = S.getSpanCommsForSpan(spanId).some(sc => H.parseHeight(sc.ocalcMS || sc.midspan || "") !== null);
    const hasProposedMidspan = S.getSpanSidesForSpan(spanId).some(side => H.parseHeight(side.ocalcMS || side.proposedMidspan || side.msProposed || "") !== null);
    return hasCommMidspan || hasProposedMidspan;
  }

  function isForespanForProposed(span, poleId) {
    const type = String(span?.type || span?.rawType || "").toLowerCase();
    if (/fore\s*span|forespan/.test(type)) return span.fromPole === poleId;
    if (/back\s*span|backspan|other/.test(type)) return false;
    return span?.fromPole === poleId;
  }

  function proposedSpansForPole(poleId) {
    const seen = new Set();
    const allowNoMidspan = S.getState().settings?.proposeForeSpanWithoutMidspan === true;
    return connectedSpansSorted(poleId)
      .filter(span => isForespanForProposed(span, poleId) || S.getSpanSide(span.spanId, poleId)?.isManualProposed)
      .filter(span => allowNoMidspan || spanHasRealMidspan(span.spanId) || S.getSpanSide(span.spanId, poleId)?.isManualProposed)
      .filter(span => {
        const side = S.getSpanSide(span.spanId, poleId);
        const key = side?.isAdditionalProposed
          ? `${span.fromPole || ""}->${span.toPole || ""}->${span.spanId}`
          : `${span.fromPole || ""}->${span.toPole || ""}`;
        if (seen.has(key)) return false;
        seen.add(key);
        return true;
      });
  }

  function groupedCommsForPole(poleId) {
    const groups = new Map();
    S.getSpanCommsForPole(poleId).forEach(sc => {
      const key = commGroupKey(sc);
      if (!groups.has(key)) {
        groups.set(key, {
          key,
          owner: commOwnerLabel(sc),
          existingHOA: normalizedHeightLabel(sc.existingHOA || ""),
          existingHOAChange: sc.existingHOAChange || "",
          isPof: false,
          rows: []
        });
      }
      const group = groups.get(key);
      if (!group.existingHOAChange && sc.existingHOAChange) group.existingHOAChange = sc.existingHOAChange;
      if (global.Calculations.isPofComm && global.Calculations.isPofComm(sc)) group.isPof = true;
      group.rows.push(sc);
    });
    return Array.from(groups.values()).sort((a, b) => {
      // The table represents the original physical pole stack, so sorting stays
      // based on Existing HOA instead of HOA Change.
      const aHeight = H.parseHeight(a.existingHOA || "");
      const bHeight = H.parseHeight(b.existingHOA || "");
      if (aHeight !== null && bHeight !== null && aHeight !== bHeight) return bHeight - aHeight;
      if (aHeight !== null && bHeight === null) return -1;
      if (aHeight === null && bHeight !== null) return 1;
      return `${a.owner}${a.existingHOA}`.localeCompare(`${b.owner}${b.existingHOA}`, undefined, { numeric: true });
    });
  }

  function commMidspanEntries(group, poleId) {
    // One comm can touch more than one span. The Span, Max Height at MS,
    // Midspan and Other Pole HOA columns all come from this list so every
    // stacked value stays visually aligned row by row.
    const seen = new Set();
    const entries = [];
    const rows = [...group.rows].sort((a, b) =>
      `${a.spanId}${a.wireIndex || ""}`.localeCompare(`${b.spanId}${b.wireIndex || ""}`, undefined, { numeric: true })
    );
    rows.forEach(sc => {
      const span = S.getSpan(sc.spanId);
      const spanType = String(span?.type || span?.rawType || "").toLowerCase();
      const ownMidspan = H.parseHeight(sc.midspan || sc.ocalcMS || "") !== null;
      const isBackspan = /back\s*span|backspan/.test(spanType);
      const isOtherReference = /other/.test(spanType) && !ownMidspan;
      const isReferenceSpan = isBackspan || isOtherReference;
      const calculatedBackspan = Boolean(global.Calculations.isCalculatedBackspanComm?.(sc));
      const midspan = displayMidspan(sc);
      const hasMidspan = H.parseHeight(midspan || "") !== null;
      const hasStoredMidspan = H.parseHeight(sc.midspan || sc.ocalcMS || sc.calculatedMidspan || sc.msProposed || sc.finalMidspan || "") !== null;
      const key = `${sc.spanId}|${midspan || ""}|${hasMidspan ? "ms" : "ref"}`;
      if (seen.has(key)) return;
      seen.add(key);
      const remote = global.Calculations.findRemoteComm(sc.spanId, sc.poleId, sc.ownerBase || sc.owner, sc.wireId || "");
      const midspanLocked = Boolean(sc.existingHOAChange || remote?.existingHOAChange);
      // Backspans are reference-only here, but Fore/Other rows without an
      // imported midspan must stay editable so the user can create the MS.
      const canEditMidspan = !isBackspan && (!midspanLocked || !ownMidspan);
      const rawMidspan = H.parseHeight(sc.calculatedMidspan || sc.midspan || sc.ocalcMS || "");
      const resagTarget = H.parseHeight("15'6\"");
      const canResagServiceDrop = Boolean(
        !isReferenceSpan
        && sc.serviceDrop
        && (sc.resagServiceDrop || (rawMidspan !== null && resagTarget !== null && rawMidspan < resagTarget))
      );
      entries.push({
        spanHtml: `<div class="comm-span-row">
          ${span ? spanColorDot(poleId, span.spanId) : ""}
          <span>${span ? `${poleLink(span.fromPole)} → ${poleLink(span.toPole)}` : escapeHtml(sc.spanId || "")}</span>
          ${isReferenceSpan ? `<em>REF</em>` : ""}
          <button class="inline-icon-action danger-action" type="button"
            data-delete-comm-span
            data-pole="${escapeHtml(sc.poleId)}"
            data-span="${escapeHtml(sc.spanId)}"
            data-owner="${escapeHtml(sc.owner)}"
            data-wire-id="${escapeHtml(sc.wireId || "")}"
            title="Delete only this span"
            aria-label="Delete only this span">&#10005;</button>
        </div>`,
        serviceDropHtml: `<div class="comm-midspan-value">
          <input type="checkbox"
            data-scope="spanComm"
            data-pole="${escapeHtml(sc.poleId)}"
            data-span="${escapeHtml(sc.spanId)}"
            data-owner="${escapeHtml(sc.owner)}"
            data-wire-id="${escapeHtml(sc.wireId || "")}"
            data-field="serviceDrop"
            ${sc.serviceDrop ? "checked" : ""}>
        </div>`,
        downGuyHtml: `<div class="comm-midspan-value">
          <input type="checkbox"
            data-scope="spanComm"
            data-pole="${escapeHtml(sc.poleId)}"
            data-span="${escapeHtml(sc.spanId)}"
            data-owner="${escapeHtml(sc.owner)}"
            data-wire-id="${escapeHtml(sc.wireId || "")}"
            data-field="downGuy"
            ${sc.downGuy ? "checked" : ""}>
        </div>`,
        resagHtml: `<div class="comm-midspan-value">
          <input type="checkbox"
            data-scope="spanComm"
            data-pole="${escapeHtml(sc.poleId)}"
            data-span="${escapeHtml(sc.spanId)}"
            data-owner="${escapeHtml(sc.owner)}"
            data-wire-id="${escapeHtml(sc.wireId || "")}"
            data-field="resagServiceDrop"
            title="Re-sag this service drop to 15'6&quot; at midspan"
            aria-label="Re-sag this service drop to 15 feet 6 inches at midspan"
            ${sc.resagServiceDrop ? "checked" : ""}
            ${canResagServiceDrop ? "" : "disabled"}>
        </div>`,
        midspanHtml: `<div class="comm-midspan-value colored-midspan ${span ? spanColorClass(poleId, span.spanId) : ""}">${canEditMidspan
          ? `<input class="input height-input remote-height-input midspan-highlight-input" data-scope="spanComm" data-pole="${escapeHtml(sc.poleId)}" data-span="${escapeHtml(sc.spanId)}" data-owner="${escapeHtml(sc.owner)}" data-wire-id="${escapeHtml(sc.wireId || "")}" data-field="midspan" value="${escapeHtml(midspan)}" placeholder="">`
          : `<strong class="midspan-highlight-display">${escapeHtml(midspan)}</strong>`}
          ${hasStoredMidspan ? `<button class="inline-icon-action danger-action" type="button"
            data-clear-comm-midspan
            data-pole="${escapeHtml(sc.poleId)}"
            data-span="${escapeHtml(sc.spanId)}"
            data-owner="${escapeHtml(sc.owner)}"
            data-wire-id="${escapeHtml(sc.wireId || "")}"
            title="Clear only this midspan"
            aria-label="Clear only this midspan">&#10005;</button>` : ""}
          ${!hasStoredMidspan ? `<span class="inline-icon-action ghost-action" aria-hidden="true"></span>` : ""}
        </div>`,
        maxHeightAtMSHtml: `<div class="comm-midspan-value"><strong>${(!isReferenceSpan || calculatedBackspan) && hasMidspan ? escapeHtml(span?.midspanMaxCommHeight || "") : ""}</strong></div>`,
        remoteHtml: (isReferenceSpan && !calculatedBackspan) || !hasMidspan || !remote
          ? `<div class="comm-midspan-value"><strong></strong></div>`
          : `<div class="comm-midspan-value"><input class="input height-input remote-height-input" data-scope="spanComm" data-pole="${escapeHtml(remote.poleId)}" data-span="${escapeHtml(remote.spanId)}" data-owner="${escapeHtml(remote.owner)}" data-wire-id="${escapeHtml(remote.wireId || "")}" data-field="existingHOAChange" value="${escapeHtml(remote.existingHOAChange || remote.existingHOA || "")}"></div>`
      });
    });
    return entries;
  }

  function renderCommSpanRefs(group, poleId) {
    const entries = commMidspanEntries(group, poleId);
    return `<div class="comm-span-list">${entries.map(entry => entry.spanHtml).join("")}</div>`;
  }

  function renderCommServiceDropValues(group, poleId) {
    const entries = commMidspanEntries(group, poleId);
    return `<div class="comm-midspan-list">${entries.map(entry => entry.serviceDropHtml).join("")}</div>`;
  }

  function renderCommDownGuyValues(group, poleId) {
    const entries = commMidspanEntries(group, poleId);
    return `<div class="comm-midspan-list">${entries.map(entry => entry.downGuyHtml).join("")}</div>`;
  }

  function renderCommTransferValues(group, poleId) {
    // Transfer belongs to the physical comm at this pole, not to each span
    // touching it. Existing files may still have the flag on only one span,
    // so any checked relation makes the single grouped control active.
    const checked = group.rows.some(row => row.transferToNewPole);
    return `<div class="comm-group-checkbox">
      <input type="checkbox"
        data-scope="commGroup"
        data-pole="${escapeHtml(poleId)}"
        data-group-key="${escapeHtml(group.key)}"
        data-field="transferToNewPole"
        title="Replace raise/lower MR with one transfer to the new pole"
        aria-label="Transfer ${escapeHtml(group.owner)} to the new pole"
        ${checked ? "checked" : ""}>
    </div>`;
  }

  function renderCommResagValues(group, poleId) {
    const entries = commMidspanEntries(group, poleId);
    return `<div class="comm-midspan-list">${entries.map(entry => entry.resagHtml).join("")}</div>`;
  }

  function renderCommMidspanValues(group, poleId) {
    const entries = commMidspanEntries(group, poleId);
    return `<div class="comm-midspan-list">${entries.map(entry => entry.midspanHtml).join("")}</div>`;
  }

  function renderCommMaxHeightAtMSValues(group, poleId) {
    const entries = commMidspanEntries(group, poleId);
    return `<div class="comm-midspan-list">${entries.map(entry => entry.maxHeightAtMSHtml).join("")}</div>`;
  }

  function renderCommRemoteValues(group) {
    const entries = commMidspanEntries(group, group.rows[0]?.poleId || "");
    return `<div class="comm-midspan-list">${entries.map(entry => entry.remoteHtml).join("")}</div>`;
  }

  function renderCommFlagging(group) {
    const messages = Array.from(new Set(group.rows
      .map(row => String(row.flaggingMessage || "").trim())
      .filter(message => message && message !== "OK")));
    const hasProblem = group.rows.some(row => row.flaggingStatus === "PROBLEM");
    const hasMissing = group.rows.some(row => row.flaggingStatus === "MISSING" || row.flaggingStatus === "MISSING_POWER");
    if (!messages.length && !hasProblem && !hasMissing) return `<span class="badge changed">OK</span>`;
    const badgeClass = hasProblem ? "danger" : "warning";
    const label = hasProblem ? "Clearance Issue" : "Missing Data";
    return `<div class="flagging-cell">
      <span class="badge ${badgeClass}">${label}</span>
      <div class="flagging-message">${messages.map(escapeHtml).join(" ")}</div>
    </div>`;
  }

  function poleIdsForSpan(spanId) {
    const span = S.getSpan(spanId);
    return span ? [span.fromPole, span.toPole].filter(Boolean) : [];
  }

  function affectedPoleIdsForElement(el, scope) {
    const ids = new Set();
    if (el?.dataset?.pole) ids.add(el.dataset.pole);
    if (el?.dataset?.span) poleIdsForSpan(el.dataset.span).forEach(id => ids.add(id));
    if (scope === "pole" && el?.dataset?.pole) {
      S.getConnectedSpans(el.dataset.pole).forEach(span => {
        if (span.fromPole) ids.add(span.fromPole);
        if (span.toPole) ids.add(span.toPole);
      });
    }
    if (scope === "spanSide" && el?.dataset?.pole && el?.dataset?.field === "proposedHOA") {
      // A proposed value on the current pole can feed Next Pole Proposed for
      // spans arriving from previous poles, even when the spanId is different.
      S.getConnectedSpans(el.dataset.pole).forEach(span => {
        if (span.fromPole) ids.add(span.fromPole);
        if (span.toPole) ids.add(span.toPole);
      });
    }
    return Array.from(ids).filter(Boolean);
  }

  function bindScrollLinks(root) {
    if (!root) return;
    root.querySelectorAll("[data-scroll-to-pole]").forEach(btn => {
      btn.addEventListener("click", () => selectPole(btn.dataset.scrollToPole));
    });
  }

  function replacePoleCard(poleId) {
    if (!els.polesOverview || !poleId || !S.getPole(poleId)) return false;
    const oldCard = els.polesOverview.querySelector(`[data-pole-card="${CSS.escape(poleId)}"]`);
    if (!oldCard || !oldCard.isConnected || !oldCard.parentNode) return false;
    // Replacing a card after a checkbox change used to reset every wide table
    // to its left edge. Keep each table viewport so controls on the right stay
    // under the pointer while the recalculated card is rendered.
    const tableScrollPositions = Array.from(oldCard.querySelectorAll(".table-wrap"))
      .map(wrapper => ({ left: wrapper.scrollLeft, top: wrapper.scrollTop }));
    const template = document.createElement("template");
    template.innerHTML = renderPoleWorkspace(poleId).trim();
    const nextCard = template.content.firstElementChild;
    if (!nextCard) return false;
    const parent = oldCard.parentNode;
    if (!parent || !parent.contains(oldCard)) return false;
    parent.replaceChild(nextCard, oldCard);
    Array.from(nextCard.querySelectorAll(".table-wrap")).forEach((wrapper, index) => {
      const position = tableScrollPositions[index];
      if (!position) return;
      wrapper.scrollLeft = position.left;
      wrapper.scrollTop = position.top;
    });
    wireEditableEvents(nextCard);
    bindScrollLinks(nextCard);
    bindLocalActions(nextCard);
    return true;
  }

  function renderAffectedPoles(poleIds) {
    const ids = Array.from(new Set((poleIds || []).filter(Boolean)));
    if (!ids.length || (S.getState().ui.filter || "all") !== "all") {
      render();
      return;
    }

    global.Calculations.recalculateAll();
    renderSummary();
    renderPoleLists();

    let replacedAny = false;
    ids.forEach(id => {
      replacedAny = replacePoleCard(id) || replacedAny;
    });

    if (!replacedAny) renderAllPolesWorkspace();
  }

  function noteList(lines) {
    const clean = (Array.isArray(lines) ? lines : [lines])
      .map(line => String(line || "").trim())
      .filter(Boolean);
    if (!clean.length) return "";
    return `<div class="auto-notes">${clean.map(line => `<div>${escapeHtml(line)}</div>`).join("")}</div>`;
  }

  function renderEditableNotes(scope, attrs, value, autoNotes, placeholder = "Notes...") {
    const attrText = Object.entries(attrs || {})
      .map(([name, attrValue]) => `data-${name}="${escapeHtml(attrValue)}"`)
      .join(" ");
    return `<div class="notes-cell">
      <textarea class="input text-input" data-scope="${escapeHtml(scope)}" ${attrText} data-field="notes" placeholder="${escapeHtml(placeholder)}">${escapeHtml(value || "")}</textarea>
    </div>`;
  }

  function commMidspanNote(row, span) {
    if (!span || !span.midspanLowPower) return "Missing Power MS to calculate Max MS Comm.";
    const midspan = typeof row === "string" ? row : displayMidspan(row);
    if (!midspan) return span.midspanMaxCommHeight ? `Missing midspan. Max ${span.midspanMaxCommHeight}.` : "Missing midspan.";
    return row?.clearanceMSMessage || `Max ${span.midspanMaxCommHeight} · Low Power MS ${span.midspanLowPower}.`;
  }

  function poleClearanceNote(height, pole, missingLabel = "Missing Data") {
    if (!pole || !pole.lowPower) return missingLabel;
    const max = pole.maxCommHeight || "";
    if (!height) return max ? `Missing data. Max ${max}.` : "Missing data.";
    const aboveMax = max && H.compareHeights(height, max) === 1;
    return aboveMax
      ? `Clearance issue. Max ${max} · Low Power ${pole.lowPower}.`
      : `OK. Max ${max} · Low Power ${pole.lowPower}.`;
  }

  function spanSideClearanceNote(side) {
    if (!side.ocalcMS && !side.proposedMidspan) return "Missing O-CALC MS.";
    const settings = S.getState().settings || {};
    if (side.clearanceMSReason === "LOW_POWER" && side.clearanceMSIssue && settings.allowLowPowerMidspanAdjustment !== false) return `Ensure min 30" to low power at midspan.`;
    if (side.clearanceMSStatus === "PENDING") {
      return `${side.clearanceMSMessage || "Clearance issue."} Adjusting to ${side.pendingMidspanFinal || ""}...`;
    }
    return side.clearanceMSMessage || "";
  }

  function renderMidspanClearanceStatus(row, span) {
    const status = typeof row === "string" ? "" : row?.clearanceMSStatus || "";
    const midspan = typeof row === "string" ? row : displayMidspan(row);
    if (!span || !span.midspanLowPower) return `<span class="badge warning">Missing Power MS</span>`;
    if (!midspan) return `<span class="badge warning">Missing Data</span>`;
    if (status === "PROBLEM") return `<span class="badge danger">Clearance Issue</span>`;
    if (status === "MISSING_POWER") return `<span class="badge warning">Missing Power MS</span>`;
    return `<span class="badge changed">OK</span>`;
  }

  function renderSpanSideMidspanStatus(side) {
    const status = side.clearanceMSStatus || "";
    if (!side.ocalcMS && !side.proposedMidspan) return `<span class="badge warning">Missing Data</span>`;
    if (status === "PENDING") return `<span class="badge danger pulse-badge">Clearance Issue</span>`;
    if (status === "PROBLEM") return `<span class="badge danger">Clearance Issue</span>`;
    if (status === "ADJUSTED") return `<span class="badge warning">Adjusted</span>`;
    if (status === "ADJUSTMENT_NEEDED") return `<span class="badge warning">Adjustment Needed</span>`;
    return `<span class="badge changed">OK</span>`;
  }

  function renderSpanSideFlagging(side) {
    if (!side.proposedFlaggingMessage || side.proposedFlaggingStatus === "OK") return `<span class="badge changed">OK</span>`;
    return `<div class="flagging-cell">
      <span class="badge danger">Clearance Issue</span>
      <div class="flagging-message">${escapeHtml(side.proposedFlaggingMessage)}</div>
    </div>`;
  }

  function renderClearanceStatus(height, pole, missingLabel = "Missing Data") {
    if (!pole || !pole.lowPower) return `<span class="badge warning">${missingLabel}</span>`;
    const max = pole.maxCommHeight || "";
    if (!height) return `<span class="badge warning">Missing Data</span>`;
    const aboveMax = max && H.compareHeights(height, max) === 1;
    if (aboveMax) return `<span class="badge danger">Clearance Issue</span>`;
    return `<span class="badge changed">OK</span>`;
  }

  function renderClearanceSettings() {
    if (!els.clearanceSettings) return;
    const settings = S.getState().settings || {};
    const profiles = global.ProjectProfiles?.PROFILES || {};
    const selectedProfile = settings.projectProfile || "INTEC";
    const projectOptions = Object.values(profiles).map(profile =>
      `<option value="${escapeHtml(profile.id)}" ${selectedProfile === profile.id ? "selected" : ""}>${escapeHtml(profile.label)}</option>`
    ).join("");
    const clearanceRows = [
      ["polePowerCommsClearance", "Pole · Power-comms", settings.polePowerCommsClearance || settings.clearanceToPower || "40\""],
      ["commClearance", "Pole · Comm-comm", settings.commClearance || "12\""],
      ["boltClearance", "Pole · Bolt-bolt", settings.boltClearance || "4\""],
      ["midspanPowerCommClearance", "Midspan · Power-comm", settings.midspanPowerCommClearance || "30\""],
      ["midspanCommCommClearance", "Midspan · Comm-comm", settings.midspanCommCommClearance || "4\""]
    ];
    if (selectedProfile === "METRONET" && String(settings.proposedOwner || "MidAm").toUpperCase() === "MIDAM") {
      clearanceRows.push(
        ["streetlightBracketCommClearance", "Pole · Streetlight bracket-comm", settings.streetlightBracketCommClearance || "20\""],
        ["streetlightDripLoopCommClearance", "Pole · Streetlight drip loop-comm", settings.streetlightDripLoopCommClearance || "12\""]
      );
    }
    const position = settings.position === "LOW_COMM" ? "LOW_COMM" : "TOP_COMM";
    const proposedOwner = settings.proposedOwner || "Wecom";
    const fiberSizes = settings.fiberSizes && typeof settings.fiberSizes === "object" ? settings.fiberSizes : {};
    const detectedFibers = new Set(Object.keys(fiberSizes));
    (S.getState().makeReadyReferences || []).forEach(ref => {
      const source = `${ref.attachmentFiber || ""} ${ref.attachmentSizeRaw || ""}`;
      const match = source.match(/\b(\d+)\s*CT\b/i);
      if (match) detectedFibers.add(`${match[1]}CT Fiber`);
    });
    const fiberTypes = Array.from(detectedFibers).sort((a, b) => {
      const aCount = Number((a.match(/\d+/) || [0])[0]);
      const bCount = Number((b.match(/\d+/) || [0])[0]);
      return aCount - bCount || a.localeCompare(b);
    });
    const ownerOptions = ["Wecom", "MidAm", "CenturyLink", "Cable One", "Cox", "Fatbeam", "Vexus", "MCI Metro"].map(owner =>
      `<option value="${escapeHtml(owner)}" ${proposedOwner === owner ? "selected" : ""}>${escapeHtml(owner)}</option>`
    ).join("");
    const renderRow = ([field, label, value]) => `
        <label class="clearance-row">
          <span>${escapeHtml(label)}</span>
          <input class="input height-input" data-scope="settings" data-field="${escapeHtml(field)}" value="${escapeHtml(value)}" />
        </label>
      `;
    els.clearanceSettings.innerHTML = `
      <div class="settings-section">
        <h3>Clearance</h3>
        <div class="settings-grid clearance-grid">${clearanceRows.map(renderRow).join("")}</div>
      </div>
      <div class="settings-section settings-section-adjustments">
        <h3>Adjustments</h3>
        <div class="settings-grid adjustments-grid">
          <label class="clearance-row position-row">
            <span>Project</span>
            <select class="input position-select" data-scope="settings" data-field="projectProfile">
              ${projectOptions}
            </select>
          </label>
          <label class="clearance-row position-row">
            <span>Position</span>
            <select class="input position-select" data-scope="settings" data-field="position">
              <option value="TOP_COMM" ${position === "TOP_COMM" ? "selected" : ""}>Top Comm</option>
              <option value="LOW_COMM" ${position === "LOW_COMM" ? "selected" : ""}>Low Comm</option>
            </select>
          </label>
          <label class="clearance-row position-row">
            <span>Make Ready</span>
            <select class="input position-select" data-scope="settings" data-field="mrCase">
              <option value="LOWER" ${(settings.mrCase || "LOWER") === "LOWER" ? "selected" : ""}>Lowercase</option>
              <option value="UPPER" ${settings.mrCase === "UPPER" ? "selected" : ""}>Uppercase</option>
            </select>
          </label>
          ${selectedProfile === "METRONET" ? `<label class="clearance-row position-row">
            <span>WI</span>
            <select class="input position-select" data-scope="settings" data-field="proposedOwner">
              <option value="MidAm" selected>MidAm</option>
            </select>
          </label>` : settings.hideProposedOwner ? "" : `<label class="clearance-row position-row">
            <span>Proposed Owner</span>
            <select class="input position-select" data-scope="settings" data-field="proposedOwner">
              ${ownerOptions}
            </select>
          </label>`}
        </div>
      </div>
      ${selectedProfile === "INTEC" && fiberTypes.length ? `<div class="settings-section settings-section-fiber">
        <h3>Fiber</h3>
        <div class="settings-grid fiber-grid">
          <label class="clearance-row">
            <span>Messenger Size</span>
            <input class="input" data-scope="attachmentSettings" data-field="attachmentMessengerSize" value="${escapeHtml(settings.attachmentMessengerSize || "")}" placeholder="0.25" />
          </label>
          ${fiberTypes.map(fiber => `<label class="clearance-row">
            <span>${escapeHtml(fiber)} Size</span>
            <input class="input" data-scope="attachmentSettings" data-field="fiberSize" data-fiber="${escapeHtml(fiber)}" value="${escapeHtml(fiberSizes[fiber] || "")}" placeholder="0.51" />
          </label>`).join("")}
        </div>
      </div>` : ""}
    `;
    wireEditableEvents(els.clearanceSettings);
  }

  function updateAutoCalculateButtonState() {
    if (!els.autoCalculateBtn) return;
    const hasPoleData = Object.keys(S.getState().poles || {}).length > 0;
    const isLowComm = (S.getState().settings?.position || "TOP_COMM") === "LOW_COMM";
    const disableAuto = !hasPoleData || isLowComm;
    els.autoCalculateBtn.disabled = disableAuto;
    els.autoCalculateBtn.classList.toggle("btn-disabled", disableAuto);
    els.autoCalculateBtn.classList.toggle("btn-primary", !disableAuto);
    els.autoCalculateBtn.textContent = !hasPoleData
      ? "Auto Calculate Moves · Import Data First"
      : isLowComm
        ? "Auto Calculate Moves · Top Comm Required"
        : "Auto Calculate Moves";

    if (els.exportProposedJsonBtn) {
      els.exportProposedJsonBtn.disabled = !hasPoleData;
      els.exportProposedJsonBtn.classList.toggle("btn-disabled", !hasPoleData);
      els.exportProposedJsonBtn.classList.toggle("btn-success", hasPoleData);
    }
    if (els.exportDebugJsonBtn) {
      els.exportDebugJsonBtn.disabled = !hasPoleData;
      els.exportDebugJsonBtn.classList.toggle("btn-disabled", !hasPoleData);
    }
  }

  function renderPoleClassResults() {
    if (!els.poleClassResults) return;
    const rows = S.getState().poleClassChecks || [];
    renderPoleClassTabWarning();
    if (!rows.length) {
      els.poleClassResults.innerHTML = `<div class="detail-placeholder">Import an Excel file with Collection data to review Height / Class.</div>`;
      return;
    }
    const heightIssueCount = rows.filter(row => poleClassSeverity(row) === "critical").length;
    const classIssueCount = rows.filter(row => poleClassSeverity(row) === "warning").length;
    const issueCount = rows.filter(row => row.status !== "OK").length;
    els.poleClassResults.innerHTML = `
      <div class="pole-class-summary">
        <span class="badge changed">OK ${rows.length - issueCount}</span>
        <span class="badge danger">Height ${heightIssueCount}</span>
        <span class="badge warning">Class ${classIssueCount}</span>
      </div>
      <div class="table-wrap"><table class="pole-class-table">
        <thead><tr>
          <th>Pole ID</th><th>Tip</th><th>Imported Circ.</th><th>Diameter</th><th>Calc Circ.</th><th>Imported Type</th><th>Calc Height</th><th>Calc Class</th><th>Expected Type</th><th>Status</th>
        </tr></thead>
        <tbody>${rows.map((row, index) => {
          const severity = poleClassSeverity(row);
          return `
          <tr class="${poleClassRowClass(severity)}">
            <td><strong>${poleLink(row.poleId)}</strong></td>
            <td>${escapeHtml(row.tip || "")}</td>
            <td>${escapeHtml(row.importedCircumference || "")}</td>
            <td><input class="input pole-class-diameter-input" data-scope="poleClassCheck" data-index="${index}" data-field="manualDiameter" value="${escapeHtml(row.manualDiameter || "")}" /></td>
            <td>${escapeHtml(row.circumference || "")}</td>
            <td>${escapeHtml(row.importedType || "")}</td>
            <td>${escapeHtml(row.calculatedHeight || "")}</td>
            <td>${escapeHtml(row.calculatedClass || "")}</td>
            <td>${escapeHtml(row.expectedType || "")}</td>
            <td>${renderPoleClassStatus(row, severity)}</td>
          </tr>
        `; }).join("")}</tbody>
      </table></div>`;
    els.poleClassResults.insertAdjacentHTML("beforeend", renderAnsiReferenceTable());
    wireEditableEvents(els.poleClassResults);
    bindScrollLinks(els.poleClassResults);
  }

  function poleClassSeverity(row) {
    const status = String(row?.status || "");
    if (status === "OK") return "ok";
    if (/Height mismatch|Missing Tip|No (?:ANSI|reference) height row/i.test(status)) return "critical";
    return "warning";
  }

  function reviewStatusBadge(status) {
    const normalized = String(status || "PASS").toUpperCase();
    const badgeClass = normalized === "ERROR" ? "danger"
      : normalized === "WARNING" || normalized === "NOT_READY" ? "warning"
        : normalized === "IGNORED" ? "ignored"
        : "changed";
    const label = normalized === "NOT_READY" ? "NOT READY" : normalized;
    return `<span class="badge ${badgeClass}">${label}</span>`;
  }

  function renderReviewCheck(item) {
    const canIgnore = item.poleId && item.status !== "PASS";
    return `<li class="excel-review-check ${String(item.status || "").toLowerCase()} ${item.level === "low" ? "low-level" : ""} ${item.ignored ? "ignored" : ""}">
      <div class="excel-review-check-heading">
        ${reviewStatusBadge(item.ignored ? "IGNORED" : item.status)}
        <strong>${escapeHtml(item.section || "Review")} · ${escapeHtml(item.title || item.code || "Check")}</strong>
        ${item.level === "low" ? `<span class="badge">Low</span>` : ""}
      </div>
      <p>${escapeHtml(item.message || "")}</p>
      ${item.expected || item.actual ? `<dl class="review-values">
        ${item.expected ? `<div><dt>Expected</dt><dd>${escapeHtml(item.expected)}</dd></div>` : ""}
        ${item.actual ? `<div><dt>Found</dt><dd>${escapeHtml(item.actual)}</dd></div>` : ""}
      </dl>` : ""}
      ${item.details?.length ? `<div class="review-source-details">${item.details.map(detail => (
        `<span>Span Index: ${escapeHtml(detail.spanIndex || "—")} · Span Id: ${escapeHtml(detail.spanId || "—")} · Type: ${escapeHtml(detail.type || "—")} · Linked: ${escapeHtml(detail.linkedCollectionTitle || "—")}</span>`
      )).join("")}</div>` : ""}
      ${canIgnore ? `<div class="review-check-actions"><button class="btn review-ignore-btn" type="button" data-review-ignore-key="${escapeHtml(item.ignoreKey)}" data-review-ignore-value="${item.ignored ? "false" : "true"}">${item.ignored ? "Restore" : "Ignore"}</button></div>` : ""}
    </li>`;
  }

  function renderExcelReviewTabWarning(review) {
    if (!els.excelReviewTabWarning || !els.excelReviewTabWarningCount) return;
    const summary = review?.summary || {};
    const activeCount = Number(summary.errors || 0) + Number(summary.warnings || 0);
    els.excelReviewTabWarningCount.textContent = String(activeCount);
    els.excelReviewTabWarning.classList.toggle("hidden", activeCount === 0);
    els.excelReviewTabWarning.title = activeCount === 1
      ? "1 pole has active Excel Review findings"
      : `${activeCount} poles have active Excel Review findings`;
  }

  function renderExcelReviewResults() {
    if (!els.excelReviewResults || !global.ExcelReview) return;
    const openPoleIds = new Set(Array.from(els.excelReviewResults.querySelectorAll("details[open][data-review-result-pole]"))
      .map(details => details.dataset.reviewResultPole));
    const review = global.ExcelReview.getReviewState();
    const summary = review.summary || {};
    renderExcelReviewTabWarning(review);
    if (els.excelReviewTimestamp) {
      els.excelReviewTimestamp.textContent = review.reviewedAt
        ? `Last reviewed ${new Date(review.reviewedAt).toLocaleString()}`
        : "No review has been run.";
    }
    if (els.rerunExcelReviewBtn) {
      const hasCollectionSource = Boolean(S.getState().excelReviewSource?.collection?.rows?.length);
      els.rerunExcelReviewBtn.disabled = !hasCollectionSource;
      els.rerunExcelReviewBtn.classList.toggle("btn-disabled", !hasCollectionSource);
    }
    if (!review.reviewedAt) {
      els.excelReviewResults.innerHTML = `<div class="detail-placeholder">Import a raw Excel file to prepare HOA and Final PLA / MR review results.</div>`;
      return;
    }

    const globalChecks = review.globalChecks || [];
    els.excelReviewResults.innerHTML = `
      <div class="excel-review-summary" aria-label="Excel Review summary">
        <div class="review-summary-item error"><strong>${Number(summary.errors || 0)}</strong><span>Errors</span></div>
        <div class="review-summary-item warning"><strong>${Number(summary.warnings || 0)}</strong><span>Warnings</span></div>
        <div class="review-summary-item pass"><strong>${Number(summary.passed || 0)}</strong><span>Passed</span></div>
      </div>
      ${globalChecks.length ? `<section class="excel-review-global">
        <h3>Workbook Issues</h3>
        <ul class="excel-review-checks">${globalChecks.map(renderReviewCheck).join("")}</ul>
      </section>` : ""}
      <div class="excel-review-poles">${(review.results || []).map(result => {
        const visibleChecks = result.checks.filter(item => item.status !== "PASS" || item.applicable === false || item.code === "ALL_APPLICABLE_CHECKS_PASSED");
        return `<details class="excel-review-pole review-${String(result.overallStatus || "pass").toLowerCase()}" data-review-result-pole="${escapeHtml(result.poleId)}">
          <summary>
            <span class="review-pole-identity"><button class="review-pole-link" type="button" data-review-pole="${escapeHtml(result.poleId)}" title="Open this pole in Calculator"><strong>${escapeHtml(result.poleId)}</strong></button>${result.sequence ? `<small>Sequence ${escapeHtml(result.sequence)}</small>` : ""}</span>
            <span class="review-phase-status"><small>HOA Review</small>${reviewStatusBadge(result.hoaStatus)}</span>
            <span class="review-phase-status"><small>Final Review</small>${reviewStatusBadge(result.finalStatus)}</span>
            <span class="review-phase-status overall"><small>Overall</small>${reviewStatusBadge(result.overallStatus)}</span>
          </summary>
          <ul class="excel-review-checks">${visibleChecks.map(renderReviewCheck).join("")}</ul>
        </details>`;
      }).join("") || `<div class="detail-placeholder">No Collection rows were available for review.</div>`}</div>`;
    els.excelReviewResults.querySelectorAll("details[data-review-result-pole]").forEach(details => {
      if (openPoleIds.has(details.dataset.reviewResultPole)) details.open = true;
    });
    els.excelReviewResults.querySelectorAll("[data-review-pole]").forEach(button => {
      button.addEventListener("click", event => {
        event.preventDefault();
        event.stopPropagation();
        const requested = button.dataset.reviewPole || "";
        const poleId = Object.keys(S.getState().poles || {}).find(id =>
          S.canonicalPoleIdentity(id) === S.canonicalPoleIdentity(requested)
        ) || requested;
        selectPole(poleId);
      });
    });
    els.excelReviewResults.querySelectorAll("[data-review-ignore-key]").forEach(button => {
      button.addEventListener("click", event => {
        event.preventDefault();
        event.stopPropagation();
        recordUndoSnapshot();
        global.ExcelReview.setCheckIgnored(button.dataset.reviewIgnoreKey, button.dataset.reviewIgnoreValue === "true");
        renderExcelReviewResults();
      });
    });
  }

  // Height discrepancies are operationally more important than class-only
  // differences. Keep this lookup independent from comm/proposed flagging so
  // the warning remains visible even when a pole is marked UG or PCO.
  function hasCriticalPoleHeightIssue(poleId) {
    const normalizedPoleId = String(poleId || "").trim().toUpperCase();
    return (S.getState().poleClassChecks || []).some(row => (
      String(row?.poleId || "").trim().toUpperCase() === normalizedPoleId
      && poleClassSeverity(row) === "critical"
    ));
  }

  function renderPoleClassTabWarning() {
    if (!els.poleClassTabWarning || !els.poleClassTabWarningCount) return;
    const criticalCount = (S.getState().poleClassChecks || [])
      .filter(row => poleClassSeverity(row) === "critical")
      .length;
    els.poleClassTabWarningCount.textContent = String(criticalCount);
    els.poleClassTabWarning.classList.toggle("hidden", criticalCount === 0);
    els.poleClassTabWarning.title = criticalCount === 1
      ? "1 pole has a critical height issue"
      : `${criticalCount} poles have critical height issues`;
  }

  function poleClassRowClass(severity) {
    if (severity === "ok") return "changed-row";
    if (severity === "critical") return "pole-class-critical-row";
    return "pole-class-warning-row";
  }

  function renderPoleClassStatus(row, severity) {
    if (severity === "ok") return `<span class="badge changed">OK</span>`;
    const label = severity === "critical" ? "Height Critical" : "Review Class";
    const badgeClass = severity === "critical" ? "danger" : "warning";
    return `<div class="flagging-cell">
      <span class="badge ${badgeClass}">${label}</span>
      <div class="flagging-message">${escapeHtml(row.status)}</div>
    </div>`;
  }

  function renderAnsiReferenceTable() {
    const classes = global.ExcelImport?.ANSI_POLE_CLASSES || [];
    const table = global.ExcelImport?.ANSI_CLASS_TABLE || {};
    const groundline = global.ExcelImport?.ANSI_APPROX_GROUNDLINE_DISTANCE || {};
    const heights = Object.keys(table).map(Number).sort((a, b) => a - b);
    return `<details class="reference-table-panel" open>
      <summary>Pole Height and Class Reference Table</summary>
      <div class="table-wrap"><table class="ansi-reference-table">
        <thead>
          <tr><th>Length of Pole (ft)</th><th>Approx. Groundline Distance from Butt (ft)</th>${classes.map(item => `<th>${escapeHtml(item)}</th>`).join("")}</tr>
        </thead>
        <tbody>${heights.map(height => `
          <tr>
            <td><strong>${height}</strong></td>
            <td>${escapeHtml(groundline[height] ?? "-")}</td>
            ${(table[height] || []).map(value => `<td>${value === null || value === undefined ? "-" : escapeHtml(value)}</td>`).join("")}
          </tr>
        `).join("")}</tbody>
      </table></div>
    </details>`;
  }

  function renderActiveView() {
    const activeView = S.getState().ui.activeView || "calculator";
    const isPoleClassView = activeView === "poleClass";
    document.querySelectorAll("[data-view-tab]").forEach(btn => {
      btn.classList.toggle("active", btn.dataset.viewTab === activeView);
    });
    document.querySelectorAll("[data-view-panel]").forEach(panel => {
      panel.classList.toggle("hidden", panel.dataset.viewPanel !== activeView);
    });
    document.body.classList.toggle("pole-class-view", isPoleClassView);
    document.body.classList.toggle("excel-review-view", activeView === "excelReview");
    if (activeView !== "calculator") setPoleIndexOpen(false);
    updatePoleIndexToggleVisibility();
  }

  function renderEnvironmentOptions(selected) {
    const current = selected || "NONE";
    return (S.ENVIRONMENT_OPTIONS || []).map(option => (
      `<option value="${escapeHtml(option.value)}" ${option.value === current ? "selected" : ""}>${escapeHtml(option.label)}</option>`
    )).join("");
  }

  function poleSummary(poleId) {
    const state = S.getState();
    const pole = S.getPole(poleId) || S.createPole({ poleId, isGenerated: /^Unknown-/i.test(poleId) });
    const spans = S.getConnectedSpans(poleId);
    const warnings = state.warnings.filter(w => w.poleId === poleId);
    const midspanCount = S.getSpanCommsForPole(poleId).filter(sc => sc.calculatedMidspan || sc.midspan || sc.ocalcMS).length;
    return {
      pole,
      spans,
      warnings,
      midspanCount,
      hasChanges: S.poleHasChanges(poleId)
    };
  }

  function poleFlaggingSummary(poleId) {
    const pole = S.getPole(poleId);
    const commIssues = S.getSpanCommsForPole(poleId)
      .filter(sc => sc.flaggingStatus === "PROBLEM")
      .length;
    const proposedIssues = S.getSpanSidesForPole(poleId)
      .filter(side => side.proposedFlaggingStatus === "PROBLEM")
      .length;
    const standaloneIssue = pole?.standaloneProposedHOA && global.Calculations.evaluateSpanSideFlagging({
      spanId: "",
      poleId,
      proposedHOA: pole.standaloneProposedHOA
    }).status === "PROBLEM" ? 1 : 0;
    const calculationIssueCount = commIssues + proposedIssues + standaloneIssue;
    const heightCritical = hasCriticalPoleHeightIssue(poleId);
    return {
      calculationIssueCount,
      heightCritical,
      issueCount: calculationIssueCount + (heightCritical ? 1 : 0),
      resolution: pole?.ugActive ? "UG" : pole?.pcoActive ? "PCO" : ""
    };
  }

  function renderSummary() {
    const state = S.getState();
    if (els.projectMeta) {
      els.projectMeta.textContent = state.importedFileName
        ? `${state.importedFileName} · ${new Date(state.importedAt || Date.now()).toLocaleString()}`
        : "No file imported";
    }
  }

  function filteredPoles() {
    const state = S.getState();
    const search = (state.ui.search || "").toLowerCase();
    const filter = state.ui.filter || "all";
    return Object.keys(state.poles).filter(poleId => {
      const summary = poleSummary(poleId);
      const pole = summary.pole;
      const comms = Array.isArray(pole.comms) ? pole.comms : [];
      const text = `${poleId} ${pole.sequence || ""} ${pole.poleHeight || ""} ${pole.lowPower || ""} ${pole.maxCommHeight || ""} ${comms.map(c => `${c.owner || ""} ${c.existingHOA || ""}`).join(" ")}`.toLowerCase();
      if (search && !text.includes(search)) return false;
      if (filter === "warnings" && summary.warnings.length === 0 && poleFlaggingSummary(poleId).issueCount === 0) return false;
      if (filter === "changed" && !summary.hasChanges) return false;
      return true;
    }).sort((a, b) => a.localeCompare(b, undefined, { numeric: true }));
  }

  function renderPoleListItem(poleId) {
    const state = S.getState();
    const { pole, warnings, hasChanges } = poleSummary(poleId);
    const flagging = poleFlaggingSummary(poleId);
    const active = state.selectedPoleId === poleId ? " active" : "";
    return `<button class="pole-index-link${active}" data-pole-select="${escapeHtml(poleId)}" type="button">
      <span>${escapeHtml(poleId)}</span>
      ${pole.isGenerated ? `<span class="mini-dot warning">Gen</span>` : ""}
      ${flagging.resolution ? `<span class="mini-dot ${flagging.resolution.toLowerCase()}">${flagging.resolution}</span>` : ""}
      ${!flagging.resolution && flagging.calculationIssueCount ? `<span class="mini-dot danger">Flag ${flagging.calculationIssueCount}</span>` : ""}
      ${flagging.heightCritical ? `<span class="mini-dot danger" title="Critical pole height issue">&#9888; Height</span>` : ""}
      ${hasChanges ? `<span class="mini-dot changed">Changed</span>` : ""}
    </button>`;
  }

  function renderPoleLists() {
    const poleIds = filteredPoles();
    els.poleSearchInput.value = S.getState().ui.search || "";
    els.warningFilterSelect.value = S.getState().ui.filter || "all";
    els.poleSearchInputTop.value = S.getState().ui.search || "";
    els.warningFilterSelectTop.value = S.getState().ui.filter || "all";
    [els.polesList, els.polesListTop].forEach(list => {
      list.innerHTML = poleIds.map(id => renderPoleListItem(id)).join("") || `<div class="detail-placeholder">No poles match that filter.</div>`;
      list.querySelectorAll("[data-pole-select]").forEach(btn => {
        btn.addEventListener("click", () => selectPole(btn.dataset.poleSelect));
      });
    });
  }

  function renderMRText(poleId) {
    const item = S.getState().mr.find(mr => mr.poleId === poleId);
    if (!item) return `<p class="muted">No Make Ready generated yet.</p>`;
    return `<pre class="mr-output">${escapeHtml(item.text)}</pre>`;
  }

  function renderPoleActions(poleId) {
    const pole = S.getPole(poleId);
    return `<div class="pole-action-buttons">
      <button class="mini-btn ${pole?.ugActive ? "active-action" : ""}" type="button" data-toggle-ug data-pole="${escapeHtml(poleId)}">UG</button>
      <button class="mini-btn ${pole?.pcoActive ? "active-action" : ""}" type="button" data-toggle-pco data-pole="${escapeHtml(poleId)}">PCO</button>
    </div>`;
  }

  function renderPoleEditableHeader(poleId) {
    const { pole, spans, hasChanges } = poleSummary(poleId);
    const flagging = poleFlaggingSummary(poleId);
    return `<div class="pole-workspace-header">
      <div>
        <h3 id="pole-${escapeHtml(poleId)}">${escapeHtml(poleId)}</h3>
        <div class="pole-meta">
          ${pole.isGenerated ? `<span class="badge warning">Editable generated other pole</span>` : ""}
          <span class="badge">Spans ${spans.length}</span>
          <span class="badge owner">Comms ${S.getSpanCommsForPole(poleId).length}</span>
          ${flagging.resolution ? `<span class="badge ${flagging.resolution.toLowerCase()}">${flagging.resolution}</span>` : ""}
          ${!flagging.resolution && flagging.calculationIssueCount ? `<span class="badge danger">Flagging ${flagging.calculationIssueCount}</span>` : ""}
          ${flagging.heightCritical ? `<span class="badge danger" title="Critical pole height issue">&#9888; Height Critical</span>` : ""}
          ${hasChanges ? `<span class="badge changed">Changed</span>` : ""}
        </div>
      </div>
      <div class="pole-kpis two-up">
        <label>Low Power on Pole<input class="input height-input" data-scope="pole" data-pole="${escapeHtml(poleId)}" data-field="lowPower" value="${escapeHtml(pole.lowPower || "")}" placeholder="30'8&quot;"></label>
        <label>Max Height on Pole<input class="input height-input muted-input" value="${escapeHtml(pole.maxCommHeight || "")}" readonly></label>
      </div>
    </div>`;
  }

  function renderSpanProposedTable(poleId) {
    const spans = proposedSpansForPole(poleId);
    const pole = S.getPole(poleId);
    const showStandalone = !spans.length || Boolean(pole?.standaloneProposedHOA);
    const standaloneFlagging = global.Calculations.evaluateSpanSideFlagging({
      spanId: "",
      poleId,
      proposedHOA: pole?.standaloneProposedHOA || ""
    });
    return `<div class="table-wrap"><table class="span-proposed-table wide-table">
      <thead><tr>
        <th>Span</th><th>Proposed</th><th>End Drop</th><th>Next Pole Proposed</th><th>O-CALC MS</th><th>MS Proposed</th><th>Max Height at MS</th><th>Adjusted Final MS</th><th>MS Flagging</th><th>Proposed Flagging</th><th>Environment</th><th>Environment Clearance</th><th>Notes</th><th>Actions</th>
      </tr></thead>
      <tbody>${spans.map(span => {
        const side = S.getSpanSide(span.spanId, poleId) || S.upsertSpanSide({ spanId: span.spanId, poleId });
        const physicalSpan = span.sourceSpanId ? S.getSpan(span.sourceSpanId) || span : span;
        const aboveMax = side.proposedHOA && H.compareHeights(side.proposedHOA, side.maxCommHeight || pole?.maxCommHeight) === 1;
        const midspanIssue = side.clearanceMSStatus === "PENDING" || side.clearanceMSStatus === "PROBLEM";
        const boltIssue = global.Calculations.evaluateProposedPoleClearance(side);
        const proposedFlaggingIssue = side.proposedFlaggingStatus === "PROBLEM";
        const rowClasses = [
          spanRowClasses(poleId, physicalSpan.spanId),
          side.proposedHOA || side.ocalcMS || side.msProposed || side.finalMidspan || side.proposedMidspan || side.endDrop ? "changed-row" : "",
          aboveMax || midspanIssue || !boltIssue.ok || proposedFlaggingIssue ? "warning-row" : ""
        ].filter(Boolean).join(" ");
        const autoNotes = [spanSideClearanceNote(side), boltIssue.message];
        return `<tr class="${rowClasses}">
          <td class="span-cell">
            <strong class="span-main-line">${spanChip(poleId, physicalSpan.spanId)}${poleLink(span.fromPole)} → ${poleLink(span.toPole)}</strong>
            ${spanLengthDisplay(physicalSpan) ? `<span class="span-distance-line">${escapeHtml(spanLengthDisplay(physicalSpan))}</span>` : ""}
          </td>
          <td><input class="input height-input" data-scope="spanSide" data-pole="${escapeHtml(poleId)}" data-span="${escapeHtml(span.spanId)}" data-field="proposedHOA" value="${escapeHtml(side.proposedHOA || "")}"></td>
          <td><span class="calculated-value">${escapeHtml(side.endDrop || "")}</span></td>
          <td><input class="input height-input" data-scope="spanSide" data-pole="${escapeHtml(poleId)}" data-span="${escapeHtml(span.spanId)}" data-field="proposedHOAChange" value="${escapeHtml(side.proposedHOAChange || "")}"></td>
          <td><input class="input decimal-height-input" data-scope="spanSide" data-pole="${escapeHtml(poleId)}" data-span="${escapeHtml(span.spanId)}" data-field="ocalcMS" value="${escapeHtml(displayDecimalFeetInput(side.ocalcMS, side.proposedMidspan))}" placeholder="XX.XX"></td>
          <td><span class="calculated-value">${escapeHtml(side.msProposed || "")}</span></td>
          <td>${escapeHtml(physicalSpan.midspanMaxCommHeight || "")}</td>
          <td><span class="calculated-value midspan-highlight-display ${spanColorClass(poleId, physicalSpan.spanId)}">${escapeHtml(side.finalMidspan || "")}</span></td>
          <td>${renderSpanSideMidspanStatus(side)}</td>
          <td>${renderSpanSideFlagging(side)}</td>
          <td><select class="input environment-input" data-scope="span" data-span="${escapeHtml(physicalSpan.spanId)}" data-field="environment">${renderEnvironmentOptions(physicalSpan.environment)}</select></td>
          <td><input class="input" data-scope="span" data-span="${escapeHtml(physicalSpan.spanId)}" data-field="environmentClearance" value="${escapeHtml(physicalSpan.environmentClearance || "")}"></td>
          <td>${renderEditableNotes("spanSide", { pole: poleId, span: span.spanId }, side.notes, autoNotes, "Own notes...")}</td>
          <td><button class="icon-action danger-action" type="button" data-delete-proposed-span data-pole="${escapeHtml(poleId)}" data-span="${escapeHtml(span.spanId)}" title="Delete proposed span" aria-label="Delete proposed span">&#10005;</button></td>
        </tr>`;
      }).join("")}${showStandalone ? `<tr class="standalone-proposed-row ${pole?.standaloneProposedHOA ? "changed-row" : ""} ${standaloneFlagging.status === "PROBLEM" ? "warning-row" : ""}">
          <td class="span-cell"></td>
          <td><input class="input height-input" data-scope="pole" data-pole="${escapeHtml(poleId)}" data-field="standaloneProposedHOA" value="${escapeHtml(pole?.standaloneProposedHOA || "")}" aria-label="Proposed attachment on terminal pole"></td>
          <td></td><td></td><td></td><td></td><td></td><td></td><td></td>
          <td>${renderSpanSideFlagging({ proposedFlaggingStatus: standaloneFlagging.status, proposedFlaggingMessage: standaloneFlagging.message })}</td>
          <td></td><td></td><td></td><td></td>
        </tr>` : ""}</tbody>
    </table></div>`;
  }

  function renderManualProposedSpanStarter(poleId, optionsConfig = {}) {
    const options = Object.keys(S.getState().poles || {})
      .filter(id => id !== poleId)
      .sort((a, b) => a.localeCompare(b, undefined, { numeric: true }))
      .map(id => `<option value="${escapeHtml(id)}">${escapeHtml(id)}</option>`)
      .join("");
    if (!options) return `<p class="muted">No other pole is available for a proposed span.</p>`;
    return `<div class="manual-proposed-starter ${optionsConfig.compact ? "compact-starter" : ""}">
      ${optionsConfig.compact ? "" : `<span class="muted">No spans with midspan are available to propose from this pole.</span>`}
      <label>
        <span>Propose to</span>
        <select class="input" data-manual-proposed-target="${escapeHtml(poleId)}">
          <option value="">Select pole</option>
          ${options}
        </select>
      </label>
      <button class="mini-btn" type="button" data-add-proposed-span data-pole="${escapeHtml(poleId)}">Add Proposed Span</button>
    </div>`;
  }

  function renderCommMovementTable(poleId) {
    const groups = groupedCommsForPole(poleId);
    if (!groups.length) return `<p class="muted">No comms imported from Span.Wire for this pole.</p>`;
    const settings = S.getState().settings || {};
    const showServiceDrop = settings.showServiceDrop !== false;
    const showResagServiceDrop = String(settings.projectProfile || "INTEC").toUpperCase() === "INTEC"
      && settings.showResagServiceDrop !== false;
    return `<div class="table-wrap"><table class="comm-movement-table">
      <thead><tr>
        <th>Owner/Comm</th><th>Existing HOA</th><th>HOA Change</th><th>Other Pole HOA</th><th>Span</th><th>Max Height at MS</th><th>Midspan</th><th>Flagging</th>${showServiceDrop ? "<th>Service Drop</th>" : ""}<th>DG</th><th>Transfer to New Pole</th>${showResagServiceDrop ? "<th>Re-sag Service Drop</th>" : ""}<th>Actions</th>
      </tr></thead>
      <tbody>${groups.map(group => {
        const pole = S.getPole(poleId);
        const effective = group.existingHOAChange || group.existingHOA;
        const aboveMax = effective && H.compareHeights(effective, pole?.maxCommHeight) === 1;
        const changed = Boolean(group.existingHOAChange || group.rows.some(row => (
          row.mr || row.transferToNewPole || row.resagServiceDrop
        )));
        const flaggingIssue = group.rows.some(row => row.flaggingStatus === "PROBLEM");
        const flaggingMissing = group.rows.some(row => row.flaggingStatus === "MISSING" || row.flaggingStatus === "MISSING_POWER");
        const rowClasses = [
          changed ? "changed-row" : "",
          aboveMax || flaggingIssue || flaggingMissing ? "warning-row" : ""
        ].filter(Boolean).join(" ");
        return `<tr class="${rowClasses}">
          <td><span class="badge owner">${escapeHtml(group.owner)}</span>${group.isPof ? ` <span class="badge pof">POF</span>` : ""}</td>
          <td>${escapeHtml(group.existingHOA || "")}</td>
          <td><input class="input height-input" data-scope="commGroup" data-pole="${escapeHtml(poleId)}" data-group-key="${escapeHtml(group.key)}" data-field="existingHOAChange" value="${escapeHtml(group.existingHOAChange || "")}"></td>
          <td>${renderCommRemoteValues(group)}</td>
          <td>${renderCommSpanRefs(group, poleId)}</td>
          <td>${renderCommMaxHeightAtMSValues(group, poleId)}</td>
          <td>${renderCommMidspanValues(group, poleId)}</td>
          <td>${renderCommFlagging(group)}</td>
          ${showServiceDrop ? `<td>${renderCommServiceDropValues(group, poleId)}</td>` : ""}
          <td>${renderCommDownGuyValues(group, poleId)}</td>
          <td>${renderCommTransferValues(group, poleId)}</td>
          ${showResagServiceDrop ? `<td>${renderCommResagValues(group, poleId)}</td>` : ""}
          <td><div class="row-actions">
            <button class="icon-action" type="button" data-edit-comm data-pole="${escapeHtml(poleId)}" data-group-key="${escapeHtml(group.key)}" title="Edit comm" aria-label="Edit comm">&#9998;</button>
            <button class="icon-action" type="button" data-edit-comm-spans data-pole="${escapeHtml(poleId)}" data-group-key="${escapeHtml(group.key)}" title="Edit comm spans" aria-label="Edit comm spans">&#8644;</button>
            <button class="icon-action danger-action" type="button" data-delete-comm data-pole="${escapeHtml(poleId)}" data-group-key="${escapeHtml(group.key)}" title="Delete full comm" aria-label="Delete full comm">&#10005;</button>
          </div></td>
        </tr>`;
      }).join("")}</tbody>
    </table></div>`;
  }

  function renderPowerTable(poleId) {
    const rows = S.getSpanPowerForPole(poleId).filter(row => H.parseHeight(row.midspan) !== null);
    if (!rows.length) return `<p class="muted">No power wires were imported for this pole.</p>`;
    return `<div class="table-wrap"><table class="power-table">
      <thead><tr><th>Span</th><th>Type</th><th>Attachment Height</th><th>Midspan</th></tr></thead>
      <tbody>${rows.map(row => {
        const span = S.getSpan(row.spanId);
        return `<tr>
          <td class="span-cell">${span ? `${poleLink(span.fromPole)} → ${poleLink(span.toPole)}` : ""}</td>
          <td><span class="badge warning">${escapeHtml(row.label)}</span></td>
          <td>${escapeHtml(row.attachmentHeight)}</td>
          <td><input class="input height-input" data-scope="spanPower" data-power-key="${escapeHtml(row.key || "")}" data-span="${escapeHtml(row.spanId)}" data-field="midspan" value="${escapeHtml(row.midspan || "")}"></td>
        </tr>`;
      }).join("")}</tbody>
    </table></div>`;
  }

  function renderPoleWorkspace(poleId) {
    const pole = S.getPole(poleId);
    return `<article class="pole-workspace-card ${pole?.ugActive ? "ug-active" : ""} ${pole?.pcoActive ? "pco-active" : ""}" data-pole-card="${escapeHtml(poleId)}">
      ${renderPoleEditableHeader(poleId)}
      <div class="workspace-grid">
        <section class="subsection wide" id="spans-${escapeHtml(poleId)}">
          <div class="subsection-title-row">
            <h4>Proposed by Span</h4>
            ${renderManualProposedSpanStarter(poleId, { compact: true })}
          </div>
          ${renderSpanProposedTable(poleId)}
        </section>
        <section class="subsection wide" id="comms-${escapeHtml(poleId)}">
          <div class="subsection-title-row">
            <h4>Existing Comm Movements</h4>
            <button class="mini-btn" type="button" data-add-comm data-pole="${escapeHtml(poleId)}">Add Comm</button>
          </div>
          <p class="muted">Move each existing comm to a new height. When the other pole on the same span changes, the calculated Midspan updates.</p>
          ${renderCommMovementTable(poleId)}
        </section>
        <section class="subsection wide">
          <h4>Imported Power / Clearance</h4>
          ${renderPowerTable(poleId)}
        </section>
        <section class="subsection make-ready-section" id="warnings-${escapeHtml(poleId)}">
          <div class="subsection-title-row">
            <h4>Make Ready</h4>
            <button class="mini-btn" type="button" data-copy-mr data-pole="${escapeHtml(poleId)}">Copy</button>
          </div>
          ${renderMRText(poleId)}
        </section>
        <section class="subsection">
          <h4>Pole Actions</h4>
          ${renderPoleActions(poleId)}
        </section>
      </div>
    </article>`;
  }

  function renderAllPolesWorkspace() {
    const poleIds = filteredPoles();
    els.polesOverview.innerHTML = poleIds.map(id => {
      try {
        return renderPoleWorkspace(id);
      } catch (error) {
        console.error(`Error rendering pole ${id}`, error);
        return `<article class="pole-workspace-card warning-row" data-pole-card="${escapeHtml(id)}">
          <div class="pole-workspace-header">
            <div>
              <h3>${escapeHtml(id)}</h3>
              <span class="badge danger">Data Error</span>
            </div>
          </div>
          <p class="muted">This pole was imported, but it has incomplete or unexpected data. Review the Excel file or import again.</p>
        </article>`;
      }
    }).join("") || `<div class="detail-placeholder">No poles to show.</div>`;
    wireEditableEvents(els.polesOverview);
    bindScrollLinks(els.polesOverview);
    bindLocalActions(els.polesOverview);
  }

  function wireEditableEvents(root) {
    root.querySelectorAll("input[data-scope], textarea[data-scope], select[data-scope]").forEach(input => {
      input.addEventListener("input", handleEditableInput);
      input.addEventListener("change", handleEditableChange);
      input.addEventListener("blur", handleEditableBlur);
      if (input.tagName !== "TEXTAREA") {
        input.addEventListener("keydown", event => {
          if (event.key === "Enter") {
            event.preventDefault();
            handleEditableBlur(event);
            handleEditableChange(event);
          }
        });
      }
    });
  }

  function bindLocalActions(root) {
    if (!root) return;
    root.querySelectorAll("[data-add-comm]").forEach(btn => btn.addEventListener("click", () => addCommToPole(btn.dataset.pole)));
    root.querySelectorAll("[data-edit-comm]").forEach(btn => btn.addEventListener("click", () => editCommGroup(btn.dataset.pole, btn.dataset.groupKey)));
    root.querySelectorAll("[data-edit-comm-spans]").forEach(btn => btn.addEventListener("click", () => editCommSpans(btn.dataset.pole, btn.dataset.groupKey)));
    root.querySelectorAll("[data-delete-comm-span]").forEach(btn => btn.addEventListener("click", () => deleteCommSpan(
      btn.dataset.span,
      btn.dataset.pole,
      btn.dataset.owner,
      btn.dataset.wireId || ""
    )));
    root.querySelectorAll("[data-clear-comm-midspan]").forEach(btn => btn.addEventListener("click", () => clearCommMidspan(
      btn.dataset.span,
      btn.dataset.pole,
      btn.dataset.owner,
      btn.dataset.wireId || ""
    )));
    root.querySelectorAll("[data-delete-comm]").forEach(btn => btn.addEventListener("click", () => deleteCommGroup(btn.dataset.pole, btn.dataset.groupKey)));
    root.querySelectorAll("[data-toggle-ug]").forEach(btn => btn.addEventListener("click", () => toggleUG(btn.dataset.pole)));
    root.querySelectorAll("[data-toggle-pco]").forEach(btn => btn.addEventListener("click", () => togglePCO(btn.dataset.pole)));
    root.querySelectorAll("[data-copy-mr]").forEach(btn => btn.addEventListener("click", () => copyMR(btn.dataset.pole)));
    root.querySelectorAll("[data-add-proposed-span]").forEach(btn => btn.addEventListener("click", () => addManualProposedSpan(btn.dataset.pole, root)));
    root.querySelectorAll("[data-delete-proposed-span]").forEach(btn => btn.addEventListener("click", () => deleteProposedSpan(btn.dataset.pole, btn.dataset.span)));
  }

  function toggleUG(poleId) {
    const pole = S.getPole(poleId);
    if (!pole) return;
    recordUndoSnapshot();
    S.upsertPole({
      ...pole,
      ugActive: !pole.ugActive,
      pcoActive: pole.ugActive ? pole.pcoActive : false
    });
    global.Calculations.recalculateSpansForPole(poleId);
    renderAffectedPoles([poleId, ...S.getConnectedSpans(poleId).map(span => S.getOtherPoleId(span, poleId)).filter(Boolean)]);
  }

  function togglePCO(poleId) {
    const pole = S.getPole(poleId);
    if (!pole) return;
    recordUndoSnapshot();
    S.upsertPole({
      ...pole,
      pcoActive: !pole.pcoActive,
      ugActive: pole.pcoActive ? pole.ugActive : false
    });
    global.Calculations.recalculateSpansForPole(poleId);
    renderAffectedPoles([poleId, ...S.getConnectedSpans(poleId).map(span => S.getOtherPoleId(span, poleId)).filter(Boolean)]);
  }

  async function copyMR(poleId) {
    const item = S.getState().mr.find(mr => mr.poleId === poleId);
    const text = item?.text || "";
    if (!text) return toast("No Make Ready to copy.", "warning");
    try {
      await navigator.clipboard.writeText(text);
      toast("Make Ready copied.", "success");
    } catch (error) {
      console.error(error);
      toast("Could not copy Make Ready.", "error");
    }
  }

  function addManualProposedSpan(poleId, root) {
    const select = root.querySelector(`[data-manual-proposed-target="${CSS.escape(poleId)}"]`);
    const targetPoleId = select?.value || "";
    if (!targetPoleId) return toast("Choose which pole the proposed span goes to.", "warning");
    const matchingSpans = connectedSpansSorted(poleId).filter(span =>
      span.fromPole === poleId && span.toPole === targetPoleId
    );
    const existing = matchingSpans.find(span =>
      !span.sourceSpanId && !/^(manual|additional)-proposed-/i.test(span.spanId)
    ) || matchingSpans[0] || null;
    recordUndoSnapshot();
    const existingSide = existing ? S.getSpanSide(existing.spanId, poleId) : null;
    const hasExistingProposal = Boolean(existingSide && (
      existingSide.isManualProposed || existingSide.proposedHOA || existingSide.proposedHOAChange ||
      existingSide.ocalcMS || existingSide.proposedMidspan || existingSide.notes
    ));
    const shouldCreateAdditional = Boolean(existing && hasExistingProposal);
    const span = shouldCreateAdditional || !existing ? S.upsertSpan({
      spanId: `${shouldCreateAdditional ? "additional" : "manual"}-proposed-${Date.now()}`,
      fromPole: poleId,
      toPole: targetPoleId,
      sourceSpanId: shouldCreateAdditional ? (existing.sourceSpanId || existing.spanId) : "",
      direction: existing?.direction || "",
      bearingDegrees: existing?.bearingDegrees ?? "",
      type: existing?.type || "Other",
      rawType: existing?.rawType || "",
      spanIndex: existing?.spanIndex || "",
      length: existing?.length || "",
      lengthDisplay: existing?.lengthDisplay || "",
      environment: existing?.environment || "NONE",
      environmentClearance: existing?.environmentClearance || "",
      midspanLowPower: existing?.midspanLowPower || "",
      midspanMaxCommHeight: existing?.midspanMaxCommHeight || "",
      rawSpanIds: existing?.rawSpanIds || [],
      isManualProposed: true
    }) : existing;
    S.upsertSpanSide({
      spanId: span.spanId,
      poleId,
      isManualProposed: true,
      isAdditionalProposed: shouldCreateAdditional,
      proposedHOA: shouldCreateAdditional ? "" : (existingSide?.proposedHOA || S.getPole(poleId)?.standaloneProposedHOA || "")
    });
    if (S.getPole(poleId)?.standaloneProposedHOA) S.updatePoleField(poleId, "standaloneProposedHOA", "");
    global.Calculations.recalculateSpan(span.spanId);
    renderAffectedPoles([poleId, targetPoleId]);
    markDirty();
  }

  async function deleteProposedSpan(poleId, spanId) {
    const side = S.getSpanSide(spanId, poleId);
    const span = S.getSpan(spanId);
    if (!side || !span) return;
    const label = shortSpanLabel(span);
    if (!(await confirmInApp("Delete Proposed Span", `Delete proposed span ${label}?`))) return;
    recordUndoSnapshot();
    if (!S.removeManualSpan(spanId)) {
      S.upsertSpanSide({
        ...side,
        proposedHOA: "",
        proposedHOAChange: "",
        nextPoleProposedAuto: false,
        proposedMidspan: "",
        ocalcMS: "",
        msProposed: "",
        finalMidspan: "",
        clearanceMSStatus: "",
        clearanceMSMessage: "",
        clearanceMSReason: "",
        clearanceMSIssue: false,
        proposedFlaggingStatus: "",
        proposedFlaggingMessage: "",
        endDrop: "",
        notes: "",
        isManualProposed: false,
        isAdditionalProposed: false
      });
    }
    global.Calculations.recalculateSpansForPole(poleId);
    renderAffectedPoles([poleId, S.getOtherPoleId(span, poleId)].filter(Boolean));
    markDirty();
  }

  function isLiveRecalcField(field) {
    return [
      "lowPower",
      "proposedHOA",
      "proposedHOAChange",
      "ocalcMS",
      "existingHOA",
      "existingHOAChange",
      "midspan",
      "environmentClearance",
      "midspanCommCommClearance",
      "midspanPowerCommClearance",
      "polePowerCommsClearance",
      "clearanceToPower",
      "commClearance",
      "boltClearance",
      "streetlightBracketCommClearance",
      "streetlightDripLoopCommClearance",
      "powerGuyCommClearance",
      "projectProfile",
      "position",
      "mrCase",
      "proposedOwner",
      "serviceDrop",
      "downGuy",
      "transferToNewPole",
      "resagServiceDrop"
    ].includes(field);
  }

  function updateCommGroupField(poleId, groupKey, field, value) {
    if (!["existingHOAChange", "transferToNewPole"].includes(field)) return [poleId].filter(Boolean);
    const nextValue = field === "transferToNewPole" ? Boolean(value) : value;
    const affected = new Set([poleId].filter(Boolean));
    S.getSpanCommsForPole(poleId)
      .filter(sc => commGroupKey(sc) === groupKey)
      .forEach(sc => {
        // SpanComm remains the persistence unit, but a comm-level field must
        // be identical on every relationship represented by this table row.
        global.Calculations.updateSpanCommField(sc.spanId, sc.poleId, sc.owner, sc.wireId || "", field, nextValue);
        poleIdsForSpan(sc.spanId).forEach(id => affected.add(id));
      });
    return Array.from(affected);
  }

  function defaultSpanForNewComm(poleId) {
    return connectedSpansSorted(poleId)[0] || null;
  }

  async function addCommToPole(poleId) {
    const values = await openAppDialog({
      title: "Add Comm",
      fields: [
        { name: "owner", label: "Owner / Comm" },
        { name: "existingHOA", label: "Existing HOA", placeholder: "20'" }
      ],
      confirmLabel: "Add Comm"
    });
    const owner = values?.owner?.trim();
    if (!owner) return;
    const existingHOA = values.existingHOA || "";
    const span = defaultSpanForNewComm(poleId);
    if (!span) return toast("That pole has no spans to associate with this comm.", "warning");
    recordUndoSnapshot();
    const wireId = `manual-${Date.now()}`;
    S.upsertComm(poleId, owner, existingHOA, "", { ownerBase: owner, rawOwner: owner, wireId });
    S.upsertSpanComm({ spanId: span.spanId, poleId, owner, ownerBase: owner, rawOwner: owner, existingHOA, wireId });
    global.Calculations.recalculateSpansForPole(poleId);
    renderAffectedPoles([poleId, S.getOtherPoleId(span, poleId)]);
  }

  async function editCommGroup(poleId, groupKey) {
    const group = groupedCommsForPole(poleId).find(item => item.key === groupKey);
    if (!group) return;
    const values = await openAppDialog({
      title: "Edit Comm",
      fields: [
        { name: "owner", label: "Owner / Comm", value: group.owner },
        { name: "existingHOA", label: "Existing HOA", value: group.existingHOA || "", placeholder: "20'" }
      ],
      confirmLabel: "Save"
    });
    const nextOwner = values?.owner?.trim();
    if (!nextOwner) return;
    const nextHOA = values.existingHOA || "";
    recordUndoSnapshot();
    group.rows.forEach(row => {
      S.removeSpanComm(row.spanId, row.poleId, row.owner, row.wireId || "");
      S.upsertComm(poleId, nextOwner, nextHOA, "", { ownerBase: nextOwner, rawOwner: nextOwner, wireId: row.wireId || "" });
      S.upsertSpanComm({ ...row, owner: nextOwner, ownerBase: nextOwner, rawOwner: nextOwner, existingHOA: nextHOA });
    });
    global.Calculations.recalculateSpansForPole(poleId);
    renderAffectedPoles([poleId]);
  }

  async function editCommSpans(poleId, groupKey) {
    const group = groupedCommsForPole(poleId).find(item => item.key === groupKey);
    if (!group) return;
    const spans = connectedSpansSorted(poleId);
    if (!spans.length) return toast("That pole has no available spans.", "warning");
    const spanValues = await openAppDialog({
      title: "Edit Comm Spans",
      description: group.owner,
      fields: [{
        name: "spanIds",
        label: "Spans",
        type: "checkbox-list",
        options: spans.map(span => ({
          value: span.spanId,
          label: shortSpanLabel(span),
          checked: group.rows.some(row => row.spanId === span.spanId)
        }))
      }],
      confirmLabel: "Continue"
    });
    if (!spanValues) return;
    const selected = new Set(spanValues.spanIds || []);
    if (!selected.size) return;
    recordUndoSnapshot();

    group.rows.forEach(row => {
      if (!selected.has(row.spanId)) S.removeSpanComm(row.spanId, row.poleId, row.owner, row.wireId || "");
    });

    for (const span of spans.filter(item => selected.has(item.spanId))) {
      const existing = group.rows.find(row => row.spanId === span.spanId);
      const midspanValues = await openAppDialog({
        title: "Midspan",
        description: shortSpanLabel(span),
        fields: [{ name: "midspan", label: "Midspan", value: existing?.midspan || "", placeholder: "Leave empty for REF" }],
        confirmLabel: "Save"
      });
      const midspan = midspanValues ? midspanValues.midspan : (existing?.midspan || "");
      if (existing) {
        global.Calculations.updateSpanCommField(existing.spanId, existing.poleId, existing.owner, existing.wireId || "", "midspan", midspan);
        continue;
      }
      const wireId = `manual-${Date.now()}-${span.spanId}`;
      S.upsertSpanComm({
        spanId: span.spanId,
        poleId,
        owner: group.rows[0].owner,
        ownerBase: group.rows[0].ownerBase || group.rows[0].owner,
        rawOwner: group.rows[0].rawOwner || group.rows[0].owner,
        existingHOA: group.existingHOA,
        existingHOAChange: group.existingHOAChange || "",
        serviceDrop: Boolean(group.rows[0].serviceDrop),
        downGuy: Boolean(group.rows[0].downGuy),
        transferToNewPole: group.rows.some(row => row.transferToNewPole),
        resagServiceDrop: false,
        midspan,
        wireId
      });
    }
    global.Calculations.recalculateSpansForPole(poleId);
    renderAffectedPoles([poleId]);
  }

  async function deleteCommGroup(poleId, groupKey) {
    const group = groupedCommsForPole(poleId).find(item => item.key === groupKey);
    if (!group || !(await confirmInApp("Delete Comm", `Delete ${group.owner}?`))) return;
    recordUndoSnapshot();
    group.rows.forEach(row => S.removeSpanComm(row.spanId, row.poleId, row.owner, row.wireId || ""));
    global.Calculations.recalculateSpansForPole(poleId);
    renderAffectedPoles([poleId]);
  }

  async function deleteCommSpan(spanId, poleId, owner, wireId = "") {
    if (!spanId || !poleId || !owner) return;
    if (!(await confirmInApp("Delete Span", "Delete only this span from the comm?"))) return;
    recordUndoSnapshot();
    S.removeSpanComm(spanId, poleId, owner, wireId || "");
    global.Calculations.recalculateSpansForPole(poleId);
    renderAffectedPoles([poleId, ...poleIdsForSpan(spanId)]);
  }

  async function clearCommMidspan(spanId, poleId, owner, wireId = "") {
    if (!spanId || !poleId || !owner) return;
    if (!(await confirmInApp("Clear Midspan", "Clear only this midspan and keep the span?"))) return;
    recordUndoSnapshot();
    global.Calculations.clearSpanCommMidspan(spanId, poleId, owner, wireId || "");
    renderAffectedPoles([poleId, ...poleIdsForSpan(spanId)]);
  }

  function editableInputKey(el) {
    return [
      el.dataset.scope || "",
      el.dataset.field || "",
      el.dataset.pole || "",
      el.dataset.span || "",
      el.dataset.owner || "",
      el.dataset.groupKey || "",
      el.dataset.wireId || "",
      el.dataset.powerKey || ""
    ].join("|");
  }

  function handleEditableInput(event) {
    const el = event.currentTarget;
    if (!el || el.tagName === "TEXTAREA" || el.tagName === "SELECT") return;
  }

  function handleEditableBlur(event) {
    const el = event.currentTarget;
    if (el.classList.contains("decimal-height-input")) {
      if (!el.value) {
        el.classList.remove("invalid");
        return;
      }
      const parsedDecimal = H.parseHeight(el.value);
      if (parsedDecimal === null) {
        el.classList.add("invalid");
        return;
      }
      el.classList.remove("invalid");
      return;
    }

    if (!el.classList.contains("height-input")) return;
    if (!el.value) {
      el.classList.remove("invalid");
      return;
    }
    const parsed = H.parseHeight(el.value);
    if (parsed === null) {
      el.classList.add("invalid");
      return;
    }
    el.classList.remove("invalid");
    el.value = H.formatHeight(parsed);
  }

  function scheduleDelayedMidspanRender(poleIds = []) {
    (poleIds || []).forEach(id => { if (id) delayedMidspanRenderPoleIds.add(id); });
    clearTimeout(delayedMidspanRenderTimer);
    delayedMidspanRenderTimer = setTimeout(() => {
      delayedMidspanRenderTimer = null;
      const ids = Array.from(delayedMidspanRenderPoleIds);
      delayedMidspanRenderPoleIds.clear();
      if (ids.length && (S.getState().ui.filter || "all") === "all") renderAffectedPoles(ids);
      else render();
    }, global.Calculations.CLEARANCE_FIX_DELAY_MS + 80);
  }

  function handleEditableChange(event) {
    const el = event.currentTarget;
    const scope = el.dataset.scope;
    const field = el.dataset.field;
    const value = el.type === "checkbox" ? el.checked : el.value;
    recordUndoSnapshot();

    if (scope === "pole") {
      S.updatePoleField(el.dataset.pole, field, value);
      global.Calculations.recalculateSpansForPole(el.dataset.pole);
    }

    if (scope === "spanSide") {
      global.Calculations.updateSpanSideField(el.dataset.span, el.dataset.pole, field, value);
    }

    if (scope === "span") {
      if (field === "environment") {
        global.Calculations.updateSpanField(el.dataset.span, "environment", value);
        global.Calculations.updateSpanField(
          el.dataset.span,
          "environmentClearance",
          S.defaultEnvironmentClearance ? S.defaultEnvironmentClearance(value) : ""
        );
      } else {
        global.Calculations.updateSpanField(el.dataset.span, field, value);
      }
    }

    if (scope === "spanComm") {
      global.Calculations.updateSpanCommField(el.dataset.span, el.dataset.pole, el.dataset.owner, el.dataset.wireId || "", field, value);
    }

    if (scope === "commGroup") {
      const affected = updateCommGroupField(el.dataset.pole, el.dataset.groupKey || "", field, value);
      renderAffectedPoles(affected);
      if (isLiveRecalcField(field)) scheduleDelayedMidspanRender(affected);
      return;
    }

    if (scope === "spanPower") {
      S.updateSpanPowerField(el.dataset.powerKey || "", field, value);
      global.Calculations.recalculateSpan(el.dataset.span);
    }

    if (scope === "settings") {
      S.updateSetting(field, value);
      global.Calculations.recalculateAll();
    }

    if (scope === "attachmentSettings") {
      const settings = S.getState().settings || {};
      if (field === "attachmentMessengerSize") settings.attachmentMessengerSize = String(value || "").trim();
      if (field === "fiberSize" && el.dataset.fiber) {
        settings.fiberSizes = settings.fiberSizes && typeof settings.fiberSizes === "object" ? settings.fiberSizes : {};
        settings.fiberSizes[el.dataset.fiber] = String(value || "").trim();
      }
      updateSaveButtonState();
      return;
    }

    if (scope === "poleClassCheck") {
      const rows = S.getState().poleClassChecks || [];
      const index = Number(el.dataset.index);
      if (Number.isInteger(index) && rows[index]) {
        rows[index] = global.ExcelImport.recalculatePoleClassCheck({ ...rows[index], [field]: value });
      }
      renderPoleClassResults();
      renderPoleLists();
      return;
    }

    const affectedPoleIds = affectedPoleIdsForElement(el, scope);
    if (scope !== "settings" && isLiveRecalcField(field)) {
      affectedPoleIds.forEach(id => global.Calculations.recalculateSpansForPole(id));
    }
    if (scope === "settings") render();
    else renderAffectedPoles(affectedPoleIds);

    if (["lowPower", "standaloneProposedHOA", "ocalcMS", "proposedMidspan", "proposedHOA", "proposedHOAChange", "existingHOA", "existingHOAChange", "midspan", "environmentClearance", "midspanCommCommClearance", "midspanPowerCommClearance", "polePowerCommsClearance", "clearanceToPower", "streetlightBracketCommClearance", "streetlightDripLoopCommClearance", "powerGuyCommClearance", "projectProfile", "position", "proposedOwner"].includes(field)) {
      scheduleDelayedMidspanRender(scope === "settings" ? [] : affectedPoleIds);
    }
  }

  function scrollToPole(poleId) {
    const card = document.querySelector(`[data-pole-card="${CSS.escape(poleId)}"]`);
    if (card) card.scrollIntoView({ behavior: "smooth", block: "start" });
  }

  function selectPole(poleId) {
    const state = S.getState();
    state.selectedPoleId = poleId;
    state.selectedSpanId = "";
    state.ui.activeView = "calculator";
    setPoleIndexOpen(false);
    render();
    scrollToPole(poleId);
  }

  function render() {
    // One top-level render always recalculates first so every table displays the
    // newest derived values after imports, edits, undo, or local restore.
    global.Calculations.recalculateAll();
    renderSummary();
    renderClearanceSettings();
    renderPoleClassResults();
    renderExcelReviewResults();
    renderPoleLists();
    renderAllPolesWorkspace();
    renderActiveView();
    updateAutoCalculateButtonState();
  }

  async function handleExcelImport(file) {
    if (!file) return;
    try {
      recordUndoSnapshot();
      saveFileHandle = null;
      await global.ExcelImport.importExcelFile(file);
      global.Calculations.recalculateAll();
      global.ExcelReview.runReview();
      render();
      markDirty();
      toast("Raw Excel imported. Available sheets were loaded.", "success");
    } catch (error) {
      console.error(error);
      toast(`Error importing Excel: ${error.message}`, "error");
    } finally {
      els.excelFileInput.value = "";
    }
  }

  async function handleExcelUpdate(file) {
    if (!file) return;
    try {
      recordUndoSnapshot();
      const previous = cloneCurrentState();
      await global.ExcelImport.importExcelFile(file);
      const imported = cloneCurrentState();
      const merged = mergeImportedUpdate(previous, imported);
      S.setState(merged);
      global.Calculations.recalculateAll();
      logExcelUpdateChanges(file.name, previous, cloneCurrentState());
      global.ExcelReview.runReview();
      render();
      markDirty();
      toast("Excel data updated. Existing movements were reconciled and midspans recalculated.", "success");
    } catch (error) {
      console.error(error);
      toast(`Error updating Excel data: ${error.message}`, "error");
    } finally {
      els.updateExcelFileInput.value = "";
    }
  }

  function bindEvents() {
    els.excelFileInput.addEventListener("change", event => handleExcelImport(event.target.files[0]));
    els.updateExcelFileInput.addEventListener("change", event => handleExcelUpdate(event.target.files[0]));
    els.exportProposedJsonBtn.addEventListener("click", () => {
      if (els.exportProposedJsonBtn.disabled) return;
      global.ProjectExport.exportProposedJson();
    });
    els.exportDebugJsonBtn.addEventListener("click", () => {
      if (els.exportDebugJsonBtn.disabled) return;
      global.ProjectExport.exportDebugJson();
    });
    els.autoCalculateBtn.addEventListener("click", () => {
      if (els.autoCalculateBtn.disabled) return;
      recordUndoSnapshot();
      const result = global.Calculations.autoCalculateMovements();
      if (result.disabled) {
        toast("Auto Calculate is only available in Top Comm mode.", "warning");
        return;
      }
      render();
      const passText = result.passes ? ` in ${result.passes} pass${result.passes === 1 ? "" : "es"}` : "";
      const stopText = result.maxPassesReached
        ? " Max pass safety limit reached."
        : result.stoppedByRepeat
          ? " Repeated state detected."
          : "";
      toast(`Auto Calculate: ${result.applied} applied, ${result.manual} need review, ${result.skipped} unchanged${passText}.${stopText}`, result.applied ? "success" : "warning");
    });
    els.rerunExcelReviewBtn.addEventListener("click", () => {
      if (els.rerunExcelReviewBtn.disabled) return;
      global.Calculations.recalculateAll();
      global.ExcelReview.clearResults();
      global.ExcelReview.runReview();
      renderExcelReviewResults();
      toast("Excel Review completed.", "success");
    });
    els.saveLocalBtn.addEventListener("click", async () => {
      try {
        await saveLocalFile();
      } catch (error) {
        if (error?.name === "AbortError") return;
        console.error(error);
        toast(`Error saving JSON: ${error.message}`, "error");
      }
    });
    els.loadLocalBtn.addEventListener("click", async () => {
      try {
        await loadLocalFile();
      } catch (error) {
        if (error?.name === "AbortError") return;
        console.error(error);
        toast(`Error loading JSON: ${error.message}`, "error");
      }
    });
    els.poleSearchInput.addEventListener("input", event => { S.getState().ui.search = event.target.value; render(); });
    els.warningFilterSelect.addEventListener("change", event => { S.getState().ui.filter = event.target.value; render(); });
    els.poleSearchInputTop.addEventListener("input", event => { S.getState().ui.search = event.target.value; render(); });
    els.warningFilterSelectTop.addEventListener("change", event => { S.getState().ui.filter = event.target.value; render(); });
    els.poleIndexToggle.addEventListener("click", () => setPoleIndexOpen(true));
    els.poleIndexClose.addEventListener("click", () => setPoleIndexOpen(false));
    els.poleIndexBackdrop.addEventListener("click", () => setPoleIndexOpen(false));
    window.addEventListener("scroll", updatePoleIndexToggleVisibility, { passive: true });
    window.addEventListener("resize", updatePoleIndexToggleVisibility);
    document.querySelectorAll("[data-view-tab]").forEach(btn => {
      btn.addEventListener("click", () => {
        S.getState().ui.activeView = btn.dataset.viewTab || "calculator";
        renderActiveView();
      });
    });
    document.addEventListener("keydown", event => {
      if (event.key === "Escape" && els.poleIndexDrawer?.classList.contains("open")) {
        setPoleIndexOpen(false);
        return;
      }
      if ((event.ctrlKey || event.metaKey) && !event.shiftKey && event.key.toLowerCase() === "z") {
        event.preventDefault();
        undoLastAction();
      }
      if ((event.ctrlKey || event.metaKey) && !event.shiftKey && event.key.toLowerCase() === "s") {
        event.preventDefault();
        saveLocalFile().catch(error => {
          if (error?.name === "AbortError") return;
          console.error(error);
          toast(`Error saving JSON: ${error.message}`, "error");
        });
      }
    });
    window.addEventListener("beforeunload", event => {
      if (!hasUnsavedChanges) return;
      event.preventDefault();
      event.returnValue = "";
    });
  }

  function init() {
    Object.assign(els, {
      excelFileInput: qs("excelFileInput"),
      updateExcelFileInput: qs("updateExcelFileInput"),
      exportProposedJsonBtn: qs("exportProposedJsonBtn"),
      exportDebugJsonBtn: qs("exportDebugJsonBtn"),
      autoCalculateBtn: qs("autoCalculateBtn"),
      saveLocalBtn: qs("saveLocalBtn"),
      loadLocalBtn: qs("loadLocalBtn"),
      projectMeta: qs("projectMeta"),
      poleSearchInput: qs("poleSearchInput"),
      warningFilterSelect: qs("warningFilterSelect"),
      polesList: qs("polesList"),
      poleSearchInputTop: qs("poleSearchInputTop"),
      warningFilterSelectTop: qs("warningFilterSelectTop"),
      polesListTop: qs("polesListTop"),
      poleIndexToggle: qs("poleIndexToggle"),
      poleIndexDrawer: qs("poleIndexDrawer"),
      poleIndexClose: qs("poleIndexClose"),
      poleIndexBackdrop: qs("poleIndexBackdrop"),
      topIndexPanel: qs("topIndexPanel"),
      polesOverview: qs("polesOverview"),
      poleClassResults: qs("poleClassResults"),
      poleClassTabWarning: qs("poleClassTabWarning"),
      poleClassTabWarningCount: qs("poleClassTabWarningCount"),
      excelReviewTabWarning: qs("excelReviewTabWarning"),
      excelReviewTabWarningCount: qs("excelReviewTabWarningCount"),
      excelReviewResults: qs("excelReviewResults"),
      excelReviewTimestamp: qs("excelReviewTimestamp"),
      rerunExcelReviewBtn: qs("rerunExcelReviewBtn"),
      appLayout: qs("appLayout"),
      clearanceSettings: qs("clearanceSettings"),
      toastHost: qs("toastHost")
    });

    bindEvents();
    global.FloatingCalculator?.setupFloatingCalculator();
    S.resetState();
    render();
    updatePoleIndexToggleVisibility();
    markClean(serializedSavePayload());
  }

  document.addEventListener("DOMContentLoaded", init);
})(window);
