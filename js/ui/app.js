(function (global) {
  "use strict";

  // app.js coordinates the browser UI: binding controls, dispatching edits,
  // persistence, and refreshing affected poles after recalculation.
  const S = global.AppStore;
  const H = global.HeightUtils;
  let mergeImportedUpdate, logExcelUpdateChanges;
  let poleActionHandlers = null;
  let history = null;
  let persistence = null;
  let mobileKeyboard = null;
  let themeControl = null;

  const poleResolutionOf = pole => S.getPoleResolution(pole);
  const ownerCanonicalOf = row => S.getOwnerCanonical(row);
  const ownerImportedOf = row => S.getOwnerImported(row);
  const spanRelationOf = span => S.getSpanRelation(span);

  const els = {};
  let poleNavigation = null;
  let spanLabel, shortSpanLabel, spanLengthDisplay, poleLink, displayMidspan, displayDecimalFeetInput, connectedSpansSorted, spanColorClass, spanRowClasses, spanChip, spanColorDot, commOwnerLabel, normalizedHeightLabel, commGroupKey, spanHasRealMidspan, isSpanEligibleForProposed, proposedSpansForPole, groupedCommsForPole, commMidspanEntries, renderCommSpanRefs, renderCommServiceDropValues, renderCommDownGuyValues, renderCommTransferValues, renderCommResagValues, renderCommMidspanValues, renderCommMaxHeightAtMSValues, renderCommRemoteValues, renderCommFlagging, poleIdsForSpan, affectedPoleIdsForElement, bindScrollLinks, replacePoleCard, renderAffectedPoles, noteList, commMidspanNote, poleClearanceNote, spanSideClearanceNote, renderMidspanClearanceStatus, renderCombinedSpanFlagging, renderCommHoaChangeValues, renderCommOtherHoaValues, renderClearanceStatus, renderClearanceSettings, updateAutoCalculateButtonState, updateFollowProposedButtonState, renderPoleClassResults, poleClassSeverity, reviewStatusBadge, renderReviewCheck, renderExcelReviewTabWarning, updateExcelReviewWarningActions, renderExcelReviewResults, setAllExcelReviewWarningsIgnored, hasCriticalPoleHeightIssue, renderPoleClassTabWarning, poleClassRowClass, renderPoleClassStatus, renderAnsiReferenceTable, renderActiveView, renderEnvironmentOptions, poleSummary, isSelfSupportingComm, poleHasProposed, poleFlaggingSummary, renderSummary, filteredPoles, renderPoleListItem, renderPoleLists, renderMRText, formatExistingViolationsText, renderMidamPreMoveViolations, renderPoleActions, renderPoleEditableHeader, renderSpanProposedTable, renderManualProposedSpanStarter, renderCommMovementTable, renderPowerTable, renderEquipmentTable, renderPoleWorkspace, renderAllPolesWorkspace;
  let delayedMidspanRenderTimer = null;
  const delayedMidspanRenderPoleIds = new Set();
  const editableInputTimers = new Map();
  const SPAN_COLOR_CLASS_COUNT = 10;
  let autoCalculateRunning = false;
  let autoCalculateCancelRequested = false;


  function qs(id) { return document.getElementById(id); }

  function renderDeploymentVersion() {
    if (!els.deploymentVersion) return;
    const build = global.AppDeploymentVersion || {};
    const version = String(build.version || "unknown").trim() || "unknown";
    const branch = String(build.branch || "").trim();
    const deployedAt = String(build.deployedAt || "").trim();
    // Keep the visible badge focused on the numeric application version. Build
    // and deployment details remain available in the tooltip without exposing
    // local/development labels in either the desktop or GitHub Pages UI.
    els.deploymentVersion.textContent = `Version ${version}`;
    const details = [
      `Version: ${version}`,
      build.commit ? `Commit: ${build.commit}` : "",
      branch ? `Branch: ${branch}` : "",
      deployedAt ? `Deployed: ${deployedAt}` : "",
      build.runId ? `GitHub Actions run: ${build.runId}` : ""
    ].filter(Boolean).join("\n");
    if (details) els.deploymentVersion.title = details;
  }

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

  // Auto Calculate runs in small asynchronous batches. This status panel is
  // updated between batches so large jobs remain visibly active and the user
  // can see which pass, pole and candidate are currently being evaluated.
  function setAutoCalculateProgress(visible, detail = {}) {
    if (!els.autoCalculateOverlay) return;
    els.autoCalculateOverlay.classList.toggle("hidden", !visible);
    document.body.classList.toggle("auto-calculate-running", visible);
    els.autoCalculateBtn?.setAttribute("aria-busy", String(visible));
    els.cancelAutoCalculateBtn?.classList.toggle("hidden", !visible);
    if (!visible) {
      if (els.cancelAutoCalculateBtn) {
        els.cancelAutoCalculateBtn.disabled = false;
        els.cancelAutoCalculateBtn.textContent = "Cancel";
      }
      return;
    }
    if (els.cancelAutoCalculateBtn) {
      els.cancelAutoCalculateBtn.disabled = autoCalculateCancelRequested;
      els.cancelAutoCalculateBtn.textContent = autoCalculateCancelRequested ? "Canceling..." : "Cancel";
    }

    const progress = Math.max(0, Math.min(100, Math.round(Number(detail.progress) || 0)));
    if (els.autoCalculateProgressBar) {
      els.autoCalculateProgressBar.value = progress;
      els.autoCalculateProgressBar.textContent = `${progress}%`;
    }
    if (els.autoCalculateProgressPercent) els.autoCalculateProgressPercent.textContent = `${progress}%`;

    let message = "Preparing pole data...";
    if (detail.phase === "canceling") {
      message = "Canceling and restoring the previous state...";
    } else if (detail.phase === "pole" || detail.phase === "candidate") {
      const passText = `Pass ${detail.pass || 1} of ${detail.maxPasses || 1}`;
      const poleText = `Pole ${detail.poleIndex || 0} of ${detail.poleCount || 0}`;
      const candidateText = detail.phase === "candidate" && detail.candidateCount
        ? ` · Candidate ${detail.candidateIndex} of ${detail.candidateCount}`
        : "";
      message = `${passText} · ${poleText}${detail.poleId ? ` · ${detail.poleId}` : ""}${candidateText}`;
    } else if (detail.phase === "complete") {
      message = "Finalizing results...";
    }
    if (els.autoCalculateProgressText) els.autoCalculateProgressText.textContent = message;
  }

  function setPoleIndexOpen(open) {
    poleNavigation?.setOpen(open);
  }

  function updateMobileKeyboardToggle() { mobileKeyboard?.updateToggle(); }
  function setMobileKeyboardEnabled(enabled) { mobileKeyboard?.setEnabled(enabled); }
  function setMobileHeightKeyboardOpen(open, input) { mobileKeyboard?.setOpen(open, input); }
  function bindMobileHeightKeyboard() { mobileKeyboard?.bind(); }
  function isMobileHeightKeyboardViewport() { return Boolean(mobileKeyboard?.isViewport()); }

  function updatePoleIndexToggleVisibility() {
    poleNavigation?.updateToggleVisibility();
  }

  const { openAppDialog, confirmInApp } = global.AppDialogs.create(escapeHtml);

  function cloneCurrentState() {
    return JSON.parse(JSON.stringify(S.getState()));
  }

  function recordUndoSnapshot() { history?.recordUndoSnapshot(); }
  function undoLastAction(source = "Ctrl+Z") { history?.undoLastAction(source); }
  function redoLastAction(source = "Ctrl+Y") { history?.redoLastAction(source); }

  function safeJobFilePart(value) {
    return global.AppPersistence.safeJobFilePart(value);
  }

  ({ mergeImportedUpdate, logExcelUpdateChanges } = global.UpdateDataReconciler.create({ S, ownerCanonicalOf, safeJobFilePart }));

  function savePayload() { return persistence?.savePayload(); }
  function serializedSavePayload() { return persistence?.serializedSavePayload() || ""; }
  function saveLocalFile() { return persistence?.saveLocalFile() || Promise.resolve(false); }
  function loadLocalFile() { return persistence?.loadLocalFile() || Promise.resolve(false); }
  function markDirty() { persistence?.markDirty(); }
  function markClean(serialized = "") { persistence?.markClean(serialized); }
  function updateSaveButtonState() { persistence?.updateSaveButtonState(); }

  function wireEditableEvents(root) {
    root.querySelectorAll("input[data-scope], textarea[data-scope], select[data-scope]").forEach(input => {
      if (input.classList.contains("height-input") || input.classList.contains("decimal-height-input")) {
        const decimal = input.classList.contains("decimal-height-input");
        const allowed = decimal ? /[0-9.\s-]/ : /[0-9.'"´′″\s-]/;
        // Feet/inch notation needs apostrophes and quotes, so use the normal
        // keyboard instead of Samsung's numeric-only layout.
        input.setAttribute("inputmode", isMobileHeightKeyboardViewport() ? "none" : "text");
        input.setAttribute("enterkeyhint", "next");
        input.addEventListener("focus", () => setMobileHeightKeyboardOpen(true, input));
        input.addEventListener("blur", () => {
          if (!mobileKeyboard?.ownsInput(input)) return;
          global.setTimeout(() => {
            if (document.activeElement !== input && !els.mobileHeightKeyboard?.contains(document.activeElement)) {
              setMobileHeightKeyboardOpen(false);
            }
          }, 0);
        });
        input.addEventListener("input", () => {
          const filtered = Array.from(input.value).filter(character => allowed.test(character)).join("");
          if (filtered !== input.value) {
            input.value = filtered;
            input.setSelectionRange(filtered.length, filtered.length);
          }
        });
      }
      if (input.classList.contains("pole-mr-editor")) {
        // Capture one undo point when a new editing session starts. The input
        // handler updates MR in place, so the textarea keeps focus while the
        // operator writes a multiline UG/PCO instruction.
        input.addEventListener("focus", () => { delete input.dataset.undoCaptured; });
      }
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
    return poleActionHandlers?.(root);
  }

  function isLiveRecalcField(field) {
    return [
      "lowPower",
      "proposedHOA",
      "proposedHOAChange",
      "ocalcMS",
      "existingHOA",
      "existingHOAChange",
      "otherHOA",
      "midspan",
      "attachmentHeightChange",
      "environmentClearance",
      "midspanCommCommClearance",
      "midspanPowerCommClearance",
      "midspanPrimaryPowerCommClearance",
      "polePowerCommsClearance",
      "primaryPowerCommsClearance",
      "clearanceToPower",
      "commClearance",
      "boltClearance",
      "streetlightBracketCommClearance",
      "streetlightDripLoopCommClearance",
      "powerGuyCommClearance",
      "projectProfile",
      "metronetWI",
      "position",
      "mrCase",
      "proposedOwner",
      "serviceDrop",
      "downGuy",
      "transferToNewPole",
      "resagServiceDrop",
      "pofActive",
      "ugSpanId",
      "ugSpanMRText",
      "ugReason",
      "ugMRText",
      "pcoMRText",
      "poleInsetReason",
      "poleInsetSpanId",
      "poleInsetHeight",
      "ugRiserDirection",
      "ancSize",
      "ancDistance",
      "ancDirection",
      "dgHoa",
      "ohgHoa",
      "ohgDirection",
      "actionActive",
      "actionHeight",
      "secureActive",
      "secondaryDripLoopActive",
      "secondaryDripLoopHoa",
      "raiseActive",
      "raiseHeight"
    ].includes(field);
  }

  function updateCommGroupField(poleId, groupKey, field, value) {
    if (!["existingHOA", "existingHOAChange", "otherHOA", "otherHOAActive", "transferToNewPole", "pofActive"].includes(field)) return [poleId].filter(Boolean);
    const nextValue = ["transferToNewPole", "otherHOAActive", "pofActive"].includes(field) ? Boolean(value) : value;
    const affected = new Set([poleId].filter(Boolean));
    const rows = S.getSpanCommsForPole(poleId).filter(sc => commGroupKey(sc) === groupKey);
    if (field === "otherHOAActive") {
      // Turning Other HOA off restores the original grouped behavior. Keep the
      // first entered per-span value as the shared HOA Change so disabling the
      // option never leaves the group with divergent heights.
      const sharedHOA = rows.map(row => String(row.existingHOAChange || "").trim()).find(Boolean) || "";
      rows.forEach(sc => {
        global.Calculations.updateSpanCommField(sc.spanId, sc.poleId, sc.owner, sc.wireId || "", field, nextValue);
        if (!nextValue) {
          global.Calculations.updateSpanCommField(sc.spanId, sc.poleId, sc.owner, sc.wireId || "", "existingHOAChange", sharedHOA);
        }
        poleIdsForSpan(sc.spanId).forEach(id => affected.add(id));
      });
      return Array.from(affected);
    }
    rows.forEach(sc => {
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
        { name: "existingHOA", label: "Existing HOA" }
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
    S.upsertComm(poleId, owner, existingHOA, "", { ownerCanonical: owner, ownerDisplay: owner, ownerImported: owner, wireId });
    S.upsertSpanComm({ spanId: span.spanId, poleId, owner, ownerCanonical: owner, ownerDisplay: owner, ownerImported: owner, existingHOA, wireId });
    global.Calculations.recalculateSpansForPole(poleId);
    renderAffectedPoles([poleId, S.getOtherPoleId(span, poleId)]);
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
        fields: [{ name: "midspan", label: "Midspan", value: existing?.midspan || "" }],
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
        ownerCanonical: S.getOwnerCanonical?.(group.rows[0]) || group.rows[0].owner,
        ownerDisplay: S.getOwnerDisplay?.(group.rows[0]) || group.rows[0].owner,
        ownerImported: S.getOwnerImported?.(group.rows[0]) || group.rows[0].owner,
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

  async function deleteSpanPower(powerKey, poleId, spanId) {
    if (!powerKey) return;
    const row = S.getSpanPowerForPole(poleId).find(item => item.key === powerKey);
    if (!row) return;
    const label = row.label || "power wire";
    if (!(await confirmInApp("Delete Imported Power", `Delete this ${label} power row?`))) return;
    recordUndoSnapshot();
    S.removeSpanPower(powerKey);
    const span = S.getSpan(spanId);
    const affected = [poleId, span ? S.getOtherPoleId(span, poleId) : ""].filter(Boolean);
    affected.forEach(id => global.Calculations.recalculateSpansForPole(id));
    renderAffectedPoles(affected);
    markDirty();
  }

  async function deletePowerEquipment(poleId, equipmentIndex) {
    const row = S.getPole(poleId)?.metadata?.powerEquipment?.[Number(equipmentIndex)];
    if (!row) return;
    const label = row.type || row.category || "equipment";
    if (!(await confirmInApp("Delete Equipment", `Delete this ${label} row?`))) return;
    recordUndoSnapshot();
    S.removePowerEquipment(poleId, Number(equipmentIndex));
    global.Calculations.recalculateSpansForPole(poleId);
    renderAffectedPoles([poleId, ...S.getConnectedSpans(poleId).flatMap(span => [span.fromPole, span.toPole]).filter(Boolean)]);
    markDirty();
  }

  function hidePole(poleId) {
    if (!S.getPole(poleId)) return;
    recordUndoSnapshot();
    const hidden = new Set(S.getState().ui.hiddenPoleIds || []);
    hidden.add(poleId);
    S.getState().ui.hiddenPoleIds = Array.from(hidden);
    S.getState().ui.revealedUnknownPoleIds = (S.getState().ui.revealedUnknownPoleIds || []).filter(id => id !== poleId);
    if (S.getState().selectedPoleId === poleId) S.getState().selectedPoleId = "";
    render();
    toast(`${poleId} hidden. Select it in the Pole Index to show it again.`, "success");
  }

  async function deletePole(poleId) {
    const pole = S.getPole(poleId);
    if (!pole) return;
    const spans = S.getConnectedSpans(poleId).length;
    const comms = groupedCommsForPole(poleId).length;
    const confirmed = await confirmInApp(
      "Delete Pole",
      `Delete ${poleId}, its ${spans} connected span${spans === 1 ? "" : "s"}, and ${comms} comm${comms === 1 ? "" : "s"}? This can be undone with Ctrl+Z.`,
      "Delete Pole"
    );
    if (!confirmed) return;

    recordUndoSnapshot();
    S.removePole(poleId);
    global.Calculations.recalculateAll();
    render();
    toast(`${poleId} deleted.`, "success");
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
      el.dataset.powerKey || "",
      el.dataset.equipmentIndex || ""
    ].join("|");
  }

  function handleEditableInput(event) {
    const el = event.currentTarget;
    if (!el) return;
    if (el.classList.contains("pole-mr-editor")) {
      if (!el.dataset.undoCaptured) {
        recordUndoSnapshot();
        el.dataset.undoCaptured = "true";
      }
      const poleId = el.dataset.pole || "";
      S.updatePoleField(poleId, el.dataset.field, el.value);
      if (el.dataset.field === "ugSpanMRText") {
        S.updatePoleField(poleId, "ugSpanMRSpanId", S.getPole(poleId)?.ugSpanId || "ALL");
      }
      const affectsUGConnections = ["ugMRText", "ugSpanId", "ugSpanMRText", "ugReason"].includes(el.dataset.field);
      const affectedPoleIds = affectsUGConnections
        ? (global.MRLogic.generateMRForUGChange?.(poleId) || [poleId])
        : (global.MRLogic.generateMRForPole(poleId), [poleId]);
      // Keep the active editor in place while refreshing adjacent cards. This
      // makes a changed UG reason appear on the previous pole's Forespan
      // without replacing the textarea under the user's cursor.
      affectedPoleIds.filter(id => id !== poleId).forEach(replacePoleCard);
      const generated = S.getState().mr.find(item => item.poleId === poleId)?.text || "";
      const output = el.closest("[data-pole-card]")?.querySelector(".mr-output");
      if (output) output.textContent = generated;
      markDirty();
      return;
    }
    if (el.tagName === "TEXTAREA" || el.tagName === "SELECT") return;
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
    let value = el.type === "checkbox" ? el.checked : el.value;
    // Multiline UG/PCO editors persist on every input so the card is not
    // replaced while typing. Blur/change only closes that editing session.
    if (el.classList.contains("pole-mr-editor")) {
      delete el.dataset.undoCaptured;
      return;
    }
    recordUndoSnapshot();

    if (scope === "pole") {
      const previousProposed = field === "standaloneProposedHOA"
        ? H.parseHeight(S.getPole(el.dataset.pole)?.standaloneProposedHOA || "")
        : null;
      if (field === "lowPower") {
        const pole = S.getPole(el.dataset.pole);
        S.upsertPole({
          ...pole,
          lowPower: value,
          metadata: { ...(pole?.metadata || {}), lowPowerBaseline: value }
        });
      } else {
        S.updatePoleField(el.dataset.pole, field, value);
      }
      if (field === "ugSpanId" && String(value || "").toUpperCase() !== "ALL") {
        const pole = S.getPole(el.dataset.pole);
        if (pole && pole.ugSpanMRSpanId !== value) {
          S.updatePoleField(el.dataset.pole, "ugSpanMRSpanId", value);
          S.updatePoleField(
            el.dataset.pole,
            "ugSpanMRText",
            global.MRLogic.getDefaultUGSpanTemplate?.(pole) || ""
          );
        }
      }
      global.Calculations.recalculateSpansForPole(el.dataset.pole);
      const nextProposed = field === "standaloneProposedHOA" ? H.parseHeight(value) : null;
      if (field === "standaloneProposedHOA"
        && (S.getState().settings?.autoMoveCommsToProposed === true
          || String(S.getState().settings?.autoMoveCommsToProposed).toLowerCase() === "true")
        && nextProposed !== null
        && nextProposed !== previousProposed) {
        global.AutoCalculateSolver?.applyProposedCommTarget?.(el.dataset.pole, nextProposed);
      }
    }

    if (scope === "equipment") {
      if (field === "raiseHeight" && value) {
        const row = S.getPole(el.dataset.pole)?.metadata?.powerEquipment?.[Number(el.dataset.equipmentIndex)];
        const source = H.parseHeight(row?.attachmentHeight || "");
        const target = H.parseHeight(value);
        if (source !== null && target !== null && target > source + 12) {
          value = H.formatHeight(source + 12);
          el.value = value;
          toast(`Streetlight Raise is limited to ${value} (1 foot maximum).`, "warning");
        } else if (source !== null && target !== null && target <= source) {
          value = "";
          el.value = "";
          toast("Streetlight New HOA must be above its Attachment Height.", "warning");
        }
      }
      S.updatePowerEquipmentField(el.dataset.pole, Number(el.dataset.equipmentIndex), field, value);
    }

    if (scope === "spanSide") {
      const previousProposed = field === "proposedHOA"
        ? H.parseHeight(S.getSpanSide(el.dataset.span, el.dataset.pole)?.proposedHOA || "")
        : null;
      global.Calculations.updateSpanSideField(el.dataset.span, el.dataset.pole, field, value);
      const nextProposed = field === "proposedHOA" ? H.parseHeight(value) : null;
      if (field === "proposedHOA"
        && (S.getState().settings?.autoMoveCommsToProposed === true
          || String(S.getState().settings?.autoMoveCommsToProposed).toLowerCase() === "true")
        && nextProposed !== null
        && nextProposed !== previousProposed) {
        global.AutoCalculateSolver?.applyProposedCommTarget?.(el.dataset.pole, nextProposed);
      }
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
      if (el.dataset.pole) global.Calculations.recalculateSpansForPole(el.dataset.pole);
      else global.Calculations.recalculateSpan(el.dataset.span);
    }

    if (scope === "settings") {
      S.updateSetting(field, value);
      if (["projectProfile", "metronetWI", "proposedOwner"].includes(field) && global.ExcelImport?.recalculatePoleClassCheck) {
        const settings = S.getState().settings || {};
        S.getState().poleClassChecks = (S.getState().poleClassChecks || [])
          .map(row => global.ExcelImport.recalculatePoleClassCheck({
            ...row,
            projectProfile: String(settings.projectProfile || "").toUpperCase(),
            metronetWI: String(settings.metronetWI || "").toUpperCase(),
            proposedOwner: String(settings.proposedOwner || "")
          }));
      }
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
      markDirty();
      return;
    }

    const affectedPoleIds = affectedPoleIdsForElement(el, scope);
    if (scope !== "settings" && isLiveRecalcField(field)) {
      affectedPoleIds.forEach(id => global.Calculations.recalculateSpansForPole(id));
    }
    if (scope === "settings") render();
    else renderAffectedPoles(affectedPoleIds);

    if (["lowPower", "standaloneProposedHOA", "poleInsetHeight", "ocalcMS", "proposedHOA", "proposedHOAChange", "existingHOA", "existingHOAChange", "otherHOA", "midspan", "attachmentHeightChange", "environmentClearance", "commClearance", "boltClearance", "midspanCommCommClearance", "midspanPowerCommClearance", "midspanPrimaryPowerCommClearance", "polePowerCommsClearance", "primaryPowerCommsClearance", "clearanceToPower", "streetlightBracketCommClearance", "streetlightGroundedBracketCommClearance", "streetlightDripLoopCommClearance", "powerGuyCommClearance", "projectProfile", "position", "proposedOwner"].includes(field)) {
      scheduleDelayedMidspanRender(scope === "settings" ? [] : affectedPoleIds);
    }
  }

  function scrollToPole(poleId) {
    poleNavigation?.scrollToPole(poleId);
  }

  function selectPole(poleId) {
    poleNavigation?.selectPole(poleId);
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
      persistence?.resetSaveFileHandle();
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
    els.followProposedBtn.addEventListener("click", () => {
      if (els.followProposedBtn.disabled) return;
      recordUndoSnapshot();
      const next = !(S.getState().settings?.autoMoveCommsToProposed === true);
      S.updateSetting("autoMoveCommsToProposed", next);
      updateFollowProposedButtonState();
      render();
      markDirty();
      toast(next ? "Follow Proposed enabled." : "Follow Proposed disabled.", "info");
    });
    els.exportDebugJsonBtn.addEventListener("click", () => {
      if (els.exportDebugJsonBtn.disabled) return;
      global.ProjectExport.exportDebugJson();
    });
    els.autoCalculateBtn.addEventListener("click", async () => {
      if (els.autoCalculateBtn.disabled || autoCalculateRunning) return;
      const autoCalculateSnapshot = cloneCurrentState();
      const undoStateBefore = history.captureTransaction();
      const dirtyStateBefore = persistence.captureDirtyState();
      recordUndoSnapshot();
      autoCalculateCancelRequested = false;
      autoCalculateRunning = true;
      updateAutoCalculateButtonState();
      setAutoCalculateProgress(true, { phase: "starting", progress: 0 });
      try {
        // Let the overlay paint before the first calculation begins.
        await new Promise(resolve => global.setTimeout(resolve, 0));
        const result = await global.Calculations.autoCalculateMovements({
          onProgress: detail => setAutoCalculateProgress(true, detail),
          isCancelled: () => autoCalculateCancelRequested
        });
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
      } catch (error) {
        if (error?.code === "AUTO_CALCULATE_CANCELLED") {
          history.restoreSnapshot(autoCalculateSnapshot);
          history.restoreTransaction(undoStateBefore);
          persistence.restoreDirtyState(dirtyStateBefore);
          toast("Auto Proposed canceled. No changes were applied.", "info");
          return;
        }
        global.console?.error?.(error);
        toast(`Auto Calculate failed: ${error.message}`, "error");
      } finally {
        autoCalculateRunning = false;
        autoCalculateCancelRequested = false;
        setAutoCalculateProgress(false);
        updateAutoCalculateButtonState();
      }
    });
    els.rerunExcelReviewBtn.addEventListener("click", () => {
      if (els.rerunExcelReviewBtn.disabled) return;
      global.Calculations.recalculateAll();
      global.ExcelReview.clearResults();
      global.ExcelReview.runReview();
      renderExcelReviewResults();
      toast("Excel Review completed.", "success");
    });
    els.ignoreHoaWarningsBtn.addEventListener("click", () => setAllExcelReviewWarningsIgnored("HOA"));
    els.ignoreFinalWarningsBtn.addEventListener("click", () => setAllExcelReviewWarningsIgnored("FINAL"));
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
    els.poleIndexToggle.addEventListener("click", () => {
      const isOpen = els.poleIndexDrawer?.classList.contains("open");
      setPoleIndexOpen(!isOpen);
    });
    els.cancelAutoCalculateBtn?.addEventListener("click", () => {
      if (!autoCalculateRunning || autoCalculateCancelRequested) return;
      autoCalculateCancelRequested = true;
      setAutoCalculateProgress(true, { phase: "canceling", progress: 0 });
    });
    els.themeToggleBtn?.addEventListener("click", () => themeControl.toggle());
    els.mobileKeyboardToggleBtn?.addEventListener("click", () => {
      setMobileKeyboardEnabled(!mobileKeyboard.isEnabled());
    });
    els.poleIndexClose.addEventListener("click", () => setPoleIndexOpen(false));
    els.poleIndexBackdrop.addEventListener("click", () => setPoleIndexOpen(false));
    els.jobNameInput?.addEventListener("change", event => {
      const nextName = String(event.target.value || "").trim()
        || S.jobNameFromFileName(S.getState().importedFileName);
      if (nextName === S.getState().jobName) return;
      recordUndoSnapshot();
      S.getState().jobName = nextName;
      event.target.value = nextName;
      markDirty();
    });
    els.jobNameInput?.addEventListener("keydown", event => {
      if (event.key !== "Enter") return;
      event.preventDefault();
      event.currentTarget.blur();
    });
    window.addEventListener("scroll", updatePoleIndexToggleVisibility, { passive: true });
    window.addEventListener("resize", updatePoleIndexToggleVisibility);
    document.querySelectorAll("[data-view-tab]").forEach(btn => {
      btn.addEventListener("click", () => {
        S.getState().ui.activeView = btn.dataset.viewTab || "calculator";
        renderActiveView();
      });
    });
    document.addEventListener("keydown", event => {
      if ((event.ctrlKey || event.metaKey) && !event.shiftKey && event.code === "Space") {
        if ((S.getState().ui.activeView || "calculator") !== "calculator") return;
        event.preventDefault();
        const isOpen = els.poleIndexDrawer?.classList.contains("open");
        setPoleIndexOpen(!isOpen);
        return;
      }
      if (event.key === "Escape" && els.poleIndexDrawer?.classList.contains("open")) {
        setPoleIndexOpen(false);
        return;
      }
      if ((event.ctrlKey || event.metaKey) && !event.shiftKey && event.key.toLowerCase() === "z") {
        event.preventDefault();
        undoLastAction("Ctrl+Z");
      }
      if ((event.ctrlKey || event.metaKey)
        && ((!event.shiftKey && event.key.toLowerCase() === "y")
          || (event.shiftKey && event.key.toLowerCase() === "z"))) {
        event.preventDefault();
        redoLastAction(event.shiftKey ? "Ctrl+Shift+Z" : "Ctrl+Y");
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
      if (!persistence?.hasUnsavedChanges()) return;
      event.preventDefault();
      event.returnValue = "";
    });
  }

  function init() {
    Object.assign(els, {
      excelFileInput: qs("excelFileInput"),
      updateExcelFileInput: qs("updateExcelFileInput"),
      exportProposedJsonBtn: qs("exportProposedJsonBtn"),
      followProposedBtn: qs("followProposedBtn"),
      exportDebugJsonBtn: qs("exportDebugJsonBtn"),
      autoCalculateBtn: qs("autoCalculateBtn"),
      mobileKeyboardToggleBtn: qs("toggleMobileKeyboardBtn"),
      autoCalculateOverlay: qs("autoCalculateOverlay"),
      autoCalculateProgressBar: qs("autoCalculateProgressBar"),
      autoCalculateProgressText: qs("autoCalculateProgressText"),
      autoCalculateProgressPercent: qs("autoCalculateProgressPercent"),
      cancelAutoCalculateBtn: qs("cancelAutoCalculateBtn"),
      saveLocalBtn: qs("saveLocalBtn"),
      loadLocalBtn: qs("loadLocalBtn"),
      themeToggleBtn: qs("themeToggleBtn"),
      projectMeta: qs("projectMeta"),
      jobNameInput: qs("jobNameInput"),
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
      ignoreHoaWarningsBtn: qs("ignoreHoaWarningsBtn"),
      ignoreFinalWarningsBtn: qs("ignoreFinalWarningsBtn"),
      appLayout: qs("appLayout"),
      clearanceSettings: qs("clearanceSettings"),
      toastHost: qs("toastHost"),
      mobileHeightKeyboard: qs("mobileHeightKeyboard")
    });

    themeControl = global.AppTheme.create({ els });
    mobileKeyboard = global.MobileHeightKeyboard.create({
      els,
      commitInput(input) {
        if (!input) return;
        handleEditableBlur({ currentTarget: input });
        handleEditableChange({ currentTarget: input });
      }
    });
    persistence = global.AppPersistence.create({
      S,
      Calculations: global.Calculations,
      ProjectExport: global.ProjectExport,
      ExcelImport: global.ExcelImport,
      ExcelReview: global.ExcelReview,
      els,
      render,
      recordUndoSnapshot,
      toast
    });
    history = global.AppHistory.create({
      S,
      Calculations: global.Calculations,
      ExcelReview: global.ExcelReview,
      markDirty,
      render,
      toast,
      clearPendingRender() {
        clearTimeout(delayedMidspanRenderTimer);
        delayedMidspanRenderTimer = null;
        delayedMidspanRenderPoleIds.clear();
      }
    });
    poleNavigation = global.PoleNavigation.create({
      document,
      els,
      mobileKeyboard,
      recordUndoSnapshot,
      render,
      S
    });

    const rendererContext = {
      get autoCalculateRunning() { return autoCalculateRunning; },
      bindLocalActions: bindLocalActions,
      els: els,
      escapeHtml: escapeHtml,
      H: H,
      markDirty: markDirty,
      ownerCanonicalOf: ownerCanonicalOf,
      ownerImportedOf: ownerImportedOf,
      poleResolutionOf: poleResolutionOf,
      recordUndoSnapshot: recordUndoSnapshot,
      render: render,
      S: S,
      scrollToPole: scrollToPole,
      selectPole: selectPole,
      setPoleIndexOpen: setPoleIndexOpen,
      SPAN_COLOR_CLASS_COUNT: SPAN_COLOR_CLASS_COUNT,
      spanRelationOf: spanRelationOf,
      toast: toast,
      updatePoleIndexToggleVisibility: updatePoleIndexToggleVisibility,
      wireEditableEvents: wireEditableEvents
    };
    ({ spanLabel, shortSpanLabel, spanLengthDisplay, poleLink, displayMidspan, displayDecimalFeetInput, connectedSpansSorted, spanColorClass, spanRowClasses, spanChip, spanColorDot, commOwnerLabel, normalizedHeightLabel, commGroupKey, spanHasRealMidspan, isSpanEligibleForProposed, proposedSpansForPole, groupedCommsForPole, commMidspanEntries, renderCommSpanRefs, renderCommServiceDropValues, renderCommDownGuyValues, renderCommTransferValues, renderCommResagValues, renderCommMidspanValues, renderCommMaxHeightAtMSValues, renderCommRemoteValues, renderCommFlagging, poleIdsForSpan, affectedPoleIdsForElement, bindScrollLinks, replacePoleCard, renderAffectedPoles, noteList, commMidspanNote, poleClearanceNote, spanSideClearanceNote, renderMidspanClearanceStatus, renderCombinedSpanFlagging, renderCommHoaChangeValues, renderCommOtherHoaValues, renderClearanceStatus, renderClearanceSettings, updateAutoCalculateButtonState, updateFollowProposedButtonState, renderPoleClassResults, poleClassSeverity, reviewStatusBadge, renderReviewCheck, renderExcelReviewTabWarning, updateExcelReviewWarningActions, renderExcelReviewResults, setAllExcelReviewWarningsIgnored, hasCriticalPoleHeightIssue, renderPoleClassTabWarning, poleClassRowClass, renderPoleClassStatus, renderAnsiReferenceTable, renderActiveView, renderEnvironmentOptions, poleSummary, isSelfSupportingComm, poleHasProposed, poleFlaggingSummary, renderSummary, filteredPoles, renderPoleListItem, renderPoleLists, renderMRText, formatExistingViolationsText, renderMidamPreMoveViolations, renderPoleActions, renderPoleEditableHeader, renderSpanProposedTable, renderManualProposedSpanStarter, renderCommMovementTable, renderPowerTable, renderEquipmentTable, renderPoleWorkspace, renderAllPolesWorkspace } = global.PoleWorkspaceRenderer.create(rendererContext));
    poleActionHandlers = global.PoleActions.create({
      S,
      H,
      addCommToPole,
      clearCommMidspan,
      confirmInApp,
      connectedSpansSorted,
      deleteCommGroup,
      deleteCommSpan,
      deletePole,
      deletePowerEquipment,
      deleteSpanPower,
      editCommSpans,
      formatExistingViolationsText,
      groupedCommsForPole,
      hidePole,
      markDirty,
      poleIdsForSpan,
      poleResolutionOf,
      recordUndoSnapshot,
      renderAffectedPoles,
      shortSpanLabel,
      toast,
      updateCommGroupField
    }).bindLocalActions;
    themeControl.apply(themeControl.savedTheme(), false);
    renderDeploymentVersion();
    updateMobileKeyboardToggle();
    bindEvents();
    bindMobileHeightKeyboard();
    global.FloatingCalculator?.setupFloatingCalculator();
    S.resetState();
    render();
    // Desktop/web layouts have enough room for the navigation drawer up
    // front. Phones keep it closed so the calculator remains unobstructed.
    if (window.matchMedia?.("(min-width: 701px)").matches) {
      setPoleIndexOpen(true);
    }
    updatePoleIndexToggleVisibility();
    markClean(serializedSavePayload());
  }

  document.addEventListener("DOMContentLoaded", init);
})(window);
