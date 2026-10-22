(function (global) {
  "use strict";

  function create(context = {}) {
    const {
      S, H, addCommToPole, clearCommMidspan, confirmInApp, connectedSpansSorted,
      deleteCommGroup, deleteCommSpan, deletePole, deletePowerEquipment, deleteSpanPower,
      editCommSpans, formatExistingViolationsText, groupedCommsForPole, hidePole, markDirty,
      poleIdsForSpan, poleResolutionOf, recordUndoSnapshot, renderAffectedPoles,
      shortSpanLabel, toast, updateCommGroupField
    } = context;
  function bindLocalActions(root) {
    if (!root) return;
    root.querySelectorAll("[data-inline-comm-edit]").forEach(cell => cell.addEventListener("dblclick", () => beginInlineCommEdit(cell)));
    root.querySelectorAll("[data-inline-equipment-edit]").forEach(cell => cell.addEventListener("dblclick", () => beginInlineEquipmentEdit(cell)));
    root.querySelectorAll("[data-add-comm]").forEach(btn => btn.addEventListener("click", () => addCommToPole(btn.dataset.pole)));
    root.querySelectorAll("[data-hide-pole]").forEach(btn => btn.addEventListener("click", () => hidePole(btn.dataset.pole)));
    root.querySelectorAll("[data-delete-pole]").forEach(btn => btn.addEventListener("click", () => deletePole(btn.dataset.pole)));
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
    root.querySelectorAll("[data-delete-span-power]").forEach(btn => btn.addEventListener("click", () => deleteSpanPower(
      btn.dataset.powerKey,
      btn.dataset.pole,
      btn.dataset.span
    )));
    root.querySelectorAll("[data-delete-equipment]").forEach(btn => btn.addEventListener("click", () => deletePowerEquipment(
      btn.dataset.pole,
      btn.dataset.equipmentIndex
    )));
    root.querySelectorAll("[data-toggle-ug]").forEach(btn => btn.addEventListener("click", () => toggleUG(btn.dataset.pole)));
    root.querySelectorAll("[data-toggle-pco]").forEach(btn => btn.addEventListener("click", () => togglePCO(btn.dataset.pole)));
    root.querySelectorAll("[data-toggle-pole-inset]").forEach(btn => btn.addEventListener("click", () => togglePoleInset(btn.dataset.pole)));
    root.querySelectorAll("[data-toggle-riser]").forEach(btn => btn.addEventListener("click", () => toggleRiser(btn.dataset.pole)));
    root.querySelectorAll("[data-toggle-secondary-drip-loop]").forEach(btn => btn.addEventListener("click", () => toggleSecondaryDripLoop(btn.dataset.pole)));
    root.querySelectorAll("[data-toggle-span-ug]").forEach(btn => btn.addEventListener("click", () => toggleSpanUG(btn.dataset.pole, btn.dataset.span)));
    root.querySelectorAll("[data-toggle-anc]").forEach(btn => btn.addEventListener("click", () => toggleANC(btn.dataset.pole)));
    root.querySelectorAll("[data-toggle-ohg]").forEach(btn => btn.addEventListener("click", () => toggleOHG(btn.dataset.pole)));
    root.querySelectorAll("[data-toggle-red-tag]").forEach(btn => btn.addEventListener("click", () => toggleRedTag(btn.dataset.pole)));
    root.querySelectorAll("[data-toggle-less-than-12-comm]").forEach(btn => btn.addEventListener("click", () => toggleLessThan12CommClearance(btn.dataset.pole)));
    root.querySelectorAll("[data-toggle-riser-secure]").forEach(btn => btn.addEventListener("click", () => toggleRiserSecure(btn.dataset.pole)));
    root.querySelectorAll("[data-toggle-comm-movements]").forEach(btn => btn.addEventListener("click", () => toggleCommMovements(btn.dataset.pole)));
    root.querySelectorAll("[data-clear-comm-movements]").forEach(btn => btn.addEventListener("click", () => clearCommMovements(btn.dataset.pole)));
    root.querySelectorAll("[data-copy-mr]").forEach(btn => btn.addEventListener("click", () => copyMR(btn.dataset.pole)));
    root.querySelectorAll("[data-copy-existing-violations]").forEach(btn => btn.addEventListener("click", async () => {
      const violations = global.Calculations?.getMidamPreMoveViolations?.(btn.dataset.pole) || [];
      const text = formatExistingViolationsText(violations);
      if (!text) return toast("No existing violations to copy.", "warning");
      try {
        await navigator.clipboard.writeText(text);
        toast("Existing violations copied.", "success");
      } catch (error) {
        console.error(error);
        toast("Could not copy existing violations.", "error");
      }
    }));
    root.querySelectorAll("[data-add-proposed-span]").forEach(btn => btn.addEventListener("click", () => addManualProposedSpan(btn.dataset.pole, root)));
    root.querySelectorAll("[data-delete-proposed-span]").forEach(btn => btn.addEventListener("click", () => deleteProposedSpan(btn.dataset.pole, btn.dataset.span)));
  }

  function toggleUG(poleId) {
    const pole = S.getPole(poleId);
    if (!pole) return;
    recordUndoSnapshot();
    const currentResolution = poleResolutionOf(pole);
    const nextUGActive = currentResolution !== "UG";
    S.upsertPole({
      ...pole,
      ugActive: nextUGActive,
      ugSpanId: pole.ugSpanId || "ALL",
      pcoActive: nextUGActive ? false : pole.pcoActive
    });
    global.Calculations.recalculateSpansForPole(poleId);
    renderAffectedPoles([poleId, ...S.getConnectedSpans(poleId).map(span => S.getOtherPoleId(span, poleId)).filter(Boolean)]);
  }

  function togglePCO(poleId) {
    const pole = S.getPole(poleId);
    if (!pole) return;
    recordUndoSnapshot();
    const currentResolution = poleResolutionOf(pole);
    const nextPCOActive = currentResolution !== "PCO";
    S.upsertPole({
      ...pole,
      pcoActive: nextPCOActive,
      ugActive: nextPCOActive ? false : pole.ugActive
    });
    global.Calculations.recalculateSpansForPole(poleId);
    renderAffectedPoles([poleId, ...S.getConnectedSpans(poleId).map(span => S.getOtherPoleId(span, poleId)).filter(Boolean)]);
  }

  function toggleRiser(poleId) {
    const pole = S.getPole(poleId);
    if (!pole || !global.MRLogic.isRiserAvailable?.(poleId)) return;
    recordUndoSnapshot();
    S.upsertPole({
      ...pole,
      riserActive: !global.MRLogic.isRiserEnabled?.(poleId)
    });
    global.Calculations.recalculateSpansForPole(poleId);
    renderAffectedPoles([poleId, ...S.getConnectedSpans(poleId).map(span => S.getOtherPoleId(span, poleId)).filter(Boolean)]);
  }

  function toggleSecondaryDripLoop(poleId) {
    const pole = S.getPole(poleId);
    if (!pole || String(S.getState().settings?.projectProfile || "").toUpperCase() !== "OLSSON_OPPD") return;
    recordUndoSnapshot();
    const active = !pole.secondaryDripLoopActive;
    S.upsertPole({ ...pole, secondaryDripLoopActive: active, secondaryDripLoopHoa: pole.secondaryDripLoopHoa || "" });
    global.Calculations.recalculateSpansForPole(poleId);
    renderAffectedPoles([poleId]);
    markDirty();
  }

  function updateCommGroupOwner(poleId, groupKey, value) {
    const nextOwner = String(value || "").trim();
    const group = groupedCommsForPole(poleId).find(item => item.key === groupKey);
    if (!group || !nextOwner) return false;
    group.rows.forEach(row => {
      S.removeSpanComm(row.spanId, row.poleId, row.owner, row.wireId || "");
      S.upsertComm(poleId, nextOwner, row.existingHOA || group.existingHOA || "", "", {
        ownerCanonical: nextOwner,
        ownerDisplay: nextOwner,
        ownerImported: nextOwner,
        wireId: row.wireId || "",
        existingHOAChange: row.existingHOAChange || "",
        otherHOA: row.otherHOA || "",
        pofActive: row.pofActive
      });
      S.upsertSpanComm({ ...row, owner: nextOwner, ownerCanonical: nextOwner, ownerDisplay: nextOwner, ownerImported: nextOwner });
    });
    return true;
  }

  function beginInlineCommEdit(cell) {
    if (!cell || cell.querySelector("input")) return;
    const poleId = cell.dataset.pole || "";
    const groupKey = cell.dataset.groupKey || "";
    const field = cell.dataset.inlineCommEdit || "";
    const group = groupedCommsForPole(poleId).find(item => item.key === groupKey);
    if (!group || !["owner", "existingHOA"].includes(field)) return;
    const original = field === "owner" ? group.owner : (group.existingHOA || "");
    const input = document.createElement("input");
    input.className = `input inline-comm-input ${field === "existingHOA" ? "height-input" : ""}`;
    input.value = original;
    input.setAttribute("aria-label", field === "owner" ? "Owner / Comm" : "Existing HOA");
    cell.replaceChildren(input);
    input.focus();
    input.select();
    let saved = false;
    const save = () => {
      if (saved) return;
      saved = true;
      const value = input.value.trim();
      if (!value) {
        renderAffectedPoles([poleId]);
        return;
      }
      if (field === "existingHOA" && H.parseHeight(value) === null) {
        toast("Existing HOA must be a valid height such as 18' 6\".", "warning");
        renderAffectedPoles([poleId]);
        return;
      }
      recordUndoSnapshot();
      const affected = new Set([poleId]);
      if (field === "owner") {
        if (!updateCommGroupOwner(poleId, groupKey, value)) return renderAffectedPoles([poleId]);
        group.rows.forEach(row => poleIdsForSpan(row.spanId).forEach(id => affected.add(id)));
      } else {
        updateCommGroupField(poleId, groupKey, "existingHOA", value).forEach(id => affected.add(id));
      }
      renderAffectedPoles(Array.from(affected));
      markDirty();
    };
    input.addEventListener("blur", save, { once: true });
    input.addEventListener("keydown", event => {
      if (event.key === "Enter") {
        event.preventDefault();
        input.blur();
      }
      if (event.key === "Escape") {
        saved = true;
        renderAffectedPoles([poleId]);
      }
    });
  }

  function beginInlineEquipmentEdit(cell) {
    if (!cell || cell.querySelector("input")) return;
    const poleId = cell.dataset.pole || "";
    const equipmentIndex = Number(cell.dataset.equipmentIndex);
    const field = cell.dataset.inlineEquipmentEdit || "";
    const row = S.getPole(poleId)?.metadata?.powerEquipment?.[equipmentIndex];
    if (!row || !["owner", "attachmentHeight"].includes(field)) return;
    const input = document.createElement("input");
    input.className = `input inline-equipment-input ${field === "attachmentHeight" ? "height-input" : ""}`;
    input.value = row[field] || "";
    input.setAttribute("aria-label", field === "owner" ? "Equipment owner" : "Attachment height");
    cell.replaceChildren(input);
    input.focus();
    input.select();
    let saved = false;
    const save = () => {
      if (saved) return;
      saved = true;
      const value = input.value.trim();
      if (!value || (field === "attachmentHeight" && H.parseHeight(value) === null)) {
        if (field === "attachmentHeight") toast("Attachment Height must be a valid height such as 18' 6\\\".", "warning");
        renderAffectedPoles([poleId]);
        return;
      }
      recordUndoSnapshot();
      S.updatePowerEquipmentField(poleId, equipmentIndex, field, value);
      global.Calculations.recalculateSpansForPole(poleId);
      renderAffectedPoles([poleId, ...S.getConnectedSpans(poleId).map(span => S.getOtherPoleId(span, poleId)).filter(Boolean)]);
      markDirty();
    };
    input.addEventListener("blur", save, { once: true });
    input.addEventListener("keydown", event => {
      if (event.key === "Enter") {
        event.preventDefault();
        input.blur();
      }
      if (event.key === "Escape") {
        saved = true;
        renderAffectedPoles([poleId]);
      }
    });
  }

  function toggleSpanUG(poleId, spanId) {
    const pole = S.getPole(poleId);
    const span = S.getSpan(spanId);
    if (!pole || !span || poleResolutionOf(pole) !== "AERIAL") return;
    recordUndoSnapshot();
    const enabling = span.ugActive !== true;
    S.updateSpanField(spanId, "ugActive", enabling);
    if (enabling && !span.ugReason) S.updateSpanField(spanId, "ugReason", "CLEARANCE VIOLATION");
    global.Calculations.recalculateSpan(spanId);
    renderAffectedPoles([span.fromPole, span.toPole].filter(Boolean));
    markDirty();
  }

  function toggleANC(poleId) {
    const pole = S.getPole(poleId);
    const projectProfile = String(S.getState().settings?.projectProfile || "").toUpperCase();
    if (!pole || !["METRONET", "INTEC"].includes(projectProfile) || poleResolutionOf(pole) !== "AERIAL") return;
    recordUndoSnapshot();
    S.upsertPole({
      ...pole,
      ancActive: global.MRLogic.isANCEnabled?.(poleId) ? false : true,
      ancSize: pole.ancSize || "8\"",
      ancDistance: pole.ancDistance || "",
      ancDistanceManual: pole.ancDistanceManual === true,
      ancDirectionManual: pole.ancDirectionManual === true
    });
    global.Calculations.recalculateSpansForPole(poleId);
    renderAffectedPoles([poleId]);
  }

  function toggleOHG(poleId) {
    const pole = S.getPole(poleId);
    const projectProfile = String(S.getState().settings?.projectProfile || "").toUpperCase();
    if (!pole || !["METRONET", "OLSSON_OPPD"].includes(projectProfile) || poleResolutionOf(pole) !== "AERIAL") return;
    recordUndoSnapshot();
    S.upsertPole({
      ...pole,
      ohgActive: !pole.ohgActive
    });
    global.Calculations.recalculateSpansForPole(poleId);
    renderAffectedPoles([poleId]);
  }

  function toggleRedTag(poleId) {
    const pole = S.getPole(poleId);
    const projectProfile = String(S.getState().settings?.projectProfile || "").toUpperCase();
    if (!pole || projectProfile !== "OLSSON_OPPD") return;
    recordUndoSnapshot();
    S.upsertPole({
      ...pole,
      redTagActive: !pole.redTagActive,
      redTagReason: pole.redTagReason || "CLEARANCES"
    });
    global.Calculations.recalculateSpansForPole(poleId);
    renderAffectedPoles([poleId]);
    markDirty();
  }

  function toggleLessThan12CommClearance(poleId) {
    const pole = S.getPole(poleId);
    if (!pole) return;
    recordUndoSnapshot();
    S.updatePoleField(poleId, "lessThan12CommClearanceActive", !pole.lessThan12CommClearanceActive);
    global.Calculations.recalculateSpansForPole(poleId);
    renderAffectedPoles([poleId]);
    markDirty();
  }

  function toggleRiserSecure(poleId) {
    const pole = S.getPole(poleId);
    const rows = Array.isArray(pole?.metadata?.powerEquipment) ? pole.metadata.powerEquipment : [];
    const riserIndexes = rows
      .map((row, index) => ({ row, index }))
      .filter(item => String(item.row.category || item.row.type || "").toUpperCase().includes("RISER"))
      .map(item => item.index);
    if (!riserIndexes.length) return;
    const active = riserIndexes.every(index => rows[index].secureActive === true);
    recordUndoSnapshot();
    riserIndexes.forEach(index => S.updatePowerEquipmentField(poleId, index, "secureActive", !active));
    global.Calculations.recalculateSpansForPole(poleId);
    renderAffectedPoles([poleId]);
    markDirty();
  }

  function toggleCommMovements(poleId) {
    const pole = S.getPole(poleId);
    const hasChanges = S.getSpanCommsForPole(poleId).some(row => Boolean(row.existingHOAChange || row.otherHOA));
    if (!pole || !hasChanges) return;
    recordUndoSnapshot();
    S.upsertPole({ ...pole, commMovementsActive: pole.commMovementsActive === false });
    global.Calculations.recalculateSpansForPole(poleId);
    renderAffectedPoles([poleId, ...S.getConnectedSpans(poleId).map(span => S.getOtherPoleId(span, poleId)).filter(Boolean)]);
  }

  async function clearCommMovements(poleId) {
    const pole = S.getPole(poleId);
    const rows = S.getSpanCommsForPole(poleId);
    const hasChanges = rows.some(row => Boolean(row.existingHOAChange || row.otherHOA));
    if (!pole || !hasChanges) return;
    const confirmed = await confirmInApp(
      "Clear Comm Movements",
      "Clear all HOA Change and Other HOA movement heights for this pole?",
      "Clear Moves"
    );
    if (!confirmed) return;

    recordUndoSnapshot();
    rows.forEach(row => S.upsertSpanComm({
      ...row,
      existingHOAChange: "",
      otherHOA: "",
      autoCalcStatus: "",
      autoCalcMessage: ""
    }));
    S.upsertPole({ ...pole, commMovementsActive: false });

    const affectedPoles = new Set([
      poleId,
      ...S.getConnectedSpans(poleId).flatMap(span => [span.fromPole, span.toPole]).filter(Boolean)
    ]);
    affectedPoles.forEach(id => global.Calculations.recalculateSpansForPole(id));
    renderAffectedPoles(Array.from(affectedPoles));
    markDirty();
    toast("Comm movements cleared.", "success");
  }

  function togglePoleInset(poleId) {
    const pole = S.getPole(poleId);
    if (!pole || poleResolutionOf(pole) !== "AERIAL") return;
    recordUndoSnapshot();
    S.upsertPole({
      ...pole,
      poleInsetActive: !pole.poleInsetActive,
      poleInsetReason: pole.poleInsetReason === "OVERLOADED" ? "OVERLOADED" : "FAILING_CLEARANCES",
      poleInsetSpanId: !pole.poleInsetActive
        ? (pole.poleInsetSpanId || global.Calculations.getPoleInsetDetails?.(poleId)?.spanId || "")
        : (pole.poleInsetSpanId || "")
    });
    global.Calculations.recalculateSpansForPole(poleId);
    renderAffectedPoles([poleId]);
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
      existingSide.ocalcMS || existingSide.proposedMidspan || existingSide.notes || existingSide.isSlack
    ));
    const shouldCreateAdditional = Boolean(existing && hasExistingProposal);
    const span = shouldCreateAdditional || !existing ? S.upsertSpan({
      spanId: `${shouldCreateAdditional ? "additional" : "manual"}-proposed-${Date.now()}`,
      fromPole: poleId,
      toPole: targetPoleId,
      sourceSpanId: shouldCreateAdditional ? (existing.sourceSpanId || existing.spanId) : "",
      direction: existing?.direction || "",
      bearingDegrees: existing?.bearingDegrees ?? "",
      spanRelation: existing?.spanRelation || existing?.type || "Other",
      importedRelation: existing?.importedRelation || existing?.rawType || "",
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
      isProposedExcluded: false,
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
        isSlack: false,
        isManualProposed: false,
        isAdditionalProposed: false,
        isProposedExcluded: true
      });
    }
    global.Calculations.recalculateSpansForPole(poleId);
    renderAffectedPoles([poleId, S.getOtherPoleId(span, poleId)].filter(Boolean));
    markDirty();
  }

    return { bindLocalActions };
  }

  global.PoleActions = { create };
})(window);