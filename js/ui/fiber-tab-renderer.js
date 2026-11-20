(function (global) {
  "use strict";

  const FIBER_TYPES = ["24CT Fiber", "48CT Fiber", "MST"];
  const DIRECTIONS = ["N", "NE", "E", "SE", "S", "SW", "W", "NW"];
  const OPPOSITE = { N: "S", NE: "SW", E: "W", SE: "NW", S: "N", SW: "NE", W: "E", NW: "SE" };
  const DIRECTION_BEARING = { N: 0, NE: 45, E: 90, SE: 135, S: 180, SW: 225, W: 270, NW: 315 };

  function text(value) {
    return String(value ?? "").trim();
  }

  function comparePoleIds(left, right) {
    const first = text(left);
    const second = text(right);
    return first.localeCompare(second, undefined, { numeric: true, sensitivity: "base" })
      || first.localeCompare(second);
  }

  function escapeHtml(value) {
    return String(value ?? "")
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&#039;");
  }

  function splitFiberComponents(value) {
    const source = text(value);
    if (!source) return [];
    const parts = [];
    let depth = 0;
    let start = 0;
    for (let index = 0; index < source.length; index += 1) {
      const character = source[index];
      if (character === "(") depth += 1;
      else if (character === ")") depth = Math.max(0, depth - 1);
      else if ((character === "/" || character === "+") && depth === 0) {
        const part = source.slice(start, index).trim();
        if (part) parts.push(part);
        start = index + 1;
      }
    }
    const last = source.slice(start).trim();
    if (last) parts.push(last);
    return parts;
  }

  function parseFiberComponent(value) {
    const raw = text(value);
    if (!raw) return null;
    const isMst = /\bMST\b/i.test(raw);
    const countMatch = raw.match(/\b(\d+)\s*CT\b/i);
    if (!isMst && !countMatch) return null;
    const directionGroup = Array.from(raw.matchAll(/\(([^)]+)\)/g))
      .map(match => text(match[1]).toUpperCase())
      .find(group => group.split(/[\/,;\s]+/).some(token => DIRECTIONS.includes(token))) || "";
    const directions = directionGroup
      .split(/[\/,;\s]+/)
      .filter(token => DIRECTIONS.includes(token))
      .filter((token, index, tokens) => tokens.indexOf(token) === index);
    const portsMatch = isMst ? raw.match(/\b(8|16)\s*[- ]?PORTS?\b/i) : null;
    return {
      raw,
      fiberType: isMst ? "MST" : `${Number(countMatch[1])}CT Fiber`,
      directions,
      mstPorts: portsMatch ? portsMatch[1] : ""
    };
  }

  function normalizedDirection(value) {
    const direction = text(value).toUpperCase();
    return DIRECTIONS.includes(direction) ? direction : "";
  }

  function directionFromBearing(value) {
    const bearing = Number(value);
    if (!Number.isFinite(bearing)) return "";
    const normalized = ((bearing % 360) + 360) % 360;
    return DIRECTIONS[Math.round(normalized / 45) % 8];
  }

  function directionAtPole(span, poleId) {
    const bearing = Number(span?.bearingDegrees);
    if (Number.isFinite(bearing)) {
      return directionFromBearing(bearing + (span.fromPole === poleId ? 0 : 180));
    }
    const imported = normalizedDirection(span?.direction);
    return span?.fromPole === poleId ? imported : (OPPOSITE[imported] || "");
  }

  function bearingAtPole(span, poleId) {
    const bearing = Number(span?.bearingDegrees);
    if (!Number.isFinite(bearing)) return null;
    return ((bearing + (span.fromPole === poleId ? 0 : 180)) % 360 + 360) % 360;
  }

  function angularDistance(first, second) {
    if (!Number.isFinite(first) || !Number.isFinite(second)) return null;
    const difference = Math.abs(first - second) % 360;
    return Math.min(difference, 360 - difference);
  }

  function relationOf(span) {
    return text(span?.spanRelation || span?.type || span?.rawType || "Other");
  }

  function relationRank(span) {
    const relation = relationOf(span).toLowerCase();
    if (relation.includes("fore")) return 0;
    if (relation.includes("other")) return 1;
    if (relation.includes("back")) return 2;
    return 3;
  }

  function connectedPhysicalSpans(state, poleId) {
    const grouped = new Map();
    Object.values(state?.spans || {}).forEach(span => {
      if (span?.fromPole !== poleId && span?.toPole !== poleId) return;
      if (span.sourceSpanId || /^(manual|additional)-proposed-/i.test(text(span.spanId))) return;
      const otherPoleId = span.fromPole === poleId ? span.toPole : span.fromPole;
      if (!otherPoleId) return;
      const key = [poleId.toUpperCase(), otherPoleId.toUpperCase()].sort().join("|");
      const existing = grouped.get(key);
      if (!existing || relationRank(span) < relationRank(existing)
        || (relationRank(span) === relationRank(existing) && text(span.spanIndex).localeCompare(text(existing.spanIndex), undefined, { numeric: true }) < 0)) {
        grouped.set(key, span);
      }
    });
    return Array.from(grouped.values()).sort((left, right) => {
      const rankDiff = relationRank(left) - relationRank(right);
      if (rankDiff) return rankDiff;
      return text(left.spanIndex).localeCompare(text(right.spanIndex), undefined, { numeric: true })
        || text(left.spanId).localeCompare(text(right.spanId));
    });
  }

  function referenceKey(reference, index) {
    const makeReadyId = text(reference?.makeReadyId);
    const sourceId = makeReadyId || text(reference?.collectionId) || "row";
    return `mr:${sourceId}:${index}`;
  }

  function sizeSettingKey(settings, fiberType) {
    const configured = settings?.fiberSizes && typeof settings.fiberSizes === "object"
      ? settings.fiberSizes
      : {};
    if (fiberType === "MST") {
      return Object.keys(configured).find(key => /^MST(?:\s|$)/i.test(key)) || "MST";
    }
    const count = String(fiberType).match(/\b(\d+)\s*CT\b/i)?.[1] || "";
    return Object.keys(configured).find(key => {
      const keyCount = String(key).match(/\b(\d+)\s*CT\b/i)?.[1] || "";
      return keyCount && Number(keyCount) === Number(count);
    }) || `${Number(count)}CT Fiber`;
  }

  function collectFiberTypes(state) {
    // A saved project can retain diameter settings from a previous job. Those
    // settings must not make unrelated fiber types appear in the current job.
    // The Excel references are the source of truth for the types displayed in
    // the plan. Keep explicitly assigned types as well so an operator's
    // intentional row edit remains selectable after a re-render.
    const values = new Set();
    (state?.makeReadyReferences || []).forEach(reference => {
      splitFiberComponents(reference?.attachmentSizeRaw || reference?.attachmentFiber || "")
        .map(parseFiberComponent)
        .filter(Boolean)
        .forEach(component => values.add(component.fiberType));
    });
    (state?.fiberAssignments || []).forEach(assignment => {
      const fiberType = text(assignment?.fiberType);
      if (fiberType === "MST" || /^(?:\d+)CT Fiber$/i.test(fiberType)) values.add(fiberType);
    });
    return Array.from(values).sort((left, right) => {
      if (left === "MST") return 1;
      if (right === "MST") return -1;
      const leftCount = Number((left.match(/\d+/) || [0])[0]);
      const rightCount = Number((right.match(/\d+/) || [0])[0]);
      return leftCount - rightCount || left.localeCompare(right);
    });
  }

  function installType(value) {
    const normalized = text(value).toUpperCase();
    if (normalized.includes("OVERLASH")) return "OVERLASH";
    if (/\bNEW\b/.test(normalized)) return "NEW";
    return "";
  }

  function buildRows(state) {
    const assignments = Array.isArray(state?.fiberAssignments) ? state.fiberAssignments : [];
    const overrides = new Map(assignments
      .filter(row => row && row.id)
      .map(row => [String(row.id), row]));
    const rows = [];
    (state?.makeReadyReferences || []).forEach((reference, referenceIndex) => {
      const poleId = text(reference?.poleId);
      if (!poleId) return;
      const raw = text(reference?.attachmentSizeRaw || reference?.attachmentFiber);
      const components = splitFiberComponents(raw).map(parseFiberComponent).filter(Boolean);
      if (!components.length) return;
      const spans = connectedPhysicalSpans(state, poleId);
      const sourceKey = referenceKey(reference, referenceIndex);
      const importedDirections = [
        ...(Array.isArray(reference.attachmentDirectionTokens) ? reference.attachmentDirectionTokens : []),
        reference.attachmentDirection
      ].flatMap(value => String(value || "").toUpperCase().match(/\b(?:NE|NW|SE|SW|N|E|S|W)\b/g) || [])
        .filter((direction, index, values) => DIRECTIONS.includes(direction) && values.indexOf(direction) === index);

      components.forEach((component, componentIndex) => {
        const componentDirections = component.directions.length ? component.directions : importedDirections;
        const slots = componentDirections.length
          ? componentDirections.flatMap(direction => {
            const matching = spans.filter(span => directionAtPole(span, poleId) === direction);
            if (matching.length) return matching.map((span, matchIndex) => ({ direction, span, matchIndex }));
            const nearest = spans.map(span => ({ span, distance: angularDistance(bearingAtPole(span, poleId), DIRECTION_BEARING[direction]) }))
              .filter(item => item.distance !== null && item.distance <= 45)
              .sort((left, right) => left.distance - right.distance)[0];
            return nearest
              ? [{ direction, span: nearest.span, matchIndex: 0 }]
              : [{ direction, span: null, matchIndex: 0 }];
          })
          : [{ direction: "", span: spans.length === 1 ? spans[0] : null, matchIndex: 0 }];

        slots.forEach(slot => {
          const baseSpanId = text(slot.span?.spanId);
          // Keep row identity independent from the selected span so changing
          // the destination does not detach the operator's other overrides.
          const id = `${sourceKey}|${componentIndex}|${slot.direction || "any"}|${slot.matchIndex}`;
          const override = overrides.get(id) || {};
          const fiberType = text(override.fiberType || component.fiberType);
          const key = sizeSettingKey(state?.settings, fiberType);
          const portDefault = text(state?.settings?.fiberPlanDefaults?.mstPorts);
          const selectedSpanId = Object.prototype.hasOwnProperty.call(override, "spanId")
            ? text(override.spanId)
            : baseSpanId;
          const selectedSpan = spans.find(span => text(span.spanId) === selectedSpanId) || null;
          rows.push({
            id,
            poleId,
            makeReadyId: text(reference.makeReadyId),
            sourceKey,
            sourceLabel: component.raw,
            componentIndex,
            sourceInstallType: text(reference.attachmentType),
            direction: slot.direction,
            fiberType,
            // Fiber size is configured once per fiber type and exported in
            // sizes.fiber. Do not keep a per-span diameter override.
            diameter: text(state?.settings?.fiberSizes?.[key]),
            mstPorts: Object.prototype.hasOwnProperty.call(override, "mstPorts")
              ? text(override.mstPorts)
              : (component.mstPorts || portDefault),
            installType: Object.prototype.hasOwnProperty.call(override, "installType")
              ? installType(override.installType)
              : installType(reference.attachmentType),
            spanId: selectedSpan ? selectedSpanId : "",
            spanDirection: selectedSpan ? directionAtPole(selectedSpan, poleId) : "",
            spans
          });
        });
      });
    });
    assignments.filter(row => row?.manual === true && row.id).forEach(assignment => {
      const poleId = text(assignment.poleId);
      if (!poleId) return;
      const spans = connectedPhysicalSpans(state, poleId);
      const selectedSpanId = text(assignment.spanId);
      const selectedSpan = spans.find(span => text(span.spanId) === selectedSpanId) || null;
      const fiberType = FIBER_TYPES.includes(text(assignment.fiberType)) ? text(assignment.fiberType) : FIBER_TYPES[0];
      const sizeKey = sizeSettingKey(state?.settings, fiberType);
      rows.push({
        id: text(assignment.id),
        poleId,
        sourceKey: "manual",
        sourceLabel: "Manual assignment",
        componentIndex: Number.MAX_SAFE_INTEGER,
        sourceInstallType: "",
        direction: "",
        fiberType,
        diameter: text(state?.settings?.fiberSizes?.[sizeKey]),
        mstPorts: Object.prototype.hasOwnProperty.call(assignment, "mstPorts")
          ? text(assignment.mstPorts)
          : text(state?.settings?.fiberPlanDefaults?.mstPorts),
        installType: installType(assignment.installType),
        spanId: selectedSpan ? selectedSpanId : "",
        spanDirection: selectedSpan ? directionAtPole(selectedSpan, poleId) : "",
        spans,
        manual: true
      });
    });
    return rows.sort((left, right) => {
      return comparePoleIds(left.poleId, right.poleId)
        || left.componentIndex - right.componentIndex
        || left.direction.localeCompare(right.direction)
        || text(left.spanId).localeCompare(text(right.spanId));
    });
  }

  function spanLabel(state, poleId, span) {
    const otherPoleId = span.fromPole === poleId ? span.toPole : span.fromPole;
    const arrow = span.fromPole === poleId ? "→" : "←";
    const direction = directionAtPole(span, poleId);
    const length = text(span.lengthDisplay || span.length);
    return [
      `${poleId} ${arrow} ${otherPoleId}`,
      relationOf(span),
      direction,
      length
    ].filter(Boolean).join(" · ");
  }

  function create(context) {
    const { els, recordUndoSnapshot, markDirty, S, toast } = context;
    const container = els.fiberPlanResults;
    if (!container) return { render() {} };
    const draft = { poleId: "", spanId: "", fiberType: FIBER_TYPES[0] };
    let manualRowSequence = 0;

    function saveDefault(target) {
      const state = S.getState();
      const field = target.dataset.fiberDefault;
      const value = text(target.value);
      state.settings = state.settings || {};
      if (field === "messenger") {
        state.settings.attachmentMessengerSize = value;
      } else if (field === "mstPorts") {
        state.settings.fiberPlanDefaults = state.settings.fiberPlanDefaults || {};
        state.settings.fiberPlanDefaults[field] = value === "8" || value === "16" ? value : "";
      } else if (field.startsWith("size:")) {
        const fiberType = field.slice("size:".length);
        const key = sizeSettingKey(state.settings, fiberType);
        state.settings.fiberSizes = state.settings.fiberSizes || {};
        state.settings.fiberSizes[key] = value;
      }
    }

    function updateAssignment(target) {
      const state = S.getState();
      const id = target.dataset.fiberAssignment || "";
      const field = target.dataset.fiberField || "";
      if (!id || !field) return;
      const rows = Array.isArray(state.fiberAssignments) ? state.fiberAssignments : [];
      const index = rows.findIndex(row => row?.id === id);
      const existing = index >= 0 ? rows[index] : { id };
      const next = { ...existing, id };
      const value = text(target.value);
      if (field === "fiberType") {
        next.fiberType = value;
        delete next.diameter;
        if (value !== "MST") delete next.mstPorts;
      } else if (field === "mstPorts") {
        next.mstPorts = value === "8" || value === "16" ? value : "";
      } else if (field === "installType") {
        next.installType = value === "NEW" || value === "OVERLASH" ? value : "";
      } else if (field === "spanId") {
        next.spanId = value;
      } else {
        return;
      }
      if (index >= 0) rows[index] = next;
      else rows.push(next);
      state.fiberAssignments = rows;
    }

    function applyToAll() {
      const state = S.getState();
      const rows = buildRows(state);
      if (!rows.length) {
        toast?.("There are no fiber assignments to update.", "warning");
        return;
      }
      const defaults = state.settings?.fiberPlanDefaults || {};
      const ports = text(defaults.mstPorts);
      const assignments = Array.isArray(state.fiberAssignments) ? state.fiberAssignments : [];
      const byId = new Map(assignments.filter(row => row?.id).map(row => [row.id, row]));
      rows.forEach(row => {
        const next = { ...(byId.get(row.id) || {}), id: row.id, poleId: row.poleId, spanId: row.spanId, fiberType: row.fiberType };
        delete next.diameter;
        if (row.fiberType === "MST" && (ports === "8" || ports === "16")) next.mstPorts = ports;
        byId.set(row.id, next);
      });
      state.fiberAssignments = Array.from(byId.values());
      markDirty();
      render();
      toast?.("Global fiber defaults applied. You can still edit individual rows.", "success");
    }

    container.addEventListener("change", event => {
      const target = event.target;
      if (target?.dataset?.fiberNew) {
        const field = target.dataset.fiberNew;
        draft[field] = text(target.value);
        if (field === "poleId") draft.spanId = "";
        render();
        return;
      }
      if (target?.dataset?.fiberDefault) {
        recordUndoSnapshot();
        saveDefault(target);
        markDirty();
        return;
      }
      if (!target?.dataset?.fiberAssignment) return;
      recordUndoSnapshot();
      updateAssignment(target);
      markDirty();
      render();
    });

    container.addEventListener("click", event => {
      const button = event.target?.closest?.("[data-fiber-action]");
      if (!button) return;
      const action = button.dataset.fiberAction;
      if (action === "apply-all") {
        recordUndoSnapshot();
        applyToAll();
        return;
      }
      if (action === "add") {
        const state = S.getState();
        const poleId = text(draft.poleId);
        const span = connectedPhysicalSpans(state, poleId).find(item => text(item.spanId) === text(draft.spanId));
        if (!poleId || !span) {
          toast?.("Select a pole and a connected span first.", "warning");
          return;
        }
        recordUndoSnapshot();
        state.fiberAssignments = Array.isArray(state.fiberAssignments) ? state.fiberAssignments : [];
        const id = `manual-fiber:${Date.now()}:${++manualRowSequence}`;
        state.fiberAssignments.push({
          id,
          manual: true,
          poleId,
          spanId: text(span.spanId),
          fiberType: draft.fiberType,
          mstPorts: draft.fiberType === "MST" ? text(state.settings?.fiberPlanDefaults?.mstPorts) : "",
          installType: ""
        });
        markDirty();
        render();
        toast?.("Fiber assignment added. Complete or adjust its details in the row.", "success");
        return;
      }
      if (action === "remove") {
        const id = button.dataset.fiberAssignment;
        const state = S.getState();
        if (!id) return;
        recordUndoSnapshot();
        state.fiberAssignments = (state.fiberAssignments || []).filter(row => row.id !== id || row.manual !== true);
        markDirty();
        render();
        toast?.("Manual fiber assignment removed.", "success");
      }
    });

    function render() {
      const state = S.getState();
      const settings = state.settings || {};
      const defaults = settings.fiberPlanDefaults || {};
      const types = collectFiberTypes(state);
      const addTypes = types.length ? types : FIBER_TYPES;
      if (!addTypes.includes(draft.fiberType)) draft.fiberType = addTypes[0];
      const rows = buildRows(state);
      const unassignedCount = rows.filter(row => !row.spanId).length;
      const hasMst = rows.some(row => row.fiberType === "MST");
      if (!draft.poleId || !state.poles?.[draft.poleId]) draft.poleId = Object.keys(state.poles || {})[0] || "";
      const newSpans = connectedPhysicalSpans(state, draft.poleId);
      if (!newSpans.some(span => text(span.spanId) === text(draft.spanId))) draft.spanId = text(newSpans[0]?.spanId);
      const poleOptions = Object.keys(state.poles || {}).sort(comparePoleIds).map(poleId =>
        `<option value="${escapeHtml(poleId)}" ${draft.poleId === poleId ? "selected" : ""}>${escapeHtml(poleId)}</option>`
      ).join("");
      const newSpanOptions = newSpans.map(span =>
        `<option value="${escapeHtml(span.spanId)}" ${text(draft.spanId) === text(span.spanId) ? "selected" : ""}>${escapeHtml(spanLabel(state, draft.poleId, span))}</option>`
      ).join("");
      const sizeInputs = types.map(type => {
        const key = sizeSettingKey(settings, type);
        return `<label class="fiber-default-field"><span>${escapeHtml(type === "MST" ? "MST Diameter" : `${type} Size`)}</span><input class="input" type="number" min="0" step="0.001" data-fiber-default="size:${escapeHtml(type)}" value="${escapeHtml(settings.fiberSizes?.[key] || "")}" /></label>`;
      }).join("");
      const rowsPerPole = new Map();
      rows.forEach(row => rowsPerPole.set(row.poleId, (rowsPerPole.get(row.poleId) || 0) + 1));
      let renderedPoleId = "";
      const table = rows.length ? `<div class="table-wrap fiber-plan-table-wrap"><table class="fiber-plan-table">
        <thead><tr><th>POLE</th><th>FIBER</th><th>DIAMETER</th><th>MST PORTS</th><th>INSTALL</th><th>DIRECTION</th><th>SPAN</th><th>SOURCE</th><th>ACTIONS</th></tr></thead>
        <tbody>${rows.map(row => {
          const mst = row.fiberType === "MST";
          const isFirstPoleRow = row.poleId !== renderedPoleId;
          renderedPoleId = row.poleId;
          const poleCell = isFirstPoleRow
            ? `<td class="fiber-pole-cell" rowspan="${rowsPerPole.get(row.poleId)}"><strong>${escapeHtml(row.poleId)}</strong></td>`
            : "";
          const spanOptions = row.spans.map(span => `<option value="${escapeHtml(span.spanId)}" ${row.spanId === span.spanId ? "selected" : ""}>${escapeHtml(spanLabel(state, row.poleId, span))}</option>`).join("");
          const needsAttention = !row.spanId || !row.installType || !row.diameter || (mst && !row.mstPorts);
          const status = !row.spanId ? "Select span" : !row.installType ? "Select New / Overlash" : !row.diameter ? "Set global diameter" : (mst && !row.mstPorts) ? "Select port count" : "Ready";
          return `<tr data-fiber-plan-row="${escapeHtml(row.id)}">
            ${poleCell}
            <td class="fiber-type-cell"><select class="input" data-fiber-assignment="${escapeHtml(row.id)}" data-fiber-field="fiberType">${types.map(type => `<option value="${escapeHtml(type)}" ${row.fiberType === type ? "selected" : ""}>${escapeHtml(type === "MST" ? "MST" : type.replace(/ Fiber$/, ""))}</option>`).join("")}</select></td>
            <td class="fiber-diameter-cell"><span class="fiber-global-size-value">${escapeHtml(row.diameter || "Set globally")}</span></td>
            <td class="fiber-ports-cell">${mst ? `<select class="input" data-fiber-assignment="${escapeHtml(row.id)}" data-fiber-field="mstPorts"><option value="">Select</option><option value="8" ${row.mstPorts === "8" ? "selected" : ""}>8 Port</option><option value="16" ${row.mstPorts === "16" ? "selected" : ""}>16 Port</option></select>` : `<span class="fiber-not-applicable">—</span>`}</td>
            <td class="fiber-install-cell"><select class="input" data-fiber-assignment="${escapeHtml(row.id)}" data-fiber-field="installType"><option value="">Select</option><option value="NEW" ${row.installType === "NEW" ? "selected" : ""}>New</option><option value="OVERLASH" ${row.installType === "OVERLASH" ? "selected" : ""}>Overlash</option></select></td>
            <td class="fiber-direction-cell">${escapeHtml(row.spanDirection || row.direction || "—")}</td>
            <td class="fiber-span-cell"><select class="input" data-fiber-assignment="${escapeHtml(row.id)}" data-fiber-field="spanId"><option value="">Select span</option>${spanOptions}</select></td>
            <td class="fiber-source-column"><div class="fiber-source-cell"><span>${escapeHtml(row.sourceLabel)}</span><small class="${needsAttention ? "fiber-row-status needs-attention" : "fiber-row-status"}">${escapeHtml(status)}</small></div></td>
            <td class="fiber-actions-cell">${row.manual ? `<button class="btn fiber-remove-btn" type="button" title="Remove manual fiber" aria-label="Remove ${escapeHtml(row.fiberType)} on ${escapeHtml(row.poleId)}" data-fiber-action="remove" data-fiber-assignment="${escapeHtml(row.id)}">×</button>` : `<span class="fiber-not-applicable">—</span>`}</td>
          </tr>`;
        }).join("")}</tbody>
      </table></div>` : `<div class="detail-placeholder">No fiber rows yet. Add a fiber below, or import Make Ready fiber references to populate the plan automatically.</div>`;

      container.innerHTML = `<div class="fiber-plan-content">
        <div class="panel-header split fiber-plan-header"><div><p class="eyebrow">Fiber planning</p><h2>Fiber by Pole and Span</h2><p class="muted">Fiber diameters are global by type. Edit each fiber's type, MST ports, installation type, and span. New / Overlash starts from Excel and can be changed per row.</p></div><div class="fiber-plan-header-actions"><div class="fiber-plan-summary"><span>${rows.length} fiber entries</span>${unassignedCount ? `<span class="warning">${unassignedCount} need a span</span>` : ""}</div><div class="fiber-plan-export-action"></div></div></div>
        <section class="fiber-defaults-panel" aria-labelledby="fiberDefaultsTitle"><div class="fiber-section-heading"><h3 id="fiberDefaultsTitle">Global defaults</h3><span>Diameter is shared by every span using that fiber type.</span></div><div class="fiber-defaults-grid">
          <label class="fiber-default-field"><span>Messenger Size</span><input class="input" type="number" min="0" step="0.001" data-fiber-default="messenger" value="${escapeHtml(settings.attachmentMessengerSize || "")}" /></label>
          ${sizeInputs}
          ${hasMst ? `<label class="fiber-default-field"><span>MST ports for Apply All</span><select class="input" data-fiber-default="mstPorts"><option value="">Keep current / select per row</option><option value="8" ${defaults.mstPorts === "8" ? "selected" : ""}>8 Port</option><option value="16" ${defaults.mstPorts === "16" ? "selected" : ""}>16 Port</option></select></label>` : ""}
          <button class="btn btn-primary fiber-apply-all-btn" type="button" data-fiber-action="apply-all">Apply to all rows</button>
        </div></section>
        <section class="fiber-add-panel" aria-labelledby="fiberAddTitle"><div class="fiber-section-heading"><h3 id="fiberAddTitle">Add a fiber</h3><span>Use this when a fiber is not listed in the imported Make Ready data.</span></div><div class="fiber-add-grid">
          <label class="fiber-default-field"><span>Pole</span><select class="input" data-fiber-new="poleId"><option value="">Select pole</option>${poleOptions}</select></label>
          <label class="fiber-default-field"><span>Fiber</span><select class="input" data-fiber-new="fiberType">${addTypes.map(type => `<option value="${escapeHtml(type)}" ${draft.fiberType === type ? "selected" : ""}>${escapeHtml(type === "MST" ? "MST" : type.replace(/ Fiber$/, ""))}</option>`).join("")}</select></label>
          <label class="fiber-default-field"><span>Span</span><select class="input" data-fiber-new="spanId" ${newSpans.length ? "" : "disabled"}><option value="">${newSpans.length ? "Select span" : "No connected span"}</option>${newSpanOptions}</select></label>
          <button class="btn btn-primary" type="button" data-fiber-action="add" ${!draft.poleId || !draft.spanId ? "disabled" : ""}>Add fiber</button>
        </div></section>
        <section class="fiber-assignments-section" aria-labelledby="fiberAssignmentsTitle"><div class="fiber-section-heading"><h3 id="fiberAssignmentsTitle">Pole-by-pole assignments</h3><span>Each row refers to one fiber component and one connected physical span.</span></div>${table}</section>
      </div>`;
      const exportAction = container.querySelector?.(".fiber-plan-export-action");
      if (exportAction && els.exportProposedJsonBtn) exportAction.append(els.exportProposedJsonBtn);
    }

    return { render, buildRows };
  }

  global.FiberTabRenderer = {
    buildRows,
    collectFiberTypes,
    create,
    parseFiberComponent,
    splitFiberComponents
  };
})(window);
