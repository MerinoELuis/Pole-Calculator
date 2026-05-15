(function (global) {
  "use strict";

  const S = global.AppStore;
  const H = global.HeightUtils;

  const els = {};
  let delayedMidspanRenderTimer = null;
  const delayedMidspanRenderPoleIds = new Set();
  const editableInputTimers = new Map();
  const SPAN_COLOR_CLASS_COUNT = 5;


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

  function spanLabel(span) {
    if (!span) return "";
    return `${span.fromPole} → ${span.toPole || "Unknown"}`;
  }

  function shortSpanLabel(span) {
    if (!span) return "";
    return `${span.fromPole || "?"} → ${span.toPole || "Unknown"}`;
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
    return label || "Sin owner";
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

  function proposedSpansForPole(poleId) {
    const seen = new Set();
    return connectedSpansSorted(poleId)
      .filter(span => span.fromPole === poleId)
      .filter(span => spanHasRealMidspan(span.spanId))
      .filter(span => {
        const key = `${span.fromPole || ""}->${span.toPole || ""}`;
        if (seen.has(key)) return false;
        seen.add(key);
        return true;
      });
  }

  function commHasImportedMidspan(sc) {
    return H.parseHeight(sc?.ocalcMS || sc?.midspan || "") !== null;
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
          rows: []
        });
      }
      const group = groups.get(key);
      if (!group.existingHOAChange && sc.existingHOAChange) group.existingHOAChange = sc.existingHOAChange;
      group.rows.push(sc);
    });
    return Array.from(groups.values()).sort((a, b) => {
      // La tabla representa el acomodo físico original del poste, por eso el
      // orden siempre depende de Existing HOA y no de Cambio de HOA.
      const aHeight = H.parseHeight(a.existingHOA || "");
      const bHeight = H.parseHeight(b.existingHOA || "");
      if (aHeight !== null && bHeight !== null && aHeight !== bHeight) return bHeight - aHeight;
      if (aHeight !== null && bHeight === null) return -1;
      if (aHeight === null && bHeight !== null) return 1;
      return `${a.owner}${a.existingHOA}`.localeCompare(`${b.owner}${b.existingHOA}`, undefined, { numeric: true });
    });
  }

  function commMidspanEntries(group, poleId) {
    // Un comm puede tocar más de un span. Las columnas Span, Midspan y
    // Otro poste HOA salen de esta misma lista para conservar la alineación
    // visual y la relación fila a fila entre los tres datos.
    const seen = new Set();
    const entries = [];
    const rows = [...group.rows].sort((a, b) =>
      `${a.spanId}${a.wireIndex || ""}`.localeCompare(`${b.spanId}${b.wireIndex || ""}`, undefined, { numeric: true })
    );
    rows.forEach(sc => {
      const span = S.getSpan(sc.spanId);
      const hasMidspan = commHasImportedMidspan(sc);
      const midspan = hasMidspan ? displayMidspan(sc) : "";
      const key = `${sc.spanId}|${midspan || ""}|${hasMidspan ? "ms" : "ref"}`;
      if (seen.has(key)) return;
      seen.add(key);
      const isBackspan = span && span.fromPole !== poleId;
      const remote = global.Calculations.findRemoteComm(sc.spanId, sc.poleId, sc.ownerBase || sc.owner, sc.wireId || "");
      const midspanLocked = Boolean(sc.existingHOAChange || remote?.existingHOAChange);
      entries.push({
        spanHtml: `<div class="comm-span-row">
          ${span ? spanColorDot(poleId, span.spanId) : ""}
          <span>${span ? `${poleLink(span.fromPole)} → ${poleLink(span.toPole)}` : escapeHtml(sc.spanId || "")}</span>
          ${hasMidspan && isBackspan ? `<em>ref</em>` : ""}
          ${!hasMidspan ? `<em>REF</em>` : ""}
        </div>`,
        midspanHtml: `<div class="comm-midspan-value">${hasMidspan && !midspanLocked
          ? `<input class="input height-input remote-height-input" data-scope="spanComm" data-pole="${escapeHtml(sc.poleId)}" data-span="${escapeHtml(sc.spanId)}" data-owner="${escapeHtml(sc.owner)}" data-wire-id="${escapeHtml(sc.wireId || "")}" data-field="midspan" value="${escapeHtml(midspan)}">`
          : `<strong>${escapeHtml(midspan)}</strong>`}</div>`,
        remoteHtml: !hasMidspan || !remote
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

  function renderCommMidspanValues(group, poleId) {
    const entries = commMidspanEntries(group, poleId);
    return `<div class="comm-midspan-list">${entries.map(entry => entry.midspanHtml).join("")}</div>`;
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
    const template = document.createElement("template");
    template.innerHTML = renderPoleWorkspace(poleId).trim();
    const nextCard = template.content.firstElementChild;
    if (!nextCard) return false;
    const parent = oldCard.parentNode;
    if (!parent || !parent.contains(oldCard)) return false;
    parent.replaceChild(nextCard, oldCard);
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

  function renderEditableNotes(scope, attrs, value, autoNotes, placeholder = "Notas...") {
    const attrText = Object.entries(attrs || {})
      .map(([name, attrValue]) => `data-${name}="${escapeHtml(attrValue)}"`)
      .join(" ");
    return `<div class="notes-cell">
      <textarea class="input text-input" data-scope="${escapeHtml(scope)}" ${attrText} data-field="notes" placeholder="${escapeHtml(placeholder)}">${escapeHtml(value || "")}</textarea>
    </div>`;
  }

  function commMidspanNote(row, span) {
    if (!span || !span.midspanLowPower) return "Falta Power MS para calcular Max MS Comm.";
    const midspan = typeof row === "string" ? row : displayMidspan(row);
    if (!midspan) return span.midspanMaxCommHeight ? `Falta midspan. Max ${span.midspanMaxCommHeight}.` : "Falta midspan.";
    return row?.clearanceMSMessage || `Max ${span.midspanMaxCommHeight} · Low Power MS ${span.midspanLowPower}.`;
  }

  function poleClearanceNote(height, pole, missingLabel = "Missing Data") {
    if (!pole || !pole.lowPower) return missingLabel;
    const max = pole.maxCommHeight || "";
    if (!height) return max ? `Falta dato. Max ${max}.` : "Falta dato.";
    const aboveMax = max && H.compareHeights(height, max) === 1;
    return aboveMax
      ? `Clearance issue. Max ${max} · Low Power ${pole.lowPower}.`
      : `OK. Max ${max} · Low Power ${pole.lowPower}.`;
  }

  function spanSideClearanceNote(side) {
    if (!side.ocalcMS && !side.proposedMidspan) return "Falta O-CALC MS.";
    if (side.clearanceMSReason === "LOW_POWER" && side.clearanceMSIssue) return `Ensure min 30" to low power at midspan.`;
    if (side.clearanceMSStatus === "PENDING") {
      return `${side.clearanceMSMessage || "Clearance issue."} Corrigiendo a ${side.pendingMidspanFinal || ""}...`;
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
    if (status === "ADJUSTED") return `<span class="badge warning">Ajustado</span>`;
    if (status === "ADJUSTMENT_NEEDED") return `<span class="badge warning">Ajuste requerido</span>`;
    return `<span class="badge changed">OK</span>`;
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
    const clearanceRows = [
      ["polePowerCommsClearance", "Pole · Power-comms", settings.polePowerCommsClearance || settings.clearanceToPower || "40\""],
      ["commClearance", "Pole · Comm-comm", settings.commClearance || "12\""],
      ["boltClearance", "Pole · Bolt-bolt", settings.boltClearance || "4\""],
      ["midspanPowerCommClearance", "Midspan · Power-comm", settings.midspanPowerCommClearance || "30\""],
      ["midspanCommCommClearance", "Midspan · Comm-comm", settings.midspanCommCommClearance || "4\""]
    ];
    const position = settings.position === "LOW_COMM" ? "LOW_COMM" : "TOP_COMM";
    const proposedOwner = settings.proposedOwner || "Wecom";
    const ownerOptions = ["Wecom", "CenturyLink", "Cable One", "Cox", "Fatbeam", "Vexus", "MCI Metro"].map(owner =>
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
        <h3>Ajustes</h3>
        <div class="settings-grid adjustments-grid">
          <label class="clearance-row position-row">
            <span>Posición</span>
            <select class="input position-select" data-scope="settings" data-field="position">
              <option value="TOP_COMM" ${position === "TOP_COMM" ? "selected" : ""}>Top Comm</option>
              <option value="LOW_COMM" ${position === "LOW_COMM" ? "selected" : ""}>Low Comm</option>
            </select>
          </label>
          <label class="clearance-row position-row">
            <span>Make Ready</span>
            <select class="input position-select" data-scope="settings" data-field="mrCase">
              <option value="LOWER" ${(settings.mrCase || "LOWER") === "LOWER" ? "selected" : ""}>Minúsculas</option>
              <option value="UPPER" ${settings.mrCase === "UPPER" ? "selected" : ""}>Mayúsculas</option>
            </select>
          </label>
          <label class="clearance-row position-row">
            <span>Owner del Proposed</span>
            <select class="input position-select" data-scope="settings" data-field="proposedOwner">
              ${ownerOptions}
            </select>
          </label>
        </div>
      </div>
    `;
    wireEditableEvents(els.clearanceSettings);
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

  function renderSummary() {
    const state = S.getState();
    if (els.projectMeta) {
      els.projectMeta.textContent = `${state.importedFileName || "Proyecto"} · ${new Date(state.importedAt || Date.now()).toLocaleString()}`;
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
      if (filter === "warnings" && summary.warnings.length === 0) return false;
      if (filter === "changed" && !summary.hasChanges) return false;
      return true;
    }).sort((a, b) => a.localeCompare(b, undefined, { numeric: true }));
  }

  function renderPoleListItem(poleId) {
    const state = S.getState();
    const { pole, warnings, hasChanges } = poleSummary(poleId);
    const active = state.selectedPoleId === poleId ? " active" : "";
    return `<button class="pole-index-link${active}" data-pole-select="${escapeHtml(poleId)}" type="button">
      <span>${escapeHtml(poleId)}</span>
      ${pole.isGenerated ? `<span class="mini-dot warning">Gen</span>` : ""}
      ${warnings.length ? `<span class="mini-dot danger">${warnings.length}</span>` : ""}
      ${hasChanges ? `<span class="mini-dot changed">Cambio</span>` : ""}
    </button>`;
  }

  function renderPoleLists() {
    const poleIds = filteredPoles();
    els.polesList.innerHTML = poleIds.map(id => renderPoleListItem(id)).join("") || `<div class="detail-placeholder">No hay postes con ese filtro.</div>`;
    els.polesList.querySelectorAll("[data-pole-select]").forEach(btn => {
      btn.addEventListener("click", () => selectPole(btn.dataset.poleSelect));
    });
  }

  function renderWarningsList(poleId) {
    const warnings = S.getState().warnings.filter(w => w.poleId === poleId);
    if (!warnings.length) return `<p class="muted">Sin warnings.</p>`;
    return `<ul class="warning-list">${warnings.map(w => `<li class="warning-item"><strong>${escapeHtml(w.code)}</strong><br><span>${escapeHtml(w.message)}</span><div class="muted">${escapeHtml(w.spanId || "Sin span")} ${w.owner ? "· " + escapeHtml(w.owner) : ""}</div></li>`).join("")}</ul>`;
  }

  function renderMRText(poleId) {
    const item = S.getState().mr.find(mr => mr.poleId === poleId);
    if (!item) return `<p class="muted">Sin MR generado todavía.</p>`;
    return `<pre class="mr-output">${escapeHtml(item.text)}</pre>`;
  }

  function renderPoleEditableHeader(poleId) {
    const { pole, spans, warnings, hasChanges } = poleSummary(poleId);
    const maxHeightAtMS = spans
      .map(span => H.parseHeight(span.midspanMaxCommHeight || ""))
      .filter(value => value !== null);
    const restrictiveMaxHeightAtMS = maxHeightAtMS.length ? H.formatHeight(Math.min(...maxHeightAtMS)) : "";
    return `<div class="pole-workspace-header">
      <div>
        <h3 id="pole-${escapeHtml(poleId)}">${escapeHtml(poleId)}</h3>
        <div class="pole-meta">
          ${pole.isGenerated ? `<span class="badge warning">Other pole generado editable</span>` : ""}
          <span class="badge">Spans ${spans.length}</span>
          <span class="badge owner">Comms ${S.getSpanCommsForPole(poleId).length}</span>
          ${warnings.length ? `<span class="badge warning">Warnings ${warnings.length}</span>` : `<span class="badge">Warnings 0</span>`}
          ${hasChanges ? `<span class="badge changed">Con cambios</span>` : ""}
        </div>
      </div>
      <div class="pole-kpis three-up">
        <label>Low Power on Pole<input class="input height-input" data-scope="pole" data-pole="${escapeHtml(poleId)}" data-field="lowPower" value="${escapeHtml(pole.lowPower || "")}" placeholder="30'8&quot;"></label>
        <label>Max Height on Pole<input class="input height-input muted-input" value="${escapeHtml(pole.maxCommHeight || "")}" readonly></label>
        <label>Max Height at MS<input class="input height-input muted-input" value="${escapeHtml(restrictiveMaxHeightAtMS)}" readonly></label>
      </div>
    </div>`;
  }

  function renderSpanProposedTable(poleId) {
    const spans = proposedSpansForPole(poleId);
    if (!spans.length) return `<p class="muted">No hay spans con midspan para proponer desde este poste.</p>`;
    return `<div class="table-wrap"><table class="span-proposed-table wide-table">
      <thead><tr>
        <th>Span</th><th>Environment</th><th>Environment Clearance</th><th>Proposed</th><th>End Drop</th><th>Next Pole Proposed</th><th>O-CALC MS</th><th>MS Proposed</th><th>MS Proposed Clearance</th><th>Adjusted Final MS</th><th>Notes</th>
      </tr></thead>
      <tbody>${spans.map(span => {
        const side = S.getSpanSide(span.spanId, poleId) || S.upsertSpanSide({ spanId: span.spanId, poleId });
        const pole = S.getPole(poleId);
        const aboveMax = side.proposedHOA && H.compareHeights(side.proposedHOA, side.maxCommHeight || pole?.maxCommHeight) === 1;
        const midspanIssue = side.clearanceMSStatus === "PENDING" || side.clearanceMSStatus === "PROBLEM";
        const boltIssue = global.Calculations.evaluateProposedBoltClearance(side);
        const rowClasses = [
          spanRowClasses(poleId, span.spanId),
          side.proposedHOA || side.ocalcMS || side.msProposed || side.finalMidspan || side.proposedMidspan || side.endDrop ? "changed-row" : "",
          aboveMax || midspanIssue || !boltIssue.ok ? "warning-row" : ""
        ].filter(Boolean).join(" ");
        const autoNotes = [spanSideClearanceNote(side), boltIssue.message];
        return `<tr class="${rowClasses}">
          <td class="span-cell"><strong>${spanChip(poleId, span.spanId)}${poleLink(span.fromPole)} → ${poleLink(span.toPole)}</strong></td>
          <td><select class="input environment-input" data-scope="span" data-span="${escapeHtml(span.spanId)}" data-field="environment">${renderEnvironmentOptions(span.environment)}</select></td>
          <td><input class="input" data-scope="span" data-span="${escapeHtml(span.spanId)}" data-field="environmentClearance" value="${escapeHtml(span.environmentClearance || "")}"></td>
          <td><input class="input height-input" data-scope="spanSide" data-pole="${escapeHtml(poleId)}" data-span="${escapeHtml(span.spanId)}" data-field="proposedHOA" value="${escapeHtml(side.proposedHOA || "")}"></td>
          <td><span class="calculated-value">${escapeHtml(side.endDrop || "")}</span></td>
          <td><input class="input height-input" data-scope="spanSide" data-pole="${escapeHtml(poleId)}" data-span="${escapeHtml(span.spanId)}" data-field="proposedHOAChange" value="${escapeHtml(side.proposedHOAChange || "")}"></td>
          <td><input class="input decimal-height-input" data-scope="spanSide" data-pole="${escapeHtml(poleId)}" data-span="${escapeHtml(span.spanId)}" data-field="ocalcMS" value="${escapeHtml(displayDecimalFeetInput(side.ocalcMS, side.proposedMidspan))}" placeholder="XX.XX"></td>
          <td><span class="calculated-value">${escapeHtml(side.msProposed || "")}</span></td>
          <td>${renderSpanSideMidspanStatus(side)}</td>
          <td><span class="calculated-value">${escapeHtml(side.finalMidspan || "")}</span></td>
          <td>${renderEditableNotes("spanSide", { pole: poleId, span: span.spanId }, side.notes, autoNotes, "Notas propias...")}</td>
        </tr>`;
      }).join("")}</tbody>
    </table></div>`;
  }

  function renderCommMovementTable(poleId) {
    const groups = groupedCommsForPole(poleId);
    if (!groups.length) return `<p class="muted">No hay comms importados desde Span.Wire para este poste.</p>`;
    return `<div class="table-wrap"><table class="comm-movement-table">
      <thead><tr>
        <th>Owner/Comm</th><th>Service Drop</th><th>Existing HOA</th><th>HOA Change</th><th>Span</th><th>Midspan</th><th>Other Pole HOA</th><th>Flagging</th><th>Actions</th>
      </tr></thead>
      <tbody>${groups.map(group => {
        const pole = S.getPole(poleId);
        const effective = group.existingHOAChange || group.existingHOA;
        const aboveMax = effective && H.compareHeights(effective, pole?.maxCommHeight) === 1;
        const changed = Boolean(group.existingHOAChange || group.rows.some(row => row.mr));
        const flaggingIssue = group.rows.some(row => row.flaggingStatus === "PROBLEM");
        const flaggingMissing = group.rows.some(row => row.flaggingStatus === "MISSING" || row.flaggingStatus === "MISSING_POWER");
        const rowClasses = [
          changed ? "changed-row" : "",
          aboveMax || flaggingIssue || flaggingMissing ? "warning-row" : ""
        ].filter(Boolean).join(" ");
        return `<tr class="${rowClasses}">
          <td><span class="badge owner">${escapeHtml(group.owner)}</span></td>
          <td><input type="checkbox" data-scope="commGroup" data-pole="${escapeHtml(poleId)}" data-group-key="${escapeHtml(group.key)}" data-field="serviceDrop" ${group.rows.some(row => row.serviceDrop) ? "checked" : ""}></td>
          <td>${escapeHtml(group.existingHOA || "")}</td>
          <td><input class="input height-input" data-scope="commGroup" data-pole="${escapeHtml(poleId)}" data-group-key="${escapeHtml(group.key)}" data-field="existingHOAChange" value="${escapeHtml(group.existingHOAChange || "")}"></td>
          <td>${renderCommSpanRefs(group, poleId)}</td>
          <td>${renderCommMidspanValues(group, poleId)}</td>
          <td>${renderCommRemoteValues(group)}</td>
          <td>${renderCommFlagging(group)}</td>
          <td><div class="row-actions">
            <button class="icon-action" type="button" data-edit-comm data-pole="${escapeHtml(poleId)}" data-group-key="${escapeHtml(group.key)}" title="Editar comm" aria-label="Editar comm">&#9998;</button>
            <button class="icon-action" type="button" data-edit-comm-spans data-pole="${escapeHtml(poleId)}" data-group-key="${escapeHtml(group.key)}" title="Editar spans del comm" aria-label="Editar spans del comm">&#8644;</button>
            <button class="icon-action danger-action" type="button" data-delete-comm data-pole="${escapeHtml(poleId)}" data-group-key="${escapeHtml(group.key)}" title="Borrar comm" aria-label="Borrar comm">&#128465;</button>
          </div></td>
        </tr>`;
      }).join("")}</tbody>
    </table></div>`;
  }

  function renderPowerTable(poleId) {
    const rows = S.getSpanPowerForPole(poleId).filter(row => H.parseHeight(row.midspan) !== null);
    if (!rows.length) return `<p class="muted">No se importaron wires de power para este poste.</p>`;
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
    return `<article class="pole-workspace-card" data-pole-card="${escapeHtml(poleId)}">
      ${renderPoleEditableHeader(poleId)}
      <div class="workspace-grid">
        <section class="subsection wide" id="spans-${escapeHtml(poleId)}">
          <h4>Proposed by Span</h4>
          ${renderSpanProposedTable(poleId)}
        </section>
        <section class="subsection wide" id="comms-${escapeHtml(poleId)}">
          <div class="subsection-title-row">
            <h4>Existing Comm Movements</h4>
            <button class="mini-btn" type="button" data-add-comm data-pole="${escapeHtml(poleId)}">Add Comm</button>
          </div>
          <p class="muted">Aquí se mueve cada comm existente con nueva altura. Si cambia el otro poste del mismo span, el Midspan calculado se actualiza.</p>
          ${renderCommMovementTable(poleId)}
        </section>
        <section class="subsection wide">
          <h4>Imported Power / Clearance</h4>
          ${renderPowerTable(poleId)}
        </section>
        <section class="subsection" id="warnings-${escapeHtml(poleId)}">
          <h4>Make Ready</h4>
          ${renderMRText(poleId)}
        </section>
        <section class="subsection">
          <h4>Warnings</h4>
          ${renderWarningsList(poleId)}
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
        console.error(`Error renderizando poste ${id}`, error);
        return `<article class="pole-workspace-card warning-row" data-pole-card="${escapeHtml(id)}">
          <div class="pole-workspace-header">
            <div>
              <h3>${escapeHtml(id)}</h3>
              <span class="badge danger">Error de datos</span>
            </div>
          </div>
          <p class="muted">Este poste se importo, pero tiene datos incompletos o inesperados. Revisa el Excel o vuelve a importar.</p>
        </article>`;
      }
    }).join("") || `<div class="detail-placeholder">No hay postes para mostrar.</div>`;
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
    root.querySelectorAll("[data-delete-comm]").forEach(btn => btn.addEventListener("click", () => deleteCommGroup(btn.dataset.pole, btn.dataset.groupKey)));
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
      "position",
      "mrCase",
      "proposedOwner"
    ].includes(field);
  }

  function updateCommGroupField(poleId, groupKey, field, value) {
    if (!["existingHOAChange", "serviceDrop"].includes(field)) return [poleId].filter(Boolean);
    const affected = new Set([poleId].filter(Boolean));
    S.getSpanCommsForPole(poleId)
      .filter(sc => commGroupKey(sc) === groupKey)
      .forEach(sc => {
        global.Calculations.updateSpanCommField(sc.spanId, sc.poleId, sc.owner, sc.wireId || "", field, value);
        poleIdsForSpan(sc.spanId).forEach(id => affected.add(id));
      });
    return Array.from(affected);
  }

  function defaultSpanForNewComm(poleId) {
    return connectedSpansSorted(poleId)[0] || null;
  }

  function addCommToPole(poleId) {
    const owner = prompt("Owner/Comm");
    if (!owner) return;
    const existingHOA = prompt("Existing HOA", "") || "";
    const span = defaultSpanForNewComm(poleId);
    if (!span) return toast("Ese poste no tiene spans para asociar el comm.", "warning");
    const wireId = `manual-${Date.now()}`;
    S.upsertComm(poleId, owner, existingHOA, "", { ownerBase: owner, rawOwner: owner, wireId });
    S.upsertSpanComm({ spanId: span.spanId, poleId, owner, ownerBase: owner, rawOwner: owner, existingHOA, wireId });
    global.Calculations.recalculateSpansForPole(poleId);
    renderAffectedPoles([poleId, S.getOtherPoleId(span, poleId)]);
  }

  function editCommGroup(poleId, groupKey) {
    const group = groupedCommsForPole(poleId).find(item => item.key === groupKey);
    if (!group) return;
    const nextOwner = prompt("Owner/Comm", group.owner);
    if (!nextOwner) return;
    const nextHOA = prompt("Existing HOA", group.existingHOA || "") || "";
    group.rows.forEach(row => {
      S.removeSpanComm(row.spanId, row.poleId, row.owner, row.wireId || "");
      S.upsertComm(poleId, nextOwner, nextHOA, "", { ownerBase: nextOwner, rawOwner: nextOwner, wireId: row.wireId || "" });
      S.upsertSpanComm({ ...row, owner: nextOwner, ownerBase: nextOwner, rawOwner: nextOwner, existingHOA: nextHOA });
    });
    global.Calculations.recalculateSpansForPole(poleId);
    renderAffectedPoles([poleId]);
  }

  function editCommSpans(poleId, groupKey) {
    const group = groupedCommsForPole(poleId).find(item => item.key === groupKey);
    if (!group) return;
    const spans = connectedSpansSorted(poleId);
    if (!spans.length) return toast("Ese poste no tiene spans disponibles.", "warning");
    const menu = spans.map((span, index) => `${index + 1}: ${shortSpanLabel(span)}`).join("\n");
    const currentIndexes = group.rows
      .map(row => spans.findIndex(span => span.spanId === row.spanId))
      .filter(index => index >= 0)
      .map(index => index + 1)
      .join(",");
    const answer = prompt(`Elige spans para ${group.owner} (separa con comas):\n${menu}`, currentIndexes);
    if (answer === null) return;
    const selected = new Set(answer.split(",")
      .map(value => Number(value.trim()) - 1)
      .filter(index => Number.isInteger(index) && spans[index])
      .map(index => spans[index].spanId));
    if (!selected.size) return;

    group.rows.forEach(row => {
      if (!selected.has(row.spanId)) S.removeSpanComm(row.spanId, row.poleId, row.owner, row.wireId || "");
    });

    spans.filter(span => selected.has(span.spanId)).forEach(span => {
      const existing = group.rows.find(row => row.spanId === span.spanId);
      const midspan = prompt(`Midspan para ${shortSpanLabel(span)} (deja vacío si solo es REF)`, existing?.midspan || "") ?? (existing?.midspan || "");
      if (existing) {
        global.Calculations.updateSpanCommField(existing.spanId, existing.poleId, existing.owner, existing.wireId || "", "midspan", midspan);
        return;
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
        midspan,
        wireId
      });
    });
    global.Calculations.recalculateSpansForPole(poleId);
    renderAffectedPoles([poleId]);
  }

  function deleteCommGroup(poleId, groupKey) {
    const group = groupedCommsForPole(poleId).find(item => item.key === groupKey);
    if (!group || !confirm(`Borrar ${group.owner}?`)) return;
    group.rows.forEach(row => S.removeSpanComm(row.spanId, row.poleId, row.owner, row.wireId || ""));
    global.Calculations.recalculateSpansForPole(poleId);
    renderAffectedPoles([poleId]);
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

    if (scope === "pole") {
      S.updatePoleField(el.dataset.pole, field, value);
      global.Calculations.recalculateSpansForPole(el.dataset.pole);
    }

    if (scope === "spanSide") {
      global.Calculations.updateSpanSideField(el.dataset.span, el.dataset.pole, field, value);
    }

    if (scope === "span") {
      if (field === "environment") {
        const option = (S.ENVIRONMENT_OPTIONS || []).find(item => item.value === value);
        global.Calculations.updateSpanField(el.dataset.span, "environment", value);
        global.Calculations.updateSpanField(el.dataset.span, "environmentClearance", option?.clearance || "");
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

    const affectedPoleIds = affectedPoleIdsForElement(el, scope);
    if (scope !== "settings" && isLiveRecalcField(field)) {
      affectedPoleIds.forEach(id => global.Calculations.recalculateSpansForPole(id));
    }
    if (scope === "settings") render();
    else renderAffectedPoles(affectedPoleIds);

    if (["lowPower", "ocalcMS", "proposedMidspan", "proposedHOA", "proposedHOAChange", "existingHOA", "existingHOAChange", "midspan", "environmentClearance", "midspanCommCommClearance", "midspanPowerCommClearance", "polePowerCommsClearance", "clearanceToPower", "position", "proposedOwner"].includes(field)) {
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
    render();
    scrollToPole(poleId);
  }

  function render() {
    global.Calculations.recalculateAll();
    renderSummary();
    renderClearanceSettings();
    renderPoleLists();
    renderAllPolesWorkspace();
  }

  async function handleExcelImport(file) {
    if (!file) return;
    try {
      await global.ExcelImport.importExcelFile(file);
      render();
      toast("Excel crudo importado. Se cargaron las hojas disponibles.", "success");
    } catch (error) {
      console.error(error);
      toast(`Error importando Excel: ${error.message}`, "error");
    } finally {
      els.excelFileInput.value = "";
    }
  }

  async function handleJsonImport(file) {
    if (!file) return;
    try {
      await global.ExcelImport.importJsonFile(file);
      render();
      toast("JSON importado. El trabajo guardado quedo cargado.", "success");
    } catch (error) {
      console.error(error);
      toast(`Error importando JSON: ${error.message}`, "error");
    } finally {
      els.jsonFileInput.value = "";
    }
  }

  function bindEvents() {
    els.excelFileInput.addEventListener("change", event => handleExcelImport(event.target.files[0]));
    els.jsonFileInput.addEventListener("change", event => handleJsonImport(event.target.files[0]));
    els.exportJsonBtn.addEventListener("click", () => global.ProjectExport.exportJson());
    els.saveLocalBtn.addEventListener("click", () => { S.saveToLocal(); toast("Guardado local en este navegador.", "success"); });
    els.loadLocalBtn.addEventListener("click", () => {
      const loaded = S.loadFromLocal();
      if (!loaded) return toast("No hay guardado local.", "warning");
      render();
      toast("Guardado local cargado.", "success");
    });
    els.resetSampleBtn.addEventListener("click", () => { S.loadSampleData(); render(); toast("Datos demo cargados.", "info"); });
    els.poleSearchInput.addEventListener("input", event => { S.getState().ui.search = event.target.value; render(); });
    els.warningFilterSelect.addEventListener("change", event => { S.getState().ui.filter = event.target.value; render(); });
  }

  function init() {
    Object.assign(els, {
      excelFileInput: qs("excelFileInput"),
      jsonFileInput: qs("jsonFileInput"),
      exportJsonBtn: qs("exportJsonBtn"),
      saveLocalBtn: qs("saveLocalBtn"),
      loadLocalBtn: qs("loadLocalBtn"),
      resetSampleBtn: qs("resetSampleBtn"),
      projectMeta: qs("projectMeta"),
      poleSearchInput: qs("poleSearchInput"),
      warningFilterSelect: qs("warningFilterSelect"),
      polesList: qs("polesList"),
      polesOverview: qs("polesOverview"),
      appLayout: qs("appLayout"),
      clearanceSettings: qs("clearanceSettings"),
      toastHost: qs("toastHost")
    });

    bindEvents();
    global.FloatingCalculator?.setupFloatingCalculator();
    S.loadSampleData();
    render();
  }

  document.addEventListener("DOMContentLoaded", init);
})(window);
