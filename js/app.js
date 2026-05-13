(function (global) {
  "use strict";

  const S = global.AppStore;
  const H = global.HeightUtils;

  const els = {};

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

  function poleSummary(poleId) {
    const state = S.getState();
    const pole = S.getPole(poleId);
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
    const poles = Object.keys(state.poles);
    const warnings = state.warnings.length;
    const changed = poles.filter(S.poleHasChanges).length;
    const midspans = Object.values(state.spanComms).filter(sc => sc.calculatedMidspan || sc.midspan || sc.ocalcMS).length;
    els.summaryGrid.innerHTML = [
      [poles.length, "Postes"],
      [Object.keys(state.spans).length, "Spans"],
      [midspans, "Midspans"],
      [warnings, "Warnings"],
      [changed, "Con cambios"]
    ].map(([value, label]) => `<div class="summary-card"><strong>${value}</strong><span>${label}</span></div>`).join("");
    els.projectMeta.textContent = `${state.importedFileName || "Proyecto"} · ${new Date(state.importedAt || Date.now()).toLocaleString()}`;
  }

  function filteredPoles() {
    const state = S.getState();
    const search = (state.ui.search || "").toLowerCase();
    const filter = state.ui.filter || "all";
    return Object.keys(state.poles).filter(poleId => {
      const summary = poleSummary(poleId);
      const pole = summary.pole;
      const text = `${poleId} ${pole.sequence || ""} ${pole.poleHeight} ${pole.lowPower || ""} ${pole.maxCommHeight || ""} ${pole.comms.map(c => `${c.owner} ${c.existingHOA}`).join(" ")}`.toLowerCase();
      if (search && !text.includes(search)) return false;
      if (filter === "warnings" && summary.warnings.length === 0) return false;
      if (filter === "changed" && !summary.hasChanges) return false;
      return true;
    }).sort((a, b) => a.localeCompare(b, undefined, { numeric: true }));
  }

  function renderPoleListItem(poleId) {
    const state = S.getState();
    const { pole, spans, warnings, midspanCount, hasChanges } = poleSummary(poleId);
    const active = state.selectedPoleId === poleId ? " active" : "";
    const generated = pole.isGenerated ? `<span class="badge warning">Editable generado</span>` : "";
    const badges = [
      generated,
      `<span class="badge">Height ${escapeHtml(pole.poleHeight || "-")}</span>`,
      pole.lowPower ? `<span class="badge">Low Power ${escapeHtml(pole.lowPower)}</span>` : `<span class="badge warning">Sin Low Power</span>`,
      pole.maxCommHeight ? `<span class="badge">Max ${escapeHtml(pole.maxCommHeight)}</span>` : "",
      `<span class="badge">Spans ${spans.length}</span>`,
      `<span class="badge owner">Comms ${S.getSpanCommsForPole(poleId).length}</span>`,
      `<span class="badge">MS ${midspanCount}</span>`,
      warnings.length ? `<span class="badge warning">Warnings ${warnings.length}</span>` : `<span class="badge">Warnings 0</span>`,
      hasChanges ? `<span class="badge changed">Cambios</span>` : ""
    ].join("");
    return `<button class="pole-list-item${active}" data-pole-select="${escapeHtml(poleId)}" type="button">
      <div class="pole-list-header">
        <h3>${escapeHtml(poleId)}</h3>
        ${warnings.length ? `<span class="badge warning">${warnings.length}</span>` : ""}
      </div>
      <div class="pole-meta">${badges}</div>
    </button>`;
  }

  function renderPoleLists() {
    const poleIds = filteredPoles();
    els.polesList.innerHTML = poleIds.map(id => renderPoleListItem(id)).join("") || `<div class="detail-placeholder">No hay postes con ese filtro.</div>`;
    document.querySelectorAll("[data-pole-select]").forEach(btn => {
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
      <div class="pole-kpis">
        <label>Altura poste<input class="input height-input" data-scope="pole" data-pole="${escapeHtml(poleId)}" data-field="poleHeight" value="${escapeHtml(pole.poleHeight || "")}" placeholder="35'"></label>
        <label>Low Power<input class="input height-input" data-scope="pole" data-pole="${escapeHtml(poleId)}" data-field="lowPower" value="${escapeHtml(pole.lowPower || "")}" placeholder="30'8&quot;"></label>
        <label>Altura Max<input class="input height-input muted-input" value="${escapeHtml(pole.maxCommHeight || "")}" readonly></label>
        <label>Top Comm<input class="input height-input muted-input" value="${escapeHtml(pole.topComm || "")}" readonly></label>
        <label>Low Comm<input class="input height-input muted-input" value="${escapeHtml(pole.lowComm || "")}" readonly></label>
      </div>
    </div>`;
  }

  function renderSpanProposedTable(poleId) {
    const spans = S.getConnectedSpans(poleId).sort((a, b) => String(a.spanIndex || a.spanId).localeCompare(String(b.spanIndex || b.spanId), undefined, { numeric: true }));
    if (!spans.length) return `<p class="muted">No hay spans conectados.</p>`;
    return `<div class="table-wrap"><table class="span-proposed-table">
      <thead><tr>
        <th>Span</th><th>Dir</th><th>Length</th><th>Other Pole</th><th>Low Power</th><th>Altura Max</th><th>Proposed HOA</th><th>Proposed Midspan / O-Calc MS</th><th>End Drop</th><th>Referencia clearance</th><th>Notas</th>
      </tr></thead>
      <tbody>${spans.map(span => {
        const side = S.getSpanSide(span.spanId, poleId) || S.upsertSpanSide({ spanId: span.spanId, poleId });
        const other = S.getOtherPoleId(span, poleId);
        const pole = S.getPole(poleId);
        const aboveMax = side.proposedHOA && H.compareHeights(side.proposedHOA, side.maxCommHeight || pole.maxCommHeight) === 1;
        return `<tr class="${side.proposedHOA || side.proposedMidspan || side.endDrop ? "changed-row" : ""} ${aboveMax ? "warning-row" : ""}">
          <td><strong>${escapeHtml(spanLabel(span))}</strong><div class="muted">${escapeHtml(span.spanId)} · ${escapeHtml(span.type || "")}</div></td>
          <td><span class="badge">${escapeHtml(span.direction || "-")}</span>${span.bearingDegrees !== "" ? `<div class="muted">${escapeHtml(span.bearingDegrees)}°</div>` : ""}</td>
          <td>${escapeHtml(span.lengthDisplay || "")}</td>
          <td>${escapeHtml(other || "")}</td>
          <td>${escapeHtml(pole.lowPower || "")}</td>
          <td>${escapeHtml(side.maxCommHeight || pole.maxCommHeight || "")}</td>
          <td><input class="input height-input" data-scope="spanSide" data-pole="${escapeHtml(poleId)}" data-span="${escapeHtml(span.spanId)}" data-field="proposedHOA" value="${escapeHtml(side.proposedHOA || "")}" placeholder="24'1&quot;"></td>
          <td><input class="input height-input" data-scope="spanSide" data-pole="${escapeHtml(poleId)}" data-span="${escapeHtml(span.spanId)}" data-field="proposedMidspan" value="${escapeHtml(side.proposedMidspan || "")}" placeholder="23'7&quot;"></td>
          <td><input class="input height-input muted-input" data-scope="spanSide" data-pole="${escapeHtml(poleId)}" data-span="${escapeHtml(span.spanId)}" data-field="endDrop" value="${escapeHtml(side.endDrop || "")}" placeholder="auto"></td>
          <td><select class="input" data-scope="spanSide" data-pole="${escapeHtml(poleId)}" data-span="${escapeHtml(span.spanId)}" data-field="clearanceReference">
            <option value="LOW_POWER" ${side.clearanceReference === "LOW_POWER" ? "selected" : ""}>Low Power</option>
            <option value="TOP_COMM" ${side.clearanceReference === "TOP_COMM" ? "selected" : ""}>Top Comm</option>
            <option value="MANUAL" ${side.clearanceReference === "MANUAL" ? "selected" : ""}>Manual</option>
          </select></td>
          <td><textarea class="input text-input" data-scope="spanSide" data-pole="${escapeHtml(poleId)}" data-span="${escapeHtml(span.spanId)}" data-field="notes" placeholder="Notas propias...">${escapeHtml(side.notes || "")}</textarea></td>
        </tr>`;
      }).join("")}</tbody>
    </table></div>`;
  }

  function renderCommMovementTable(poleId) {
    const rows = S.getSpanCommsForPole(poleId).sort((a, b) => `${a.spanId}${a.owner}${a.wireIndex}`.localeCompare(`${b.spanId}${b.owner}${b.wireIndex}`, undefined, { numeric: true }));
    if (!rows.length) return `<p class="muted">No hay comms importados desde Span.Wire para este poste.</p>`;
    return `<div class="table-wrap"><table class="comm-movement-table">
      <thead><tr>
        <th>Owner/Comm</th><th>Existing HOA</th><th>Existing HOA Change</th><th>Diferencia</th><th>Span</th><th>Dir</th><th>Other Pole</th><th>HOA otro poste</th><th>Midspan importado</th><th>O-Calc MS</th><th>Midspan calculado</th><th>Clearance</th><th>Notas</th>
      </tr></thead>
      <tbody>${rows.map(sc => {
        const span = S.getSpan(sc.spanId);
        const pole = S.getPole(poleId);
        const effective = sc.existingHOAChange || sc.existingHOA;
        const aboveMax = effective && H.compareHeights(effective, pole?.maxCommHeight) === 1;
        const changed = Boolean(sc.existingHOAChange || sc.notes || sc.mr);
        return `<tr class="${changed ? "changed-row" : ""} ${aboveMax ? "warning-row" : ""}">
          <td><span class="badge owner">${escapeHtml(sc.owner)}</span><div class="muted">${escapeHtml(sc.size || sc.wireId || "")}</div></td>
          <td>${escapeHtml(sc.existingHOA || "")}</td>
          <td><input class="input height-input" data-scope="spanComm" data-pole="${escapeHtml(poleId)}" data-span="${escapeHtml(sc.spanId)}" data-owner="${escapeHtml(sc.owner)}" data-wire-id="${escapeHtml(sc.wireId || "")}" data-field="existingHOAChange" value="${escapeHtml(sc.existingHOAChange || "")}" placeholder="nueva altura"></td>
          <td>${escapeHtml(sc.difference || "")}</td>
          <td>${escapeHtml(spanLabel(span))}<div class="muted">${escapeHtml(sc.spanId)}</div></td>
          <td><span class="badge">${escapeHtml(span?.direction || "-")}</span></td>
          <td>${escapeHtml(sc.remotePoleId || (span ? S.getOtherPoleId(span, poleId) : ""))}</td>
          <td>${escapeHtml(sc.remoteHOA || "")}</td>
          <td>${escapeHtml(sc.midspan || "")}</td>
          <td><input class="input height-input" data-scope="spanComm" data-pole="${escapeHtml(poleId)}" data-span="${escapeHtml(sc.spanId)}" data-owner="${escapeHtml(sc.owner)}" data-wire-id="${escapeHtml(sc.wireId || "")}" data-field="ocalcMS" value="${escapeHtml(sc.ocalcMS || "")}" placeholder="21'8&quot;"></td>
          <td>${escapeHtml(sc.calculatedMidspan || "")}</td>
          <td>${aboveMax ? `<span class="badge danger">Sobre Max</span>` : `<span class="badge changed">OK</span>`}</td>
          <td><textarea class="input text-input" data-scope="spanComm" data-pole="${escapeHtml(poleId)}" data-span="${escapeHtml(sc.spanId)}" data-owner="${escapeHtml(sc.owner)}" data-wire-id="${escapeHtml(sc.wireId || "")}" data-field="notes" placeholder="Notas propias...">${escapeHtml(sc.notes || "")}</textarea></td>
        </tr>`;
      }).join("")}</tbody>
    </table></div>`;
  }

  function renderPowerTable(poleId) {
    const rows = S.getSpanPowerForPole(poleId);
    if (!rows.length) return `<p class="muted">No se importaron wires de power para este poste.</p>`;
    return `<div class="table-wrap"><table class="power-table">
      <thead><tr><th>Span</th><th>Tipo</th><th>Attachment Height</th><th>Midspan</th><th>Size</th></tr></thead>
      <tbody>${rows.map(row => {
        const span = S.getSpan(row.spanId);
        return `<tr><td>${escapeHtml(spanLabel(span))}</td><td><span class="badge warning">${escapeHtml(row.label)}</span></td><td>${escapeHtml(row.attachmentHeight)}</td><td>${escapeHtml(row.midspan)}</td><td>${escapeHtml(row.size)}</td></tr>`;
      }).join("")}</tbody>
    </table></div>`;
  }

  function renderPoleWorkspace(poleId) {
    return `<article class="pole-workspace-card" data-pole-card="${escapeHtml(poleId)}">
      ${renderPoleEditableHeader(poleId)}
      <div class="workspace-grid">
        <section class="subsection wide">
          <h4>Proposed por span</h4>
          <p class="muted">Este Proposed es por span/lado del poste, no por comm. End Drop se calcula con Proposed HOA y Proposed Midspan.</p>
          ${renderSpanProposedTable(poleId)}
        </section>
        <section class="subsection wide">
          <h4>Movimientos de comms existentes</h4>
          <p class="muted">Aquí se mueve cada comm existente con nueva altura. Si cambia el otro poste del mismo span, el Midspan calculado se actualiza.</p>
          ${renderCommMovementTable(poleId)}
        </section>
        <section class="subsection">
          <h4>Power / clearance importado</h4>
          ${renderPowerTable(poleId)}
        </section>
        <section class="subsection">
          <h4>MR único del poste</h4>
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
    els.polesOverview.innerHTML = poleIds.map(id => renderPoleWorkspace(id)).join("") || `<div class="detail-placeholder">No hay postes para mostrar.</div>`;
    wireEditableEvents(els.polesOverview);
  }

  function renderSelectedPoleDetail() {
    const state = S.getState();
    const poleId = state.selectedPoleId;
    if (!poleId || !S.getPole(poleId)) {
      els.selectedPoleDetail.innerHTML = `<div class="detail-placeholder">Selecciona un poste para enfocarlo.</div>`;
      return;
    }
    els.selectedPoleDetail.innerHTML = `<div class="detail-title"><div><p class="eyebrow">Detalle enfocado</p><h2>${escapeHtml(poleId)}</h2></div><button class="btn" type="button" data-scroll-to-pole="${escapeHtml(poleId)}">Ir al poste en la vista general</button></div>${renderPoleWorkspace(poleId)}`;
    wireEditableEvents(els.selectedPoleDetail);
    const btn = els.selectedPoleDetail.querySelector("[data-scroll-to-pole]");
    if (btn) btn.addEventListener("click", () => scrollToPole(btn.dataset.scrollToPole));
  }

  function wireEditableEvents(root) {
    root.querySelectorAll("input[data-scope], textarea[data-scope], select[data-scope]").forEach(input => {
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

  function handleEditableBlur(event) {
    const el = event.currentTarget;
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

  function handleEditableChange(event) {
    const el = event.currentTarget;
    const scope = el.dataset.scope;
    const field = el.dataset.field;
    const value = el.value;

    if (scope === "pole") {
      S.updatePoleField(el.dataset.pole, field, value);
      global.Calculations.recalculateSpansForPole(el.dataset.pole);
    }

    if (scope === "spanSide") {
      global.Calculations.updateSpanSideField(el.dataset.span, el.dataset.pole, field, value);
    }

    if (scope === "spanComm") {
      global.Calculations.updateSpanCommField(el.dataset.span, el.dataset.pole, el.dataset.owner, el.dataset.wireId || "", field, value);
    }

    render();
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

  function selectSpan(spanId) {
    const state = S.getState();
    state.selectedSpanId = spanId;
    const span = S.getSpan(spanId);
    if (span) state.selectedPoleId = span.fromPole;
    render();
  }

  function renderGraph() {
    global.GraphView.renderGraph(els.graphCanvas, selectPole, selectSpan);
  }

  function render() {
    global.Calculations.recalculateAll();
    renderSummary();
    renderPoleLists();
    renderAllPolesWorkspace();
    renderGraph();
    renderSelectedPoleDetail();
  }

  async function handleImport(file) {
    if (!file) return;
    try {
      await global.ExcelImport.importDataFile(file);
      render();
      const isJson = file.name.toLowerCase().endsWith(".json") || file.type === "application/json";
      toast(isJson ? "Datos JSON importados y cargados." : "Archivo importado. Se leyeron los datos disponibles.", "success");
    } catch (error) {
      console.error(error);
      toast(`Error importando datos: ${error.message}`, "error");
    } finally {
      els.excelFileInput.value = "";
    }
  }

  function bindEvents() {
    els.excelFileInput.addEventListener("change", event => handleImport(event.target.files[0]));
    els.exportExcelBtn.addEventListener("click", () => global.ExcelExport.exportData());
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
      exportExcelBtn: qs("exportExcelBtn"),
      saveLocalBtn: qs("saveLocalBtn"),
      loadLocalBtn: qs("loadLocalBtn"),
      resetSampleBtn: qs("resetSampleBtn"),
      poleSearchInput: qs("poleSearchInput"),
      warningFilterSelect: qs("warningFilterSelect"),
      summaryGrid: qs("summaryGrid"),
      polesList: qs("polesList"),
      polesOverview: qs("polesOverview"),
      graphCanvas: qs("graphCanvas"),
      selectedPoleDetail: qs("selectedPoleDetail"),
      toastHost: qs("toastHost")
    });

    bindEvents();
    global.FloatingCalculator?.setupFloatingCalculator();
    S.loadSampleData();
    render();
  }

  document.addEventListener("DOMContentLoaded", init);
})(window);
