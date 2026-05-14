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

  function shortSpanLabel(span) {
    if (!span) return "";
    return `${span.fromPole || "?"} → ${span.toPole || "Unknown"}`;
  }

  function poleLink(poleId) {
    if (!poleId) return "";
    return `<button class="link-button" type="button" data-scroll-to-pole="${escapeHtml(poleId)}">${escapeHtml(poleId)}</button>`;
  }

  function displayMidspan(sc) {
    return sc.calculatedMidspan || sc.midspan || sc.ocalcMS || "";
  }

  function renderClearanceStatus(height, pole, missingLabel = "Missing Data") {
    if (!pole || !pole.lowPower) return `<span class="badge warning">${missingLabel}</span>`;
    const max = pole.maxCommHeight || "";
    if (!height) return `<span class="badge warning">Missing Data</span>${max ? `<div class="cell-hint">Max ${escapeHtml(max)}</div>` : ""}`;
    const aboveMax = max && H.compareHeights(height, max) === 1;
    if (aboveMax) return `<span class="badge danger">Clearance Issue</span><div class="cell-hint">Max ${escapeHtml(max)} · Low Power ${escapeHtml(pole.lowPower)}</div>`;
    return `<span class="badge changed">OK</span><div class="cell-hint">Max ${escapeHtml(max)} · Low Power ${escapeHtml(pole.lowPower)}</div>`;
  }

  function renderClearanceSettings() {
    if (!els.clearanceSettings) return;
    const settings = S.getState().settings || {};
    const rows = [
      ["polePowerCommsClearance", "Pole · Power-comms", settings.polePowerCommsClearance || settings.clearanceToPower || "40\""],
      ["commClearance", "Pole · Comm-comm", settings.commClearance || "12\""],
      ["boltClearance", "Pole · Bolt-bolt", settings.boltClearance || "4\""],
      ["midspanPowerCommClearance", "Midspan · Power-comm", settings.midspanPowerCommClearance || "30\""],
      ["midspanCommCommClearance", "Midspan · Comm-comm", settings.midspanCommCommClearance || "4'"]
    ];
    els.clearanceSettings.innerHTML = rows.map(([field, label, value]) => `
      <label class="clearance-row">
        <span>${escapeHtml(label)}</span>
        <input class="input height-input" data-scope="settings" data-field="${escapeHtml(field)}" value="${escapeHtml(value)}" />
      </label>
    `).join("");
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
    document.querySelectorAll("[data-pole-select]").forEach(btn => {
      btn.addEventListener("click", () => selectPole(btn.dataset.poleSelect));
    });
    document.querySelectorAll("[data-scroll-to-pole]").forEach(btn => {
      btn.addEventListener("click", () => selectPole(btn.dataset.scrollToPole));
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
        <th>Span</th><th>Environment</th><th>Env Clearance</th><th>Low Power</th><th>Max Comm</th><th>Proposed</th><th>End Drop</th><th>Cambio Proposed</th><th>O-Calc Midspan</th><th>Clearance</th><th>Notas</th>
      </tr></thead>
      <tbody>${spans.map(span => {
        const side = S.getSpanSide(span.spanId, poleId) || S.upsertSpanSide({ spanId: span.spanId, poleId });
        const pole = S.getPole(poleId);
        const aboveMax = side.proposedHOA && H.compareHeights(side.proposedHOA, side.maxCommHeight || pole?.maxCommHeight) === 1;
        return `<tr class="${side.proposedHOA || side.proposedMidspan || side.endDrop ? "changed-row" : ""} ${aboveMax ? "warning-row" : ""}">
          <td class="span-cell"><strong>${poleLink(span.fromPole)} → ${poleLink(span.toPole)}</strong></td>
          <td><select class="input environment-input" data-scope="span" data-span="${escapeHtml(span.spanId)}" data-field="environment">${renderEnvironmentOptions(span.environment)}</select></td>
          <td><input class="input" data-scope="span" data-span="${escapeHtml(span.spanId)}" data-field="environmentClearance" value="${escapeHtml(span.environmentClearance || "")}"></td>
          <td>${escapeHtml(pole?.lowPower || "")}</td>
          <td>${escapeHtml(side.maxCommHeight || pole?.maxCommHeight || "")}</td>
          <td><input class="input height-input" data-scope="spanSide" data-pole="${escapeHtml(poleId)}" data-span="${escapeHtml(span.spanId)}" data-field="proposedHOA" value="${escapeHtml(side.proposedHOA || "")}"></td>
          <td><span class="calculated-value">${escapeHtml(side.endDrop || "")}</span></td>
          <td><input class="input height-input" data-scope="spanSide" data-pole="${escapeHtml(poleId)}" data-span="${escapeHtml(span.spanId)}" data-field="proposedHOAChange" value="${escapeHtml(side.proposedHOAChange || "")}"></td>
          <td><input class="input height-input" data-scope="spanSide" data-pole="${escapeHtml(poleId)}" data-span="${escapeHtml(span.spanId)}" data-field="proposedMidspan" value="${escapeHtml(side.proposedMidspan || "")}"></td>
          <td>${renderClearanceStatus(side.proposedHOA, pole)}</td>
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
        <th>Owner/Comm</th><th>Existing HOA</th><th>Cambio de HOA</th><th>Span</th><th>HOA otro poste</th><th>Midspan</th><th>Clearance</th>
      </tr></thead>
      <tbody>${rows.map(sc => {
        const span = S.getSpan(sc.spanId);
        const pole = S.getPole(poleId);
        const effective = sc.existingHOAChange || sc.existingHOA;
        const aboveMax = effective && H.compareHeights(effective, pole?.maxCommHeight) === 1;
        const changed = Boolean(sc.existingHOAChange || sc.notes || sc.mr);
        return `<tr class="${changed ? "changed-row" : ""} ${aboveMax ? "warning-row" : ""}">
          <td><span class="badge owner">${escapeHtml(sc.owner)}</span></td>
          <td>${escapeHtml(sc.existingHOA || "")}</td>
          <td><input class="input height-input" data-scope="spanComm" data-pole="${escapeHtml(poleId)}" data-span="${escapeHtml(sc.spanId)}" data-owner="${escapeHtml(sc.owner)}" data-wire-id="${escapeHtml(sc.wireId || "")}" data-field="existingHOAChange" value="${escapeHtml(sc.existingHOAChange || "")}"></td>
          <td class="span-cell">${span ? `${poleLink(span.fromPole)} → ${poleLink(span.toPole)}` : ""}</td>
          <td>${escapeHtml(sc.remoteHOA || "")}</td>
          <td>${escapeHtml(displayMidspan(sc))}</td>
          <td>${renderClearanceStatus(effective, pole)}</td>
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
        return `<tr>
          <td class="span-cell">${span ? `${poleLink(span.fromPole)} → ${poleLink(span.toPole)}` : ""}</td>
          <td><span class="badge warning">${escapeHtml(row.label)}</span></td>
          <td>${escapeHtml(row.attachmentHeight)}</td>
          <td><input class="input height-input" data-scope="spanPower" data-power-key="${escapeHtml(row.key || "")}" data-span="${escapeHtml(row.spanId)}" data-field="midspan" value="${escapeHtml(row.midspan || "")}"></td>
          <td>${escapeHtml(row.size)}</td>
        </tr>`;
      }).join("")}</tbody>
    </table></div>`;
  }

  function renderPoleWorkspace(poleId) {
    return `<article class="pole-workspace-card" data-pole-card="${escapeHtml(poleId)}">
      ${renderPoleEditableHeader(poleId)}
      <div class="workspace-grid">
        <section class="subsection wide">
          <h4>Proposed por span</h4>
          <p class="muted">Este Proposed es por span/lado del poste. End Drop se calcula contra Cambio Proposed o contra el Proposed del otro poste.</p>
          ${renderSpanProposedTable(poleId)}
        </section>
        <section class="subsection wide">
          <h4>Movimientos de comms existentes</h4>
          <p class="muted">Aquí se mueve cada comm existente con nueva altura. Si cambia el otro poste del mismo span, el Midspan calculado se actualiza.</p>
          ${renderCommMovementTable(poleId)}
        </section>
        <section class="subsection wide">
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
    els.polesOverview.querySelectorAll("[data-scroll-to-pole]").forEach(btn => {
      btn.addEventListener("click", () => selectPole(btn.dataset.scrollToPole));
    });
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

    if (scope === "spanPower") {
      S.updateSpanPowerField(el.dataset.powerKey || "", field, value);
      global.Calculations.recalculateSpan(el.dataset.span);
    }

    if (scope === "settings") {
      S.updateSetting(field, value);
      global.Calculations.recalculateAll();
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

  function render() {
    global.Calculations.recalculateAll();
    renderSummary();
    renderClearanceSettings();
    renderPoleLists();
    renderAllPolesWorkspace();
    renderSelectedPoleDetail();
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
    els.exportExcelBtn.addEventListener("click", () => global.ExcelExport.exportExcel());
    els.exportJsonBtn.addEventListener("click", () => global.ExcelExport.exportJson());
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
      exportExcelBtn: qs("exportExcelBtn"),
      exportJsonBtn: qs("exportJsonBtn"),
      saveLocalBtn: qs("saveLocalBtn"),
      loadLocalBtn: qs("loadLocalBtn"),
      resetSampleBtn: qs("resetSampleBtn"),
      projectMeta: qs("projectMeta"),
      poleSearchInput: qs("poleSearchInput"),
      warningFilterSelect: qs("warningFilterSelect"),
      polesList: qs("polesList"),
      polesOverview: qs("polesOverview"),
      selectedPoleDetail: qs("selectedPoleDetail"),
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
