(function (global) {
  "use strict";

  const WARNING_ICON = '<span class="pole-class-year-warning" role="img" aria-label="Additional loading required for poles installed in 2018 or later" title="Poles installed in 2018 or later require additional load cases."><svg class="pole-class-year-warning-icon" viewBox="0 0 24 24" aria-hidden="true" focusable="false"><path d="M12 3 2.8 19h18.4L12 3Z"/><path d="M12 9v4"/><path d="M12 16.5h.01"/></svg></span>';

  function create(context) {
    const { els, S, escapeHtml, wireEditableEvents } = context;
    const container = els.plaExistingResults;
    if (!container) return { render() {} };

    function caseLabel(loadCase) {
      if (loadCase.caseId === "NESC_MEDIUM_250B_GRADE_C") return `${escapeHtml(loadCase.name)} · 4 psf / 0.25" ice`;
      if (loadCase.caseId === "NESC_23_LIGHT_250B_GRADE_C") return escapeHtml(loadCase.name);
      if (loadCase.caseId === "NESC_23_EXT_WIND_250C_GRADE_C_LE_100_MPH") return `${escapeHtml(loadCase.name)} · 90 mph`;
      if (loadCase.caseId === "NESC_23_EXT_WIND_ICE_250D_GRADE_C") return `${escapeHtml(loadCase.name)} · 30 mph / 0.25" ice`;
      return loadCase.windSpeedMph
        ? `${escapeHtml(loadCase.name)} · ${escapeHtml(loadCase.windSpeedMph)} mph`
        : `${escapeHtml(loadCase.name)} · enter wind >100 mph`;
    }

    function renderLoadingList(plan) {
      if (plan.status === "YEAR_REQUIRED") return '<span class="pla-existing-empty">No loading selected until the year is known.</span>';
      const isModernPole = plan.yearInstalled >= 2018;
      const loadCases = plan.loadings.filter(loadCase => !isModernPole || loadCase.role === "ADDITIONAL");
      if (plan.status === "WIND_SPEED_REQUIRED") loadCases.unshift(global.AutoExistingExport.CASES.EXT_WIND);
      return `<ul class="pla-existing-load-list">${loadCases.map(loadCase => `<li>${caseLabel(loadCase)}</li>`).join("")}</ul>`;
    }

    function render() {
      const state = S.getState();
      const result = global.AutoExistingExport?.inspect(state) || { canExport: false, reason: "AutoExisting export is unavailable.", plans: [] };
      const supported = String(state.settings?.projectProfile || "").toUpperCase() === "INTEC";
      const exportButton = els.exportAutoExistingBtn;
      if (exportButton) {
        exportButton.disabled = !supported || !result.canExport;
        exportButton.title = !supported ? "PLA Existing is available only for INTEC." : (result.reason || "Export existing pole load cases for the O-Calc plugin.");
        exportButton.classList.toggle("btn-disabled", exportButton.disabled);
        exportButton.classList.toggle("btn-success", !exportButton.disabled);
      }
      if (els.plaExistingSummary) {
        const ready = result.plans.filter(plan => plan.status === "READY").length;
        const pendingYears = result.missingYearPoleIds.length;
        const pendingWind = result.missingWindPoleIds.length;
        els.plaExistingSummary.innerHTML = `<span class="badge changed">${ready} ready</span>${pendingYears ? `<span class="badge warning">${pendingYears} need year</span>` : ""}${pendingWind ? `<span class="badge warning">${pendingWind} need wind speed</span>` : ""}`;
      }

      if (!supported) {
        container.innerHTML = '<div class="detail-placeholder">PLA Existing is available only for INTEC jobs.</div>';
        return;
      }
      if (!result.plans.length) {
        container.innerHTML = '<div class="detail-placeholder">Import Collection data to assign existing pole load cases.</div>';
        return;
      }

      const loadingPrincipal = global.AutoExistingExport.loadingPrincipal(state);
      container.innerHTML = `<div class="pla-existing-layout">
        <div class="pla-existing-toolbar">
          <label class="pla-existing-global-loading"><span>Loading Principal</span><select class="input" data-scope="plaExistingGlobal" data-field="loadingPrincipal" aria-label="Loading Principal for this job">
            <option value="MEDIUM" ${loadingPrincipal === "MEDIUM" ? "selected" : ""}>NESC Medium (250B) Grade C</option>
            <option value="LIGHT" ${loadingPrincipal === "LIGHT" ? "selected" : ""}>NESC 23 Light (250B) Grade C</option>
          </select></label>
        </div>
        <div class="table-wrap pla-existing-table-wrap"><table class="pla-existing-table">
        <thead><tr><th>POLE</th><th>YEAR INSTALLED</th><th>WIND SPEED (&gt;100 MPH)</th><th>LOADINGS</th><th>STATUS</th></tr></thead>
        <tbody>${result.plans.map(plan => {
          const isModern = plan.yearInstalled !== null && plan.yearInstalled >= 2018;
          const yearHtml = plan.yearInstalled === null
            ? `<input class="input pla-existing-year-input" type="number" min="1800" max="2100" step="1" inputmode="numeric" data-scope="plaExisting" data-pole="${escapeHtml(plan.id)}" data-field="yearInstalled" value="${escapeHtml(plan.yearInputValue)}" placeholder="Enter year" aria-label="Year Installed for ${escapeHtml(plan.id)}">`
            : `<span class="pla-existing-year-value">${isModern ? WARNING_ICON : ""}<span>${escapeHtml(plan.yearInstalled)}</span></span>`;
          const windHtml = isModern
            ? `<label class="pla-existing-wind-field"><input class="input" type="number" min="101" step="1" inputmode="decimal" data-scope="plaExisting" data-pole="${escapeHtml(plan.id)}" data-field="windSpeedMph" value="${escapeHtml(plan.windSpeedMphInput)}" placeholder="&gt;100" aria-label="Wind speed in mph for ${escapeHtml(plan.id)}"><span>mph</span></label>`
            : '<span class="pla-existing-empty">—</span>';
          const status = plan.status === "YEAR_REQUIRED"
            ? '<span class="badge warning">Year required</span>'
            : plan.status === "WIND_SPEED_REQUIRED"
              ? '<span class="badge warning">Enter &gt;100 mph</span>'
              : '<span class="badge changed">Ready</span>';
          return `<tr data-pla-existing-pole="${escapeHtml(plan.id)}">
            <td><strong>${escapeHtml(plan.id)}</strong></td>
            <td>${yearHtml}</td>
            <td>${windHtml}</td>
            <td>${renderLoadingList(plan)}</td>
            <td>${status}</td>
          </tr>`;
        }).join("")}</tbody>
      </table></div></div>`;
      wireEditableEvents(container);
    }

    return { render };
  }

  global.PlaExistingRenderer = Object.freeze({ create });
})(window);
