(function (global) {
  "use strict";

  global.PoleSettingsRenderer = {
    create(context) {
      const { els, escapeHtml, S, wireEditableEvents } = context;

      function renderClearanceSettings() {
        if (!els.clearanceSettings) return;
        const settings = S.getState().settings || {};
        const profiles = global.ProjectProfiles?.PROFILES || {};
        const selectedProfile = settings.projectProfile || "INTEC";
        const projectOptions = Object.values(profiles)
          .filter(profile => profile.visible !== false)
          .filter((profile, index, all) => all.findIndex(item => item.id === profile.id) === index)
          .map(profile =>
            `<option value="${escapeHtml(profile.id)}" ${selectedProfile === profile.id ? "selected" : ""}>${escapeHtml(profile.label)}</option>`
          ).join("");
        const isComedWI = String(settings.metronetWI || "").toUpperCase() === "COMED";
        const poleClearanceRows = [
          ["polePowerCommsClearance", "Pole · Power-comms", settings.polePowerCommsClearance || settings.clearanceToPower || "40\""],
          ["commClearance", "Pole · Comm-comm", settings.commClearance || "12\""],
          ["boltClearance", "Pole · Bolt-bolt", settings.boltClearance || "4\""]
        ];
        const midspanClearanceRows = [
          ["midspanPowerCommClearance", "Midspan · Power-comm", settings.midspanPowerCommClearance || "30\""],
          ["midspanCommCommClearance", "Midspan · Comm-comm", settings.midspanCommCommClearance || "4\""]
        ];
        const streetlightClearanceRows = [];
        if (!isComedWI) {
          poleClearanceRows.splice(1, 0, ["primaryPowerCommsClearance", "Pole · Primary-comm", settings.primaryPowerCommsClearance || "43\""]);
          midspanClearanceRows.splice(1, 0, ["midspanPrimaryPowerCommClearance", "Midspan · Primary-comm", settings.midspanPrimaryPowerCommClearance || "33\""]);
        }
        if (selectedProfile === "INTEC") {
          streetlightClearanceRows.push(
            ["streetlightBracketCommClearance", "Streetlight bracket-comm", settings.streetlightBracketCommClearance || "40\""],
            ["streetlightGroundedBracketCommClearance", "Streetlight bracket-comm (Grounding)", settings.streetlightGroundedBracketCommClearance || "12\""]
          );
        }
        if ((selectedProfile === "METRONET" && String(settings.proposedOwner || "MidAm").toUpperCase() === "MIDAM")
          || (selectedProfile === "METRONET" && isComedWI)) {
          streetlightClearanceRows.push(
            ["streetlightBracketCommClearance", "Streetlight bracket-comm", settings.streetlightBracketCommClearance || (isComedWI ? '4"' : '20"')],
            ["streetlightDripLoopCommClearance", "Streetlight drip loop-comm", settings.streetlightDripLoopCommClearance || "12\""]
          );
        }
        const rawPosition = String(settings.position || "TOP_COMM").toUpperCase();
        const position = rawPosition === "LOW_COMM" ? "LOW_COMM"
          : rawPosition === "OVERLASH" ? "OVERLASH" : "TOP_COMM";
        const proposedOwner = settings.proposedOwner || "Wecom";
        const metronetWI = String(settings.metronetWI || (String(proposedOwner).toUpperCase() === "MNT" ? "CSU" : String(proposedOwner).toUpperCase() === "POWER" ? "COMED" : "MIDAM")).toUpperCase();
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
        const profileSettings = global.ProjectProfiles?.getProfileSettings?.(selectedProfile) || {};
        const proposedOwners = selectedProfile === "OLSSON_OPPD"
          ? ["Cox", "Verizon", "Uniti"]
          : (Array.isArray(profileSettings.proposedOwners) && profileSettings.proposedOwners.length
            ? profileSettings.proposedOwners
            : ["Wecom", "MidAm", "CenturyLink", "Cable One", "Cox", "Fatbeam", "Vexus", "MCI Metro"]);
        const ownerOptions = proposedOwners.map(owner =>
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
            <div class="clearance-settings-groups">
              <div class="settings-grid clearance-grid">${poleClearanceRows.map(renderRow).join("")}</div>
              <div class="settings-grid clearance-grid">${midspanClearanceRows.map(renderRow).join("")}</div>
              ${streetlightClearanceRows.length ? `<div class="settings-grid clearance-grid">${streetlightClearanceRows.map(renderRow).join("")}</div>` : ""}
            </div>
          </div>
          <div class="settings-section settings-section-adjustments">
            <h3>Adjustments</h3>
            <div class="settings-grid adjustments-grid">
              <label class="clearance-row position-row">
                <span>Project</span>
                <select class="input position-select" data-scope="settings" data-field="projectProfile">${projectOptions}</select>
              </label>
              <label class="clearance-row position-row">
                <span class="position-label">Position</span>
                <select class="input position-select" data-scope="settings" data-field="position">
                  <option value="TOP_COMM" ${position === "TOP_COMM" ? "selected" : ""}>Top Comm</option>
                  <option value="LOW_COMM" ${position === "LOW_COMM" ? "selected" : ""}>Low Comm</option>
                  <option value="OVERLASH" ${position === "OVERLASH" ? "selected" : ""}>Overlash</option>
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
                <select class="input position-select" data-scope="settings" data-field="metronetWI">
                  <option value="MIDAM" ${metronetWI === "MIDAM" ? "selected" : ""}>MidAm</option>
                  <option value="CSU" ${metronetWI === "CSU" ? "selected" : ""}>CSU</option>
                  <option value="COMED" ${metronetWI === "COMED" ? "selected" : ""}>ComEd</option>
                </select>
              </label>` : settings.hideProposedOwner ? "" : `<label class="clearance-row position-row">
                <span>Proposed Owner</span>
                <select class="input position-select" data-scope="settings" data-field="proposedOwner">${ownerOptions}</select>
              </label>`}
            </div>
          </div>
          ${selectedProfile === "INTEC" && fiberTypes.length ? `<div class="settings-section settings-section-fiber">
            <h3>Fiber</h3>
            <div class="settings-grid fiber-grid">
              <label class="clearance-row">
                <span>Messenger Size</span>
                <input class="input" data-scope="attachmentSettings" data-field="attachmentMessengerSize" value="${escapeHtml(settings.attachmentMessengerSize || "")}" />
              </label>
              ${fiberTypes.map(fiber => `<label class="clearance-row">
                <span>${escapeHtml(fiber)} Size</span>
                <input class="input" data-scope="attachmentSettings" data-field="fiberSize" data-fiber="${escapeHtml(fiber)}" value="${escapeHtml(fiberSizes[fiber] || "")}" />
              </label>`).join("")}
            </div>
          </div>` : ""}
        `;
        wireEditableEvents(els.clearanceSettings);
      }

      function updateAutoCalculateButtonState() {
        if (!els.autoCalculateBtn) {
          updateFollowProposedButtonState();
          return;
        }
        const state = S.getState();
        const hasPoleData = Object.keys(state.poles || {}).length > 0;
        const hasProposedFiber = global.CompactAutoProposed?.hasProposedFiber?.(state) === true;
        const supportedExportProject = global.CompactAutoProposed?.isExportProjectSupported?.(state) === true;
        const allowsOverlashWithoutFiber = global.CompactAutoProposed?.allowsOverlashExportWithoutProposedFiber?.(state) === true;
        const position = String(state.settings?.position || "TOP_COMM").toUpperCase();
        const isLowComm = position === "LOW_COMM";
        const isOverlash = position === "OVERLASH";
        const disableAuto = !hasPoleData || isLowComm || isOverlash || context.autoCalculateRunning;
        els.autoCalculateBtn.disabled = disableAuto;
        els.autoCalculateBtn.classList.toggle("btn-disabled", disableAuto);
        els.autoCalculateBtn.classList.toggle("btn-primary", !disableAuto);
        els.autoCalculateBtn.textContent = context.autoCalculateRunning
          ? "Auto Proposed · Processing"
          : !hasPoleData
            ? "Auto Proposed · Import Data First"
            : isLowComm
              ? "Auto Proposed · Top Comm Required"
              : isOverlash
                ? "Auto Proposed · Manual Overlash"
                : "Auto Proposed";

        if (els.exportProposedJsonBtn) {
          const canExportProposed = hasPoleData && (hasProposedFiber || allowsOverlashWithoutFiber) && supportedExportProject;
          els.exportProposedJsonBtn.disabled = !canExportProposed;
          els.exportProposedJsonBtn.classList.toggle("btn-disabled", !canExportProposed);
          els.exportProposedJsonBtn.classList.toggle("btn-success", canExportProposed);
          els.exportProposedJsonBtn.title = !hasPoleData
            ? "Import pole data first."
            : !supportedExportProject
              ? "AutoProposed export is currently available for INTEC and Olsson OPPD only."
              : !hasProposedFiber
                ? allowsOverlashWithoutFiber
                  ? "Export Olsson OPPD overlash data without a newly proposed fiber."
                  : "Add at least one proposed fiber before exporting."
                : "Export the proposed fiber data.";
        }
        if (els.exportDebugJsonBtn) {
          const buildSource = String(global.AppDeploymentVersion?.source || "development").toLowerCase();
          const localOnly = buildSource !== "github-pages";
          els.exportDebugJsonBtn.disabled = !hasPoleData || !localOnly;
          els.exportDebugJsonBtn.classList.toggle("btn-disabled", !hasPoleData || !localOnly);
          els.exportDebugJsonBtn.classList.toggle("hidden", !localOnly);
        }
        updateFollowProposedButtonState();
      }

      function updateFollowProposedButtonState() {
        const button = els.followProposedBtn;
        if (!button) return;
        const state = S.getState();
        const hasPoleData = Object.keys(state.poles || {}).length > 0;
        const position = String(state.settings?.position || "TOP_COMM").toUpperCase();
        const active = state.settings?.autoMoveCommsToProposed === true
          || String(state.settings?.autoMoveCommsToProposed).toLowerCase() === "true";
        const disabled = !hasPoleData || position !== "TOP_COMM" || context.autoCalculateRunning;
        button.disabled = disabled;
        button.setAttribute("aria-pressed", String(active));
        button.classList.toggle("btn-disabled", disabled);
        button.classList.toggle("btn-primary", active && !disabled);
        button.textContent = active ? "Follow Proposed: ON" : "Follow Proposed: OFF";
        button.title = position !== "TOP_COMM"
          ? "Follow Proposed is available only in Top Comm mode."
          : "When enabled, changing Proposed recalculates existing comm movements.";
      }

      return { renderClearanceSettings, updateAutoCalculateButtonState, updateFollowProposedButtonState };
    }
  };
})(window);
