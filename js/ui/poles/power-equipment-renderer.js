(function (global) {
  "use strict";

  function create(context = {}) {
    const { S, H, escapeHtml, poleLink, getCalculations } = context;
    const calculations = () => typeof getCalculations === "function" ? getCalculations() : context.Calculations;

    function renderPowerTable(poleId) {
      const powerGroups = [];
      const heightGroups = new Map();
      S.getSpanPowerForPole(poleId).forEach(row => {
        const sourceHeight = H.parseHeight(row.attachmentHeight || "");
        if (sourceHeight === null) {
          powerGroups.push({ rows: [row] });
          return;
        }
        const key = String(sourceHeight);
        let group = heightGroups.get(key);
        if (!group) {
          group = { rows: [] };
          heightGroups.set(key, group);
          powerGroups.push(group);
        }
        group.rows.push(row);
      });
      powerGroups.sort((a, b) => {
        const aHeight = H.parseHeight(a.rows[0]?.attachmentHeight || "");
        const bHeight = H.parseHeight(b.rows[0]?.attachmentHeight || "");
        if (aHeight === null && bHeight === null) return 0;
        if (aHeight === null) return 1;
        if (bHeight === null) return -1;
        return bHeight - aHeight;
      });
      if (!powerGroups.length) return `<p class="muted">No power wires were imported for this pole.</p>`;
      return `<div class="table-wrap"><table class="power-table">
        <thead><tr><th>Span</th><th>Type</th><th>Attachment Height</th><th>HOA Change</th><th>Midspan</th><th>Actions</th></tr></thead>
        <tbody>${powerGroups.map(group => {
          const groupRows = [...group.rows].sort((a, b) =>
            String(a.spanId || "").localeCompare(String(b.spanId || ""), undefined, { numeric: true })
          );
          const representative = groupRows[0];
          const sharedHoaChange = groupRows.find(row => row.attachmentHeightChange)?.attachmentHeightChange || "";
          const spanEntries = groupRows.map(row => {
            const span = S.getSpan(row.spanId);
            return `<div class="power-group-entry">${span ? `${poleLink(span.fromPole)} → ${poleLink(span.toPole)}` : ""}</div>`;
          }).join("");
          const typeEntries = groupRows.map(row =>
            `<div class="power-group-entry"><span class="badge warning">${escapeHtml(row.label || "Power")}</span></div>`
          ).join("");
          const midspanEntries = groupRows.map(row => {
            const originalMidspan = String(row.midspan || "");
            const calculatedMidspan = String(row.calculatedMidspan || originalMidspan);
            const originalTitle = originalMidspan
              ? `Original Midspan before movements: ${originalMidspan}`
              : "No original Midspan was imported";
            return `<div class="power-group-entry"><input class="input height-input" data-scope="spanPower" data-power-key="${escapeHtml(row.key || "")}" data-pole="${escapeHtml(poleId)}" data-span="${escapeHtml(row.spanId)}" data-field="midspan" value="${escapeHtml(calculatedMidspan)}" title="${escapeHtml(originalTitle)}" aria-label="Power Midspan for ${escapeHtml(row.spanId)}"></div>`;
          }).join("");
          const actionEntries = groupRows.map(row => {
            const span = S.getSpan(row.spanId);
            const label = span ? `${span.fromPole} to ${span.toPole}` : "unlinked span";
            return `<button class="icon-action danger-action" type="button" data-delete-span-power data-power-key="${escapeHtml(row.key || "")}" data-pole="${escapeHtml(poleId)}" data-span="${escapeHtml(row.spanId)}" title="Delete ${escapeHtml(row.label || "power wire")} on ${escapeHtml(label)}" aria-label="Delete ${escapeHtml(row.label || "power wire")} on ${escapeHtml(label)}">&#10005;</button>`;
          }).join("");
          return `<tr>
            <td class="span-cell"><div class="power-group-stack">${spanEntries}</div></td>
            <td><div class="power-group-stack">${typeEntries}</div></td>
            <td><input class="input height-input" data-scope="spanPower" data-power-key="${escapeHtml(representative.key || "")}" data-pole="${escapeHtml(poleId)}" data-span="${escapeHtml(representative.spanId)}" data-field="attachmentHeight" value="${escapeHtml(representative.attachmentHeight || "")}" aria-label="Power attachment height shared by wires at this height"></td>
            <td><input class="input height-input" data-scope="spanPower" data-power-key="${escapeHtml(representative.key || "")}" data-pole="${escapeHtml(poleId)}" data-span="${escapeHtml(representative.spanId)}" data-field="attachmentHeightChange" value="${escapeHtml(sharedHoaChange)}" aria-label="Power HOA Change" title="Target attachment height shared by power wires at this pole and height"></td>
            <td><div class="power-group-stack">${midspanEntries}</div></td>
            <td><div class="power-group-stack">${actionEntries}</div></td>
          </tr>`;
        }).join("")}</tbody>
      </table></div>`;
    }

    function renderCollapseEyeIcon() {
      return `<span class="collapse-hint" aria-hidden="true">
        <svg class="collapse-show-icon" viewBox="0 0 24 24" focusable="false"><path d="M2.5 12s3.5-7 9.5-7 9.5 7 9.5 7-3.5 7-9.5 7-9.5-7-9.5-7Z"/><circle cx="12" cy="12" r="3"/></svg>
        <svg class="collapse-hide-icon" viewBox="0 0 24 24" focusable="false"><path d="M3 3l18 18M10.6 10.6a2 2 0 002.8 2.8M9.9 5.2A10.8 10.8 0 0112 5c5.5 0 9 7 9 7a15 15 0 01-3.1 3.9M6.2 6.2C3.9 7.8 3 12 3 12s3.5 7 9 7a10.8 10.8 0 004.1-.8"/></svg>
      </span>`;
    }

    function renderEquipmentTable(poleId) {
      const pole = S.getPole(poleId);
      const rows = Array.isArray(pole?.metadata?.powerEquipment) ? pole.metadata.powerEquipment : [];
      if (!rows.length) return `<p class="muted">No Power Equipment was imported for this pole.</p>`;
      const labels = {
        STREETLIGHT: "Streetlight",
        TRANSFORMER: "Transformer",
        RISER: "Power Riser"
      };
      const actionLabels = {
        STREETLIGHT: "Ground",
        TRANSFORMER: "Redress",
        RISER: "Raise"
      };
      return `<div class="table-wrap"><table class="power-table equipment-table">
        <thead><tr><th>Equipment</th><th>Owner</th><th>Attachment Height</th><th>Bottom Height</th><th>Drip Loop Height</th><th>Orientation</th><th>Action</th><th>New HOA</th><th>Max Comm Height</th><th>Actions</th></tr></thead>
        <tbody>${rows.map((row, index) => {
          const category = String(row.category || "").toUpperCase();
          const needsHeight = category === "TRANSFORMER" || category === "RISER";
          const settings = S.getState().settings || {};
          const requiredMidAmGround = category === "STREETLIGHT"
            && String(settings.projectProfile || "").toUpperCase() === "METRONET"
            && String(settings.proposedOwner || "MidAm").toUpperCase() === "MIDAM";
          const isIntecStreetlight = category === "STREETLIGHT"
            && String(settings.projectProfile || "INTEC").toUpperCase() === "INTEC";
          const isMidAmStreetlight = requiredMidAmGround;
          const isRaiseableStreetlight = isIntecStreetlight || isMidAmStreetlight;
          const isRiser = category === "RISER";
          const isTransformer = category === "TRANSFORMER";
          const actionControls = isRaiseableStreetlight
            ? `<div class="equipment-action-stack">
                <label class="equipment-action-control">
                  <input type="checkbox" data-scope="equipment" data-pole="${escapeHtml(poleId)}" data-equipment-index="${index}" data-field="actionActive" ${row.actionActive || requiredMidAmGround ? "checked" : ""} ${requiredMidAmGround ? "disabled title=\"Required by MidAm\"" : ""}>
                  <span>Ground</span>
                </label>
                <label class="equipment-action-control">
                  <input type="checkbox" data-scope="equipment" data-pole="${escapeHtml(poleId)}" data-equipment-index="${index}" data-field="raiseActive" ${row.raiseActive ? "checked" : ""}>
                  <span>Raise</span>
                </label>
              </div>`
            : isRiser
              ? `<div class="equipment-action-stack">
                  <label class="equipment-action-control">
                    <input type="checkbox" data-scope="equipment" data-pole="${escapeHtml(poleId)}" data-equipment-index="${index}" data-field="actionActive" ${row.actionActive ? "checked" : ""}>
                    <span>Raise</span>
                  </label>
                  <label class="equipment-action-control">
                    <input type="checkbox" data-scope="equipment" data-pole="${escapeHtml(poleId)}" data-equipment-index="${index}" data-field="secureActive" ${row.secureActive ? "checked" : ""}>
                    <span>Secure</span>
                  </label>
                </div>`
              : `<label class="equipment-action-control">
                <input type="checkbox" data-scope="equipment" data-pole="${escapeHtml(poleId)}" data-equipment-index="${index}" data-field="actionActive" ${row.actionActive || requiredMidAmGround ? "checked" : ""} ${requiredMidAmGround ? "disabled title=\"Required by MidAm\"" : ""}>
                <span>${escapeHtml(isTransformer ? "Secure" : (actionLabels[category] || "Apply"))}</span>
              </label>`;
          const heightField = isRaiseableStreetlight ? "raiseHeight" : "actionHeight";
          const heightEnabled = isRaiseableStreetlight ? row.raiseActive : row.actionActive;
          return `<tr>
          <td><span class="badge warning" title="${escapeHtml(row.type || "")}">${escapeHtml(labels[row.category] || row.category || "Equipment")}</span></td>
          <td class="inline-edit-cell" data-inline-equipment-edit="owner" data-pole="${escapeHtml(poleId)}" data-equipment-index="${index}" title="Double-click to edit equipment owner">${escapeHtml(row.owner || "")}</td>
          <td class="inline-edit-cell" data-inline-equipment-edit="attachmentHeight" data-pole="${escapeHtml(poleId)}" data-equipment-index="${index}" title="Double-click to edit attachment height">${escapeHtml(row.attachmentHeight || "")}</td>
          <td>${escapeHtml(row.bottomHeight || "")}</td>
          <td>${escapeHtml(row.dripLoopHeight || "")}</td>
          <td>${escapeHtml(row.orientation || "")}</td>
          <td>${actionControls}</td>
          <td>${needsHeight || isRaiseableStreetlight ? `<input class="input height-input equipment-action-height" data-scope="equipment" data-pole="${escapeHtml(poleId)}" data-equipment-index="${index}" data-field="${heightField}" value="${escapeHtml(row[heightField] || "")}" ${heightEnabled ? "" : "disabled"} ${isRaiseableStreetlight ? `title="Maximum raise: 1 foot above ${escapeHtml(row.attachmentHeight || "the imported HOA")}"` : ""}>` : `<span class="muted">&mdash;</span>`}</td>
          <td><strong>${escapeHtml(calculations().getPowerEquipmentCeiling(row) || "")}</strong></td>
          <td><button class="icon-action danger-action" type="button" data-delete-equipment data-pole="${escapeHtml(poleId)}" data-equipment-index="${index}" title="Delete equipment" aria-label="Delete equipment">&#10005;</button></td>
        </tr>`;
        }).join("")}</tbody>
      </table></div>`;
    }

    return { renderPowerTable, renderCollapseEyeIcon, renderEquipmentTable };
  }

  global.PowerEquipmentRenderer = { create };
})(window);
