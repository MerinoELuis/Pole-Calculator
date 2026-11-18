(function (global) {
  "use strict";

  function create(context = {}) {
    const { S, escapeHtml, renderCollapseEyeIcon } = context;

    function renderAnchorTable(poleId) {
      const storedAnchors = S.getPole(poleId)?.metadata?.anchors;
      const anchors = Array.isArray(storedAnchors) ? storedAnchors : [];
      if (!anchors.length) return `<p class="muted">No anchor data was imported for this pole.</p>`;
      return `<div class="anchor-record-list">${anchors.map((anchor, anchorIndex) => {
        const guys = Array.isArray(anchor.guys) ? anchor.guys : [];
        return `<section class="anchor-record" aria-label="Anchor ${anchorIndex + 1}">
          <div class="anchor-fields-grid">
            <label>Owner
              <input class="input" data-scope="anchor" data-pole="${escapeHtml(poleId)}" data-anchor-index="${anchorIndex}" data-field="owner" value="${escapeHtml(anchor.owner || "")}" aria-label="Anchor owner">
            </label>
            <label>Lead Length
              <input class="input height-input" data-scope="anchor" data-pole="${escapeHtml(poleId)}" data-anchor-index="${anchorIndex}" data-field="leadLength" value="${escapeHtml(anchor.leadLength || "")}" aria-label="Anchor lead length">
            </label>
            <label>Bearing / Angle
              <input class="input" data-scope="anchor" data-pole="${escapeHtml(poleId)}" data-anchor-index="${anchorIndex}" data-field="bearing" value="${escapeHtml(anchor.bearing || "")}" aria-label="Anchor bearing or angle">
            </label>
            <label class="anchor-violation-field">
              <span>Groundline Violation (&lt;6&quot; or &gt;18&quot;)</span>
              <span class="anchor-violation-control">
                <input type="checkbox" data-scope="anchor" data-pole="${escapeHtml(poleId)}" data-anchor-index="${anchorIndex}" data-field="groundlineViolation" ${anchor.groundlineViolation ? "checked" : ""} aria-label="Groundline violation">
                <span>${anchor.groundlineViolation ? "Yes" : "No"}</span>
              </span>
            </label>
          </div>
          ${guys.length
            ? `<div class="table-wrap anchor-guys-wrap"><table class="anchor-guys-table">
              <thead><tr><th colspan="3">Guys</th></tr><tr><th>Size</th><th>Owner</th><th>Attachment HOA</th></tr></thead>
              <tbody>${guys.map((guy, guyIndex) => `<tr>
                <td><input class="input" data-scope="anchor" data-pole="${escapeHtml(poleId)}" data-anchor-index="${anchorIndex}" data-guy-index="${guyIndex}" data-field="size" value="${escapeHtml(guy.size || "")}" aria-label="Guy size"></td>
                <td><input class="input" data-scope="anchor" data-pole="${escapeHtml(poleId)}" data-anchor-index="${anchorIndex}" data-guy-index="${guyIndex}" data-field="owner" value="${escapeHtml(guy.owner || "")}" aria-label="Guy owner"></td>
                <td><input class="input height-input" data-scope="anchor" data-pole="${escapeHtml(poleId)}" data-anchor-index="${anchorIndex}" data-guy-index="${guyIndex}" data-field="attachmentHeight" value="${escapeHtml(guy.attachmentHeight || "")}" aria-label="Guy attachment HOA"></td>
              </tr>`).join("")}</tbody>
            </table></div>`
            : `<p class="muted anchor-no-guys">No guys were associated with this anchor.</p>`}
        </section>`;
      }).join("")}</div>`;
    }

    function renderAnchorSection(poleId) {
      const anchors = S.getPole(poleId)?.metadata?.anchors;
      const openAttribute = Array.isArray(anchors) && anchors.length ? " open" : "";
      return `<details class="subsection wide anchor-section"${openAttribute}>
        <summary class="subsection-collapse-summary">
          <h4>ANC / Guys</h4>
          ${renderCollapseEyeIcon()}
        </summary>
        ${renderAnchorTable(poleId)}
      </details>`;
    }

    return { renderAnchorTable, renderAnchorSection };
  }

  global.AnchorRenderer = { create };
})(window);
