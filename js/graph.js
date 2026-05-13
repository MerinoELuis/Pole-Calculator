(function (global) {
  "use strict";

  const S = () => global.AppStore;

  function buildLayout(width, height) {
    const state = S().getState();
    const nodes = {};
    const poles = Object.keys(state.poles).sort((a, b) => a.localeCompare(b, undefined, { numeric: true }));
    const known = poles.filter(id => !/^Unknown-/i.test(id));
    const generated = poles.filter(id => /^Unknown-/i.test(id));
    const columns = Math.max(1, Math.ceil(Math.sqrt(known.length || 1)));
    const xGap = width / (columns + 1);
    const yGap = height / (Math.ceil((known.length || 1) / columns) + 1);

    known.forEach((poleId, index) => {
      const col = index % columns;
      const row = Math.floor(index / columns);
      nodes[poleId] = { x: xGap * (col + 1), y: yGap * (row + 1) };
    });

    generated.forEach((poleId, index) => {
      nodes[poleId] = { x: width - 120, y: 70 + index * 64 };
    });

    Object.values(state.spans).forEach(span => {
      if (span.fromPole && !nodes[span.fromPole]) nodes[span.fromPole] = { x: 80, y: height - 70 };
      if (span.toPole && !nodes[span.toPole]) nodes[span.toPole] = { x: width - 100, y: height - 70 };
    });

    return nodes;
  }

  function escapeHtml(value) {
    return String(value ?? "")
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&#039;");
  }

  function renderGraph(container, onSelectPole, onSelectSpan) {
    const state = S().getState();
    if (!container) return;
    const width = Math.max(container.clientWidth || 1100, 760);
    const height = 420;
    const nodes = buildLayout(width, height);
    const selectedPole = state.selectedPoleId;
    const selectedSpan = state.selectedSpanId;

    const spanLines = Object.values(state.spans).map(span => {
      const a = nodes[span.fromPole];
      const b = nodes[span.toPole];
      if (!a || !b) return "";
      const active = span.spanId === selectedSpan ? " active" : "";
      const mx = (a.x + b.x) / 2;
      const my = (a.y + b.y) / 2 - 8;
      const label = [span.direction || "", span.lengthDisplay || ""].filter(Boolean).join(" · ");
      return `<g data-span-id="${escapeHtml(span.spanId)}">
        <line class="graph-edge${active}" x1="${a.x}" y1="${a.y}" x2="${b.x}" y2="${b.y}" />
        <text class="graph-label" x="${mx}" y="${my}" text-anchor="middle">${escapeHtml(label || span.spanId)}</text>
      </g>`;
    }).join("");

    const nodeItems = Object.entries(nodes).map(([poleId, pos]) => {
      const active = poleId === selectedPole ? " active" : "";
      const unknown = /^Unknown-/i.test(poleId) ? " unknown" : "";
      const label = poleId.length > 14 ? poleId.slice(0, 12) + "…" : poleId;
      return `<g class="graph-node${active}${unknown}" data-pole-id="${escapeHtml(poleId)}" transform="translate(${pos.x}, ${pos.y})">
        <circle r="31"></circle>
        <text y="5" text-anchor="middle">${escapeHtml(label)}</text>
      </g>`;
    }).join("");

    container.innerHTML = `<svg class="graph-svg" viewBox="0 0 ${width} ${height}" role="img" aria-label="Gráfica de postes y spans">${spanLines}${nodeItems}</svg>`;
    container.querySelectorAll("[data-pole-id]").forEach(node => {
      node.addEventListener("click", () => onSelectPole(node.dataset.poleId));
    });
    container.querySelectorAll("[data-span-id]").forEach(edge => {
      edge.addEventListener("click", () => onSelectSpan(edge.dataset.spanId));
    });
  }

  global.GraphView = { renderGraph };
})(window);
