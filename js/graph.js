(function (global) {
  "use strict";

  // The map is a relative engineering layout, not a geographic basemap. The
  // first pole in each connected component starts at 0,0 and every next pole is
  // projected from the span length and bearing (0=N, 90=E, 180=S, 270=W).
  const S = () => global.AppStore;
  const H = () => global.HeightUtils;
  const mapViews = new WeakMap();

  function escapeHtml(value) {
    return String(value ?? "")
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&#039;");
  }

  function spanLengthFeet(span) {
    const displayInches = H().parseHeight(span?.lengthDisplay || "");
    if (displayInches !== null && displayInches > 0) return displayInches / 12;
    const rawMeters = Number(span?.length);
    return Number.isFinite(rawMeters) && rawMeters > 0 ? rawMeters * 3.28084 : null;
  }

  function normalizedBearing(span) {
    const value = Number(span?.bearingDegrees);
    if (!Number.isFinite(value)) return null;
    return ((value % 360) + 360) % 360;
  }

  function physicalEdgeKey(span) {
    return [span?.fromPole || "", span?.toPole || ""].sort().join("|");
  }

  function edgeScore(span) {
    const type = String(span?.type || "").toLowerCase();
    return (spanLengthFeet(span) !== null ? 100 : 0) +
      (normalizedBearing(span) !== null ? 100 : 0) +
      (/fore/.test(type) ? 20 : /other/.test(type) ? 10 : 0);
  }

  function physicalEdges() {
    const byPair = new Map();
    Object.values(S().getState().spans || {}).forEach(span => {
      if (!span.fromPole || !span.toPole || span.fromPole === span.toPole) return;
      const key = physicalEdgeKey(span);
      const current = byPair.get(key);
      if (!current || edgeScore(span) > edgeScore(current)) byPair.set(key, span);
    });
    return Array.from(byPair.values());
  }

  function fallbackBearing(spanId) {
    let hash = 0;
    String(spanId || "").split("").forEach(char => { hash = ((hash * 31) + char.charCodeAt(0)) >>> 0; });
    return hash % 360;
  }

  function project(origin, span, forward) {
    const length = spanLengthFeet(span) || 80;
    const baseBearing = normalizedBearing(span);
    const bearing = ((baseBearing === null ? fallbackBearing(span.spanId) : baseBearing) + (forward ? 0 : 180)) % 360;
    const radians = bearing * Math.PI / 180;
    return {
      x: origin.x + (Math.sin(radians) * length),
      y: origin.y - (Math.cos(radians) * length),
      inferred: baseBearing === null || spanLengthFeet(span) === null
    };
  }

  function connectedComponents(poleIds, edges) {
    const adjacency = new Map(poleIds.map(id => [id, []]));
    edges.forEach(edge => {
      if (!adjacency.has(edge.fromPole)) adjacency.set(edge.fromPole, []);
      if (!adjacency.has(edge.toPole)) adjacency.set(edge.toPole, []);
      adjacency.get(edge.fromPole).push({ edge, next: edge.toPole, forward: true });
      adjacency.get(edge.toPole).push({ edge, next: edge.fromPole, forward: false });
    });

    const seen = new Set();
    const components = [];
    Array.from(adjacency.keys()).sort(naturalSort).forEach(root => {
      if (seen.has(root)) return;
      const ids = [];
      const queue = [root];
      seen.add(root);
      while (queue.length) {
        const current = queue.shift();
        ids.push(current);
        (adjacency.get(current) || []).forEach(item => {
          if (seen.has(item.next)) return;
          seen.add(item.next);
          queue.push(item.next);
        });
      }
      components.push({ ids, adjacency });
    });
    return components;
  }

  function naturalSort(a, b) {
    return String(a).localeCompare(String(b), undefined, { numeric: true });
  }

  function componentLayout(component) {
    const nodes = {};
    const root = component.ids.slice().sort((a, b) => {
      const unknownDiff = Number(/^Unknown-/i.test(a)) - Number(/^Unknown-/i.test(b));
      return unknownDiff || naturalSort(a, b);
    })[0];
    nodes[root] = { x: 0, y: 0, inferred: false, root: true };
    const queue = [root];

    while (queue.length) {
      const current = queue.shift();
      (component.adjacency.get(current) || [])
        .slice()
        .sort((a, b) => naturalSort(a.edge.spanId, b.edge.spanId))
        .forEach(item => {
          if (nodes[item.next]) return;
          nodes[item.next] = project(nodes[current], item.edge, item.forward);
          queue.push(item.next);
        });
    }
    return nodes;
  }

  function boundsForNodes(nodes) {
    const values = Object.values(nodes);
    if (!values.length) return { minX: 0, minY: 0, maxX: 1, maxY: 1, width: 1, height: 1 };
    const xs = values.map(node => node.x);
    const ys = values.map(node => node.y);
    const minX = Math.min(...xs);
    const minY = Math.min(...ys);
    const maxX = Math.max(...xs);
    const maxY = Math.max(...ys);
    return { minX, minY, maxX, maxY, width: Math.max(maxX - minX, 1), height: Math.max(maxY - minY, 1) };
  }

  function buildLayout() {
    const state = S().getState();
    const poleIds = Object.keys(state.poles || {}).sort(naturalSort);
    const edges = physicalEdges();
    const components = connectedComponents(poleIds, edges);
    const nodes = {};
    const componentGap = 140;
    let cursorX = 0;

    components.forEach(component => {
      const local = componentLayout(component);
      const bounds = boundsForNodes(local);
      Object.entries(local).forEach(([poleId, node]) => {
        nodes[poleId] = {
          ...node,
          x: node.x - bounds.minX + cursorX,
          y: node.y - bounds.minY
        };
      });
      cursorX += bounds.width + componentGap;
    });

    return { nodes, edges, bounds: boundsForNodes(nodes) };
  }

  function paddedViewBox(bounds) {
    const largestDimension = Math.max(bounds.width, bounds.height, 100);
    const padding = Math.max(36, largestDimension * 0.09);
    return {
      x: bounds.minX - padding,
      y: bounds.minY - padding,
      width: bounds.width + (padding * 2),
      height: bounds.height + (padding * 2)
    };
  }

  function viewBoxText(box) {
    return `${box.x} ${box.y} ${box.width} ${box.height}`;
  }

  function shortPoleLabel(poleId) {
    return String(poleId || "").length > 20 ? `${String(poleId).slice(0, 18)}...` : String(poleId || "");
  }

  function renderGraph(container, onSelectPole, onSelectSpan) {
    if (!container) return;
    const state = S().getState();
    const layout = buildLayout();
    // Geometry belongs in the signature so Update Data can refit the map when
    // a span keeps its ID but receives a corrected length or bearing.
    const edgeSignature = layout.edges.map(edge => [
      edge.spanId,
      edge.fromPole,
      edge.toPole,
      spanLengthFeet(edge) ?? "",
      normalizedBearing(edge) ?? ""
    ].join(":")).sort().join("|");
    const signature = `${Object.keys(layout.nodes).sort().join("|")}::${edgeSignature}`;
    const fitBox = paddedViewBox(layout.bounds);
    const previous = mapViews.get(container);
    const currentBox = previous?.signature === signature ? previous.viewBox : { ...fitBox };
    mapViews.set(container, { signature, fitBox, viewBox: currentBox });

    if (!Object.keys(layout.nodes).length) {
      container.innerHTML = `<div class="detail-placeholder">Import pole and span data to build the map.</div>`;
      return;
    }

    const scaleReference = Math.max(fitBox.width, fitBox.height, 100);
    const nodeRadius = Math.max(4, Math.min(9, scaleReference / 110));
    const labelSize = Math.max(7, Math.min(12, scaleReference / 78));
    const edgeWidth = Math.max(1.2, scaleReference / 650);

    const edgeMarkup = layout.edges.map(span => {
      const from = layout.nodes[span.fromPole];
      const to = layout.nodes[span.toPole];
      if (!from || !to) return "";
      const active = state.selectedSpanId === span.spanId ? " active" : "";
      const inferred = normalizedBearing(span) === null || spanLengthFeet(span) === null ? " inferred" : "";
      return `<g class="map-edge-group${active}${inferred}" data-span-id="${escapeHtml(span.spanId)}" role="button" aria-label="${escapeHtml(`${span.fromPole} to ${span.toPole}`)}">
        <title>${escapeHtml(`${span.fromPole} -> ${span.toPole} | ${span.lengthDisplay || "No length"} | ${normalizedBearing(span) ?? "No bearing"} degrees`)}</title>
        <line class="map-edge-hit" x1="${from.x}" y1="${from.y}" x2="${to.x}" y2="${to.y}" />
        <line class="map-edge" style="stroke-width:${edgeWidth}" x1="${from.x}" y1="${from.y}" x2="${to.x}" y2="${to.y}" />
      </g>`;
    }).join("");

    const nodeMarkup = Object.entries(layout.nodes).map(([poleId, node]) => {
      const active = state.selectedPoleId === poleId ? " active" : "";
      const unknown = /^Unknown-/i.test(poleId) ? " unknown" : "";
      const inferred = node.inferred ? " inferred" : "";
      const origin = node.root ? " origin" : "";
      return `<g class="map-node${active}${unknown}${inferred}${origin}" data-pole-id="${escapeHtml(poleId)}" role="button" aria-label="Pole ${escapeHtml(poleId)}" transform="translate(${node.x} ${node.y})">
        <title>${escapeHtml(poleId)}${node.root ? " | component origin 0,0" : ""}</title>
        <circle r="${nodeRadius}"></circle>
        <text x="${nodeRadius + 3}" y="${-nodeRadius - 2}" style="font-size:${labelSize}px">${escapeHtml(shortPoleLabel(poleId))}</text>
      </g>`;
    }).join("");

    container.innerHTML = `<svg class="pole-map-svg" viewBox="${viewBoxText(currentBox)}" role="img" aria-label="Relative pole map based on span length and bearing" preserveAspectRatio="xMidYMid meet">
      <g class="pole-map-viewport">${edgeMarkup}${nodeMarkup}</g>
    </svg>`;

    const svg = container.querySelector("svg");
    container.querySelectorAll("[data-pole-id]").forEach(node => {
      node.addEventListener("click", event => {
        event.stopPropagation();
        onSelectPole(node.dataset.poleId);
      });
    });
    container.querySelectorAll("[data-span-id]").forEach(edge => {
      edge.addEventListener("click", event => {
        event.stopPropagation();
        onSelectSpan(edge.dataset.spanId);
      });
    });
    bindPan(container, svg);
  }

  function bindPan(container, svg) {
    if (!svg) return;
    let drag = null;
    svg.addEventListener("pointerdown", event => {
      if (event.target.closest("[data-pole-id], [data-span-id]")) return;
      const view = mapViews.get(container);
      if (!view) return;
      drag = { x: event.clientX, y: event.clientY, box: { ...view.viewBox } };
      svg.setPointerCapture(event.pointerId);
      svg.classList.add("dragging");
    });
    svg.addEventListener("pointermove", event => {
      if (!drag) return;
      const rect = svg.getBoundingClientRect();
      const view = mapViews.get(container);
      if (!view || !rect.width || !rect.height) return;
      view.viewBox = {
        ...drag.box,
        x: drag.box.x - ((event.clientX - drag.x) * drag.box.width / rect.width),
        y: drag.box.y - ((event.clientY - drag.y) * drag.box.height / rect.height)
      };
      svg.setAttribute("viewBox", viewBoxText(view.viewBox));
    });
    const stopDrag = event => {
      if (!drag) return;
      drag = null;
      svg.classList.remove("dragging");
      if (svg.hasPointerCapture(event.pointerId)) svg.releasePointerCapture(event.pointerId);
    };
    svg.addEventListener("pointerup", stopDrag);
    svg.addEventListener("pointercancel", stopDrag);
  }

  function changeZoom(container, factor) {
    const view = mapViews.get(container);
    const svg = container?.querySelector("svg");
    if (!view || !svg) return;
    const box = view.viewBox;
    const nextWidth = box.width * factor;
    const nextHeight = box.height * factor;
    view.viewBox = {
      x: box.x + ((box.width - nextWidth) / 2),
      y: box.y + ((box.height - nextHeight) / 2),
      width: nextWidth,
      height: nextHeight
    };
    svg.setAttribute("viewBox", viewBoxText(view.viewBox));
  }

  function fit(container) {
    const view = mapViews.get(container);
    const svg = container?.querySelector("svg");
    if (!view || !svg) return;
    view.viewBox = { ...view.fitBox };
    svg.setAttribute("viewBox", viewBoxText(view.viewBox));
  }

  global.GraphView = {
    renderGraph,
    zoomIn: container => changeZoom(container, 0.78),
    zoomOut: container => changeZoom(container, 1.28),
    fit
  };
})(window);
