(function (global) {
  "use strict";

  const COLOR_CLASS_COUNT = 5;
  const COLOR_CLASS_PATTERN = /^span-color-\d+$/;
  let observer = null;
  let refreshQueued = false;

  function removeSpanColorClasses(element) {
    if (!element?.classList) return;
    Array.from(element.classList).forEach(className => {
      if (COLOR_CLASS_PATTERN.test(className)) element.classList.remove(className);
    });
  }

  function applySpanColorClass(element, className) {
    if (!element?.classList) return;
    removeSpanColorClasses(element);
    element.classList.add(className);
  }

  function physicalSpanId(store, spanId) {
    const id = String(spanId || "").trim();
    if (!id || !store?.getSpan) return id;
    const span = store.getSpan(id);
    if (!span?.sourceSpanId) return span?.spanId || id;
    const source = store.getSpan(span.sourceSpanId);
    return source?.spanId || span.sourceSpanId || span.spanId || id;
  }

  function naturalSpanKey(store, spanId) {
    const span = store?.getSpan?.(spanId) || {};
    return String(span.spanIndex || span.spanId || spanId || "");
  }

  function sortVisibleSpanIds(store, poleId, spanIds) {
    const visible = Array.from(new Set((spanIds || []).filter(Boolean)));
    const connectedOrder = new Map();
    const connected = store?.getConnectedSpans?.(poleId) || [];
    [...connected]
      .sort((left, right) => naturalSpanKey(store, left.spanId).localeCompare(
        naturalSpanKey(store, right.spanId),
        undefined,
        { numeric: true }
      ))
      .forEach((span, index) => connectedOrder.set(physicalSpanId(store, span.spanId), index));

    return visible.sort((left, right) => {
      const leftIndex = connectedOrder.has(left) ? connectedOrder.get(left) : Number.MAX_SAFE_INTEGER;
      const rightIndex = connectedOrder.has(right) ? connectedOrder.get(right) : Number.MAX_SAFE_INTEGER;
      if (leftIndex !== rightIndex) return leftIndex - rightIndex;
      return naturalSpanKey(store, left).localeCompare(naturalSpanKey(store, right), undefined, { numeric: true });
    });
  }

  function classForVisibleIndex(index) {
    return `span-color-${Math.max(0, Number(index) || 0) % COLOR_CLASS_COUNT}`;
  }

  function firstSpanId(element) {
    const ownSpanId = String(element?.dataset?.span || "").trim();
    if (ownSpanId) return ownSpanId;
    return String(element?.querySelector?.("[data-span]")?.dataset?.span || "").trim();
  }

  function registerTarget(targetsBySpan, store, spanId, elements) {
    const resolvedSpanId = physicalSpanId(store, spanId);
    if (!resolvedSpanId) return;
    if (!targetsBySpan.has(resolvedSpanId)) targetsBySpan.set(resolvedSpanId, new Set());
    const targets = targetsBySpan.get(resolvedSpanId);
    (elements || []).filter(Boolean).forEach(element => targets.add(element));
  }

  function collectCommRowPairs(tableRow) {
    if (!tableRow?.querySelectorAll) return [];

    const spanRows = Array.from(
      tableRow.querySelectorAll(".comm-span-list .comm-span-row")
    );
    const midspanRows = Array.from(
      tableRow.querySelectorAll(".comm-midspan-list .comm-midspan-value.colored-midspan")
    );

    return spanRows.map((spanRow, index) => ({
      spanId: firstSpanId(spanRow),
      spanRow,
      midspanContainer: midspanRows[index] || null
    }));
  }

  function registerCommTableRowTargets(targetsBySpan, store, tableRow) {
    collectCommRowPairs(tableRow).forEach(pair => {
      if (!pair.spanId) return;

      registerTarget(targetsBySpan, store, pair.spanId, [
        ...pair.spanRow.querySelectorAll(".span-color-dot")
      ]);

      if (!pair.midspanContainer) return;
      registerTarget(targetsBySpan, store, pair.spanId, [
        pair.midspanContainer,
        ...pair.midspanContainer.querySelectorAll(
          ".midspan-highlight-input, .midspan-highlight-display"
        )
      ]);
    });
  }

  function collectTargetsForCard(card, store) {
    const targetsBySpan = new Map();

    card.querySelectorAll(".span-proposed-table tbody tr").forEach(row => {
      const spanId = firstSpanId(row);
      if (!spanId) return;
      registerTarget(targetsBySpan, store, spanId, [
        row,
        ...row.querySelectorAll(
          ".span-color-chip, .midspan-highlight-display, .midspan-highlight-input"
        )
      ]);
    });

    card.querySelectorAll(".comm-movement-table tbody > tr").forEach(tableRow => {
      registerCommTableRowTargets(targetsBySpan, store, tableRow);
    });

    // Fallback for future markup where a Midspan container carries data-span
    // directly but is not inside the currently grouped communication row.
    card.querySelectorAll(".comm-midspan-value.colored-midspan[data-span]").forEach(container => {
      const spanId = firstSpanId(container);
      if (!spanId) return;
      registerTarget(targetsBySpan, store, spanId, [
        container,
        ...container.querySelectorAll(
          ".midspan-highlight-input, .midspan-highlight-display"
        )
      ]);
    });

    return targetsBySpan;
  }

  function resetCardSpanColors(card, store = global.AppStore) {
    if (!card?.querySelectorAll || !store) return 0;
    const poleId = String(card.dataset?.poleCard || "").trim();
    if (!poleId) return 0;

    const targetsBySpan = collectTargetsForCard(card, store);
    const orderedSpanIds = sortVisibleSpanIds(
      store,
      poleId,
      Array.from(targetsBySpan.keys())
    );

    orderedSpanIds.forEach((spanId, index) => {
      const className = classForVisibleIndex(index);
      targetsBySpan.get(spanId)?.forEach(element => {
        applySpanColorClass(element, className);
      });
    });

    return orderedSpanIds.length;
  }

  function resetVisibleSpanColors(root = global.document, store = global.AppStore) {
    if (!root?.querySelectorAll || !store) return 0;
    let coloredSpans = 0;
    root.querySelectorAll("[data-pole-card]").forEach(card => {
      coloredSpans += resetCardSpanColors(card, store);
    });
    return coloredSpans;
  }

  function queueRefresh() {
    if (refreshQueued) return;
    refreshQueued = true;
    const run = () => {
      refreshQueued = false;
      resetVisibleSpanColors();
    };
    if (typeof global.queueMicrotask === "function") global.queueMicrotask(run);
    else global.setTimeout?.(run, 0);
  }

  function start() {
    const document = global.document;
    if (!document) return;

    queueRefresh();
    if (typeof global.MutationObserver !== "function" || observer) return;

    observer = new global.MutationObserver(queueRefresh);
    observer.observe(
      document.getElementById("polesOverview") ||
        document.body ||
        document.documentElement,
      {
        childList: true,
        subtree: true
      }
    );
  }

  if (global.document?.readyState === "loading") {
    global.document.addEventListener("DOMContentLoaded", start, { once: true });
  } else {
    start();
  }

  global.SpanColorReset = {
    COLOR_CLASS_COUNT,
    physicalSpanId,
    sortVisibleSpanIds,
    classForVisibleIndex,
    firstSpanId,
    collectCommRowPairs,
    resetCardSpanColors,
    resetVisibleSpanColors
  };
})(window);
