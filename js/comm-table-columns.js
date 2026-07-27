(function (global) {
  "use strict";

  const COLUMN_LABEL = "Other Pole HOA";
  let observer = null;
  let refreshQueued = false;

  function normalizeLabel(value) {
    return String(value ?? "").replace(/\s+/g, " ").trim().toLowerCase();
  }

  function findColumnIndex(headers, label = COLUMN_LABEL) {
    const expected = normalizeLabel(label);
    return Array.from(headers || []).findIndex(header =>
      normalizeLabel(header?.textContent) === expected
    );
  }

  function removeColumnFromTable(table, columnIndex) {
    if (!table || columnIndex < 0) return false;

    let removed = false;
    Array.from(table.querySelectorAll("tr")).forEach(row => {
      const cell = row.children?.[columnIndex];
      if (!cell) return;
      cell.remove();
      removed = true;
    });
    return removed;
  }

  function removeOtherPoleHoaColumn(root = global.document) {
    if (!root?.querySelectorAll) return 0;

    let changedTables = 0;
    Array.from(root.querySelectorAll("table.comm-movement-table")).forEach(table => {
      const columnIndex = findColumnIndex(table.querySelectorAll("thead th"));
      if (columnIndex < 0) return;
      if (removeColumnFromTable(table, columnIndex)) changedTables++;
    });
    return changedTables;
  }

  function queueRefresh() {
    if (refreshQueued) return;
    refreshQueued = true;
    const run = () => {
      refreshQueued = false;
      removeOtherPoleHoaColumn();
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
    observer.observe(document.body || document.documentElement, {
      childList: true,
      subtree: true
    });
  }

  if (global.document?.readyState === "loading") {
    global.document.addEventListener("DOMContentLoaded", start, { once: true });
  } else {
    start();
  }

  global.CommTableColumns = {
    COLUMN_LABEL,
    findColumnIndex,
    removeColumnFromTable,
    removeOtherPoleHoaColumn
  };
})(window);
