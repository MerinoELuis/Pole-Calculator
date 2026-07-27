(function (global) {
  "use strict";

  const HIDDEN_COLUMN_KEY = "other-pole-hoa";
  const HIDDEN_COLUMN_LABEL = "Other Pole HOA";

  function normalizeLabel(value) {
    const helper = global.UiDomContract?.normalizeLabel;
    return helper
      ? helper(value)
      : String(value ?? "").replace(/\s+/g, " ").trim().toLowerCase();
  }

  function findColumnIndex(headers) {
    const list = Array.from(headers || []);
    const byContract = list.findIndex(header => header?.dataset?.column === HIDDEN_COLUMN_KEY);
    if (byContract >= 0) return byContract;
    const expected = normalizeLabel(HIDDEN_COLUMN_LABEL);
    return list.findIndex(header => normalizeLabel(header?.textContent) === expected);
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

  function refresh(root = global.document) {
    if (!root?.querySelectorAll) return 0;
    let changed = 0;
    root.querySelectorAll("table.comm-movement-table, table[data-component='comm-movement-table']")
      .forEach(table => {
        const columnIndex = findColumnIndex(table.querySelectorAll("thead th"));
        if (columnIndex >= 0 && removeColumnFromTable(table, columnIndex)) changed += 1;
      });
    return changed;
  }

  global.CommTableUI = {
    HIDDEN_COLUMN_KEY,
    HIDDEN_COLUMN_LABEL,
    findColumnIndex,
    removeColumnFromTable,
    refresh
  };
})(window);
