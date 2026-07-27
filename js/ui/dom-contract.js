(function (global) {
  "use strict";

  function text(value) {
    return String(value ?? "").trim();
  }

  function normalizeLabel(value) {
    return text(value)
      .toLowerCase()
      .normalize("NFD")
      .replace(/[\u0300-\u036f]/g, "")
      .replace(/[^a-z0-9]+/g, " ")
      .trim();
  }

  function columnKey(value) {
    return normalizeLabel(value).replace(/\s+/g, "-") || "column";
  }

  function spanIdFrom(element) {
    const own = text(element?.dataset?.spanId || element?.dataset?.span);
    if (own) return own;
    const nested = element?.querySelector?.("[data-span-id], [data-span]");
    return text(nested?.dataset?.spanId || nested?.dataset?.span);
  }

  function poleIdFromCard(card) {
    return text(card?.dataset?.poleId || card?.dataset?.poleCard);
  }

  function annotateTableColumns(table) {
    if (!table?.querySelectorAll) return [];
    const headers = Array.from(table.querySelectorAll("thead th"));
    const keys = headers.map(header => {
      const key = columnKey(header.textContent);
      header.dataset.column = key;
      return key;
    });

    table.querySelectorAll("tbody tr").forEach(row => {
      Array.from(row.children || []).forEach((cell, index) => {
        if (keys[index]) cell.dataset.column = keys[index];
      });
    });
    return keys;
  }

  function pairCommRows(tableRow) {
    if (!tableRow?.querySelectorAll) return [];
    const spanRows = Array.from(tableRow.querySelectorAll(".comm-span-list .comm-span-row"));
    const midspanRows = Array.from(tableRow.querySelectorAll(".comm-midspan-list .comm-midspan-value"));
    return spanRows.map((spanRow, index) => ({
      spanId: spanIdFrom(spanRow),
      spanRow,
      midspanRow: midspanRows[index] || null
    }));
  }

  function setIdentity(element, poleId, spanId) {
    if (!element?.dataset) return;
    if (poleId) element.dataset.poleId = poleId;
    if (spanId) element.dataset.spanId = spanId;
  }

  function annotateProposedRows(card, poleId) {
    card.querySelectorAll(".span-proposed-table tbody tr").forEach(row => {
      const spanId = spanIdFrom(row);
      setIdentity(row, poleId, spanId);
      row.dataset.component = "proposed-span-row";
      row.querySelectorAll(".span-color-chip, .midspan-highlight-input, .midspan-highlight-display")
        .forEach(element => setIdentity(element, poleId, spanId));
    });
  }

  function annotateCommRows(card, poleId) {
    card.querySelectorAll(".comm-movement-table tbody > tr").forEach(tableRow => {
      tableRow.dataset.component = "comm-movement-row";
      if (poleId) tableRow.dataset.poleId = poleId;
      pairCommRows(tableRow).forEach(pair => {
        setIdentity(pair.spanRow, poleId, pair.spanId);
        pair.spanRow.dataset.component = "comm-span-row";
        pair.spanRow.querySelectorAll(".span-color-dot")
          .forEach(element => setIdentity(element, poleId, pair.spanId));

        if (!pair.midspanRow) return;
        setIdentity(pair.midspanRow, poleId, pair.spanId);
        pair.midspanRow.dataset.component = "comm-midspan-row";
        pair.midspanRow.querySelectorAll(".midspan-highlight-input, .midspan-highlight-display")
          .forEach(element => setIdentity(element, poleId, pair.spanId));
      });
    });
  }

  function annotatePoleCard(card) {
    if (!card?.querySelectorAll || !card.dataset) return false;
    const poleId = poleIdFromCard(card);
    card.dataset.component = "pole-card";
    if (poleId) card.dataset.poleId = poleId;

    card.querySelectorAll("table").forEach(table => annotateTableColumns(table));
    card.querySelectorAll(".span-proposed-table").forEach(table => {
      table.dataset.component = "proposed-span-table";
    });
    card.querySelectorAll(".comm-movement-table").forEach(table => {
      table.dataset.component = "comm-movement-table";
    });
    card.querySelectorAll(".power-table").forEach(table => {
      table.dataset.component = "power-table";
    });
    card.querySelectorAll(".equipment-table").forEach(table => {
      table.dataset.component = "equipment-table";
    });

    annotateProposedRows(card, poleId);
    annotateCommRows(card, poleId);
    return true;
  }

  function apply(root = global.document) {
    if (!root?.querySelectorAll) return 0;
    const cards = Array.from(root.querySelectorAll("[data-pole-card], [data-component='pole-card']"));
    cards.forEach(annotatePoleCard);
    return cards.length;
  }

  global.UiDomContract = {
    normalizeLabel,
    columnKey,
    spanIdFrom,
    pairCommRows,
    annotateTableColumns,
    annotatePoleCard,
    apply
  };
})(window);
