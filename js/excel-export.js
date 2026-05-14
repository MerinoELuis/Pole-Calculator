(function (global) {
  "use strict";

  const S = () => global.AppStore;

  function stateToSheets() {
    global.Calculations.recalculateAll();
    const state = S().getState();

    const poles = [["poleId", "collectionId", "sequence", "poleHeight", "tipHeight", "lowPower", "maxCommHeight", "topComm", "lowComm", "owner", "poleType", "isGenerated", "commCount", "spanCount", "warningCount", "hasChanges", "notes"]];
    Object.values(state.poles).forEach(pole => {
      poles.push([
        pole.poleId,
        pole.collectionId || "",
        pole.sequence || "",
        pole.poleHeight || "",
        pole.tipHeight || "",
        pole.lowPower || "",
        pole.maxCommHeight || "",
        pole.topComm || "",
        pole.lowComm || "",
        pole.owner || "",
        pole.poleType || "",
        pole.isGenerated ? "Yes" : "No",
        pole.comms.length,
        S().getConnectedSpans(pole.poleId).length,
        state.warnings.filter(w => w.poleId === pole.poleId).length,
        S().poleHasChanges(pole.poleId) ? "Yes" : "No",
        pole.notes || ""
      ]);
    });

    const spans = [["spanId", "fromPole", "toPole", "direction", "bearingDegrees", "type", "spanIndex", "length", "lengthDisplay", "environment", "environmentClearance", "midspanLowPower", "midspanMaxCommHeight", "rawType", "rawSpanIds", "sourceCollectionId", "linkedCollectionId", "isGeneratedOtherPole", "notes"]];
    Object.values(state.spans).forEach(span => spans.push([
      span.spanId,
      span.fromPole,
      span.toPole,
      span.direction,
      span.bearingDegrees ?? "",
      span.type || "",
      span.spanIndex || "",
      span.length || "",
      span.lengthDisplay || "",
      span.environment || "NONE",
      span.environmentClearance || "",
      span.midspanLowPower || "",
      span.midspanMaxCommHeight || "",
      span.rawType || "",
      Array.isArray(span.rawSpanIds) ? span.rawSpanIds.join(" | ") : "",
      span.sourceCollectionId || "",
      span.linkedCollectionId || "",
      span.isGeneratedOtherPole ? "Yes" : "No",
      span.notes || ""
    ]));

    const spanSides = [["spanId", "poleId", "proposedHOA", "proposedHOAChange", "proposedMidspan", "ocalcMS", "msProposed", "finalMidspan", "clearanceMSReason", "endDrop", "clearanceReference", "maxCommHeight", "topComm", "lowComm", "notes"]];
    Object.values(state.spanSides).forEach(side => spanSides.push([
      side.spanId,
      side.poleId,
      side.proposedHOA,
      side.proposedHOAChange || "",
      side.proposedMidspan,
      side.ocalcMS || "",
      side.msProposed || "",
      side.finalMidspan || "",
      side.clearanceMSReason || "",
      side.endDrop,
      side.clearanceReference,
      side.maxCommHeight,
      side.topComm,
      side.lowComm,
      side.notes || ""
    ]));

    const comms = [["poleId", "owner", "ownerBase", "existingHOA", "existingHOAChange", "rawOwner", "unknownOwner", "size", "wireId", "notes"]];
    Object.values(state.poles).forEach(pole => pole.comms.forEach(comm => comms.push([
      pole.poleId,
      comm.owner,
      comm.ownerBase || comm.owner,
      comm.existingHOA,
      comm.existingHOAChange || "",
      comm.rawOwner || "",
      comm.unknownOwner ? "Yes" : "No",
      comm.size || "",
      comm.wireId || "",
      comm.notes || ""
    ])));

    const spanComms = [["spanId", "poleId", "owner", "ownerBase", "existingHOA", "existingHOAChange", "difference", "remotePoleId", "remoteHOA", "ocalcMS", "midspan", "calculatedMidspan", "mr", "notes", "rawOwner", "unknownOwner", "size", "construction", "insulator", "wireId", "wireIndex"]];
    Object.values(state.spanComms).forEach(sc => spanComms.push([
      sc.spanId,
      sc.poleId,
      sc.owner,
      sc.ownerBase || sc.owner,
      sc.existingHOA,
      sc.existingHOAChange,
      sc.difference,
      sc.remotePoleId,
      sc.remoteHOA,
      sc.ocalcMS,
      sc.midspan,
      sc.calculatedMidspan,
      sc.mr,
      sc.notes,
      sc.rawOwner || "",
      sc.unknownOwner ? "Yes" : "No",
      sc.size || "",
      sc.construction || "",
      sc.insulator || "",
      sc.wireId || "",
      sc.wireIndex || ""
    ]));

    const spanPower = [["spanId", "poleId", "label", "attachmentHeight", "midspan", "size", "owner", "wireId"]];
    Object.values(state.spanPower).forEach(row => spanPower.push([
      row.spanId,
      row.poleId,
      row.label,
      row.attachmentHeight,
      row.midspan,
      row.size,
      row.owner,
      row.wireId
    ]));

    const midspans = [["spanId", "span", "poleId", "owner", "existingHOA", "existingHOAChange", "remotePoleId", "remoteHOA", "importedMidspan", "ocalcMS", "calculatedMidspan"]];
    Object.values(state.spanComms).forEach(sc => {
      const span = S().getSpan(sc.spanId);
      midspans.push([sc.spanId, span ? `${span.fromPole} → ${span.toPole}` : "", sc.poleId, sc.owner, sc.existingHOA, sc.existingHOAChange, sc.remotePoleId, sc.remoteHOA, sc.midspan, sc.ocalcMS, sc.calculatedMidspan]);
    });

    const movements = [["poleId", "spanId", "owner", "existingHOA", "existingHOAChange", "difference", "midspanCalculated"]];
    Object.values(state.spanComms).forEach(sc => {
      if (sc.existingHOAChange) movements.push([sc.poleId, sc.spanId, sc.owner, sc.existingHOA, sc.existingHOAChange, sc.difference, sc.calculatedMidspan]);
    });

    const mr = [["poleId", "mr"]];
    state.mr.forEach(item => mr.push([item.poleId, item.text]));

    const warnings = [["poleId", "spanId", "owner", "code", "level", "message"]];
    state.warnings.forEach(w => warnings.push([w.poleId, w.spanId, w.owner, w.code, w.level, w.message]));

    const appState = [["json"], [JSON.stringify(state)]];

    return [
      { name: "Poles", rows: poles },
      { name: "Spans", rows: spans },
      { name: "SpanSides", rows: spanSides },
      { name: "Comms", rows: comms },
      { name: "SpanComms", rows: spanComms },
      { name: "SpanPower", rows: spanPower },
      { name: "Midspans", rows: midspans },
      { name: "Movements", rows: movements },
      { name: "MR", rows: mr },
      { name: "Warnings", rows: warnings },
      { name: "AppState", rows: appState }
    ];
  }

  function exportExcel() {
    const date = new Date().toISOString().slice(0, 10);
    global.MiniXLSX.writeFile(`pole-calculator-save-${date}.xlsx`, stateToSheets());
  }

  function downloadJson(filename, data) {
    const blob = new Blob([JSON.stringify(data, null, 2)], { type: "application/json;charset=utf-8" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    a.remove();
    URL.revokeObjectURL(url);
  }

  function exportJson() {
    global.Calculations.recalculateAll();
    const state = S().getState();
    const date = new Date().toISOString().slice(0, 10);
    downloadJson(`pole-calculator-datos-${date}.json`, {
      app: "pole-calculator",
      exportedAt: new Date().toISOString(),
      version: state.version || S().CURRENT_VERSION,
      state
    });
  }

  function rowsToCsv(rows) {
    return rows.map(row => row.map(value => {
      const text = String(value ?? "");
      if (/[",\n]/.test(text)) return `"${text.replace(/"/g, '""')}"`;
      return text;
    }).join(",")).join("\n");
  }

  function downloadCsv(filename, rows) {
    const blob = new Blob([rowsToCsv(rows)], { type: "text/csv;charset=utf-8" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    a.remove();
    URL.revokeObjectURL(url);
  }

  function exportCsv(type) {
    const sheets = stateToSheets();
    const map = {
      results: "Movements",
      midspans: "Midspans",
      mr: "MR",
      warnings: "Warnings"
    };
    const sheet = sheets.find(s => s.name === map[type]);
    if (!sheet) return;
    downloadCsv(`${map[type].toLowerCase()}-${new Date().toISOString().slice(0, 10)}.csv`, sheet.rows);
  }

  global.ExcelExport = {
    stateToSheets,
    exportExcel,
    exportJson,
    exportData: exportExcel,
    exportCsv,
    downloadJson,
    downloadCsv
  };
})(window);
