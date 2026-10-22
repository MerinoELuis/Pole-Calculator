(function (global) {
  "use strict";

  function create({ S, Calculations, ExcelReview, markDirty, render, toast, clearPendingRender }) {
    const undoHistory = [];
    const redoHistory = [];
    const undoTrace = [];
    const MAX_UNDO_STEPS = 100;
    const MAX_UNDO_TRACE = 200;
    let undoTraceSequence = 0;
    let restoring = false;

    function cloneState() {
      return JSON.parse(JSON.stringify(S.getState()));
    }

    function traceValue(value) {
      if (value === undefined || value === null || typeof value !== "object") return value ?? null;
      if (Array.isArray(value)) return { type: "array", length: value.length };
      return { type: "object", keys: Object.keys(value).length };
    }

    function collectChanges(before, after, path, changes, metadata) {
      if (Object.is(before, after)) return;
      if (changes.length >= metadata.limit) {
        metadata.truncated = true;
        return;
      }
      const beforeObject = before && typeof before === "object" && !Array.isArray(before);
      const afterObject = after && typeof after === "object" && !Array.isArray(after);
      if (beforeObject && afterObject) {
        const keys = Array.from(new Set([...Object.keys(before), ...Object.keys(after)])).sort();
        keys.forEach(key => collectChanges(
          before[key], after[key], path ? `${path}.${key}` : key, changes, metadata
        ));
        return;
      }
      changes.push({ path: path || "$", before: traceValue(before), after: traceValue(after) });
    }

    function recordTrace(action, source, before, after, details = {}) {
      const changes = [];
      const metadata = { limit: 300, truncated: false };
      collectChanges(before, after, "", changes, metadata);
      undoTrace.push({
        sequence: ++undoTraceSequence,
        timestamp: new Date().toISOString(),
        action,
        source,
        result: changes.length ? "CHANGED" : "NO_CHANGE",
        undoDepthBefore: details.undoDepthBefore ?? undoHistory.length,
        undoDepthAfter: details.undoDepthAfter ?? undoHistory.length,
        redoDepthBefore: details.redoDepthBefore ?? redoHistory.length,
        redoDepthAfter: details.redoDepthAfter ?? redoHistory.length,
        changes,
        changesTruncated: metadata.truncated
      });
      if (undoTrace.length > MAX_UNDO_TRACE) undoTrace.shift();
    }

    function restoreSnapshot(snapshot) {
      restoring = true;
      try {
        clearPendingRender?.();
        S.setState(snapshot);
        Calculations.recalculateAll();
        if (S.getState().excelReviewSource?.collection?.rows?.length) ExcelReview.runReview();
        render();
      } finally {
        restoring = false;
      }
    }

    function recordUndoSnapshot() {
      if (restoring) return;
      const snapshot = cloneState();
      const serialized = JSON.stringify(snapshot);
      const previous = undoHistory[undoHistory.length - 1];
      if (previous && previous.serialized === serialized) return;
      undoHistory.push({ serialized, snapshot });
      if (undoHistory.length > MAX_UNDO_STEPS) undoHistory.shift();
      redoHistory.length = 0;
      markDirty();
    }

    function undoLastAction(source = "Ctrl+Z") {
      const previous = undoHistory.pop();
      const currentSnapshot = cloneState();
      const undoDepthBefore = undoHistory.length + (previous ? 1 : 0);
      const redoDepthBefore = redoHistory.length;
      if (!previous) {
        recordTrace("UNDO", source, currentSnapshot, currentSnapshot, {
          undoDepthBefore, undoDepthAfter: undoHistory.length,
          redoDepthBefore, redoDepthAfter: redoHistory.length
        });
        return toast("No changes to undo.", "info");
      }
      redoHistory.push({ serialized: JSON.stringify(currentSnapshot), snapshot: currentSnapshot });
      if (redoHistory.length > MAX_UNDO_STEPS) redoHistory.shift();
      recordTrace("UNDO", source, currentSnapshot, previous.snapshot, {
        undoDepthBefore, undoDepthAfter: undoHistory.length,
        redoDepthBefore, redoDepthAfter: redoHistory.length
      });
      restoreSnapshot(previous.snapshot);
      markDirty();
      toast("Last change undone.", "success");
    }

    function redoLastAction(source = "Ctrl+Y") {
      const next = redoHistory.pop();
      const currentSnapshot = cloneState();
      const undoDepthBefore = undoHistory.length;
      const redoDepthBefore = redoHistory.length + (next ? 1 : 0);
      if (!next) {
        recordTrace("REDO", source, currentSnapshot, currentSnapshot, {
          undoDepthBefore, undoDepthAfter: undoHistory.length,
          redoDepthBefore, redoDepthAfter: redoHistory.length
        });
        return toast("No changes to redo.", "info");
      }
      undoHistory.push({ serialized: JSON.stringify(currentSnapshot), snapshot: currentSnapshot });
      if (undoHistory.length > MAX_UNDO_STEPS) undoHistory.shift();
      recordTrace("REDO", source, currentSnapshot, next.snapshot, {
        undoDepthBefore, undoDepthAfter: undoHistory.length,
        redoDepthBefore, redoDepthAfter: redoHistory.length
      });
      restoreSnapshot(next.snapshot);
      markDirty();
      toast("Last change redone.", "success");
    }

    function captureTransaction() {
      return { undoHistory: undoHistory.slice(), redoHistory: redoHistory.slice() };
    }

    function restoreTransaction(snapshot) {
      if (!snapshot) return;
      undoHistory.splice(0, undoHistory.length, ...(snapshot.undoHistory || []));
      redoHistory.splice(0, redoHistory.length, ...(snapshot.redoHistory || []));
    }

    global.AppUndoTrace = { getTrace: () => undoTrace.map(entry => JSON.parse(JSON.stringify(entry))) };
    return { cloneState, recordUndoSnapshot, undoLastAction, redoLastAction, restoreSnapshot, captureTransaction, restoreTransaction };
  }

  global.AppHistory = Object.freeze({ create });
})(window);
