(function (global) {
  "use strict";

  // Lightweight performance tracing for real browser jobs. Samples are kept
  // in memory for the current session and copied into the local Debug JSON;
  // normal users do not receive console output unless verbose mode is enabled.
  const MAX_RECENT_SAMPLES = 200;
  let enabled = true;
  let verbose = false;
  const totals = new Map();
  const recent = [];

  function now() {
    return typeof global.performance?.now === "function"
      ? global.performance.now()
      : Date.now();
  }

  function rounded(value) {
    return Math.round(Number(value || 0) * 100) / 100;
  }

  function safeDetails(details) {
    if (!details || typeof details !== "object") return details ?? null;
    const result = {};
    Object.entries(details).forEach(([key, value]) => {
      if (["string", "number", "boolean"].includes(typeof value) || value === null) {
        result[key] = value;
      } else if (Array.isArray(value)) {
        result[key] = value.length;
      }
    });
    return result;
  }

  function record(label, durationMs, details = null) {
    if (!enabled) return;
    const duration = Math.max(0, Number(durationMs) || 0);
    const key = String(label || "unknown");
    const entry = totals.get(key) || { count: 0, totalMs: 0, maxMs: 0, lastMs: 0 };
    entry.count += 1;
    entry.totalMs = rounded(entry.totalMs + duration);
    entry.maxMs = Math.max(entry.maxMs, rounded(duration));
    entry.lastMs = rounded(duration);
    totals.set(key, entry);

    recent.push({ label: key, durationMs: rounded(duration), details: safeDetails(details) });
    if (recent.length > MAX_RECENT_SAMPLES) recent.shift();
    if (verbose) console.debug("[PoleCalc perf]", key, `${rounded(duration)} ms`, details || "");
  }

  function measure(label, callback, details = null) {
    if (!enabled) return callback();
    const startedAt = now();
    try {
      return callback();
    } finally {
      record(label, now() - startedAt, typeof details === "function" ? details() : details);
    }
  }

  function reset() {
    totals.clear();
    recent.length = 0;
    return snapshot();
  }

  function snapshot() {
    return {
      enabled,
      verbose,
      totals: Object.fromEntries(Array.from(totals.entries()).map(([label, value]) => [label, { ...value }])),
      recent: recent.map(sample => ({ ...sample, details: sample.details && { ...sample.details } }))
    };
  }

  function report() {
    const result = snapshot();
    if (typeof console !== "undefined") {
      console.info("[PoleCalc perf] summary", result);
      if (typeof console.table === "function") console.table(result.totals);
    }
    return result;
  }

  function enable(options = {}) {
    enabled = true;
    verbose = options.verbose === true;
    if (options.reset !== false) reset();
    if (typeof console !== "undefined") {
      console.info("[PoleCalc perf] enabled. Run PoleCalcPerformance.report() to inspect timings.");
    }
    return snapshot();
  }

  function disable() {
    enabled = false;
    verbose = false;
    return snapshot();
  }

  global.PoleCalcPerformance = {
    enable,
    disable,
    reset,
    measure,
    record,
    report,
    snapshot,
    isEnabled: () => enabled
  };
})(window);
