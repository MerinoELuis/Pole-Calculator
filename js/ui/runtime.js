(function (global) {
  "use strict";

  let observer = null;
  let refreshQueued = false;

  function runRefresh(root = global.document) {
    if (!root) return { cards: 0, tablesChanged: 0, coloredSpans: 0, autoCalculateStatuses: 0 };
    const cards = global.UiDomContract?.apply?.(root) || 0;
    const tablesChanged = global.CommTableUI?.refresh?.(root) || 0;
    if (tablesChanged) global.UiDomContract?.apply?.(root);
    const coloredSpans = global.SpanColorUI?.refresh?.(root, global.AppStore) || 0;
    const autoCalculateStatuses = global.AutoCalculateStatusUI?.refresh?.(root) || 0;
    return { cards, tablesChanged, coloredSpans, autoCalculateStatuses };
  }

  function queueRefresh() {
    if (refreshQueued) return;
    refreshQueued = true;
    const run = () => {
      refreshQueued = false;
      runRefresh();
    };
    if (typeof global.queueMicrotask === "function") global.queueMicrotask(run);
    else global.setTimeout?.(run, 0);
  }

  function start() {
    const document = global.document;
    if (!document) return false;
    queueRefresh();
    if (observer || typeof global.MutationObserver !== "function") return true;
    const root = document.getElementById("polesOverview") || document.body || document.documentElement;
    if (!root) return true;
    observer = new global.MutationObserver(queueRefresh);
    observer.observe(root, { childList: true, subtree: true });
    return true;
  }

  function stop() {
    observer?.disconnect?.();
    observer = null;
    refreshQueued = false;
  }

  if (global.document?.readyState === "loading") {
    global.document.addEventListener("DOMContentLoaded", start, { once: true });
  } else {
    start();
  }

  global.PoleCalculatorUI = {
    runRefresh,
    queueRefresh,
    start,
    stop
  };
})(window);
