(function (global) {
  "use strict";

  const store = global.AppStore;
  if (!store || store.__autoCalculateSourceCompatInstalled) return;

  const PROPOSED_SOURCE_FIELDS = [
    "autoCalcProposedStatus",
    "autoCalcProposedMode"
  ];

  function text(value) {
    return String(value ?? "").trim();
  }

  function keyForSide(side = {}) {
    return store.keyForSpanSide?.(side.spanId || "", side.poleId || "")
      || `${text(side.spanId)}__${text(side.poleId)}`;
  }

  function copySourceFields(target, source) {
    if (!target || !source) return target;
    PROPOSED_SOURCE_FIELDS.forEach(field => {
      if (!Object.prototype.hasOwnProperty.call(source, field)) return;
      target[field] = text(source[field]);
    });
    return target;
  }

  const originalUpsertSpanSide = store.upsertSpanSide?.bind(store);
  if (originalUpsertSpanSide) {
    store.upsertSpanSide = function (data = {}) {
      const result = originalUpsertSpanSide(data);
      return copySourceFields(result, data);
    };
  }

  const originalSetState = store.setState?.bind(store);
  if (originalSetState) {
    store.setState = function (nextState) {
      const sourceByKey = new Map(
        Object.values(nextState?.spanSides || {}).map(side => [keyForSide(side), side])
      );
      const normalized = originalSetState(nextState);
      Object.values(normalized?.spanSides || store.getState?.()?.spanSides || {}).forEach(side => {
        copySourceFields(side, sourceByKey.get(keyForSide(side)));
      });
      return normalized;
    };
  }

  store.__autoCalculateSourceCompatInstalled = true;

  global.AutoCalculateSourceCompat = {
    PROPOSED_SOURCE_FIELDS,
    keyForSide,
    copySourceFields
  };
})(window);
