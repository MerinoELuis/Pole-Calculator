(function (global) {
  "use strict";

  // Translate the historical enumerable UG/PCO flags and the newer canonical
  // poleResolution field at the store boundary. Keep this migration separate
  // from the in-memory store so saved JSON remains backward compatible.
  function normalizePoleInput(source = {}) {
    const resolution = global.PoleResolution.forPole(source);
    return {
      ...source,
      ...global.PoleResolution.flags(resolution)
    };
  }

  function unwrapSavedState(value) {
    if (!value || typeof value !== "object") return value;
    if (value.state && typeof value.state === "object") return value.state;
    if (value.data?.state && typeof value.data.state === "object") return value.data.state;
    return value;
  }

  global.StateCompatibility = Object.freeze({ normalizePoleInput, unwrapSavedState });
})(window);
