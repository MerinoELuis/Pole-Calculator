(function (global) {
  "use strict";

  const MODES = Object.freeze(["AERIAL", "UG", "PCO"]);

  function normalize(value) {
    const mode = String(value || "").trim().toUpperCase();
    return MODES.includes(mode) ? mode : "AERIAL";
  }

  function forPole(pole) {
    if (!pole) return "AERIAL";
    const resolutionAlias = Object.getOwnPropertyDescriptor(pole, "poleResolution");
    if (resolutionAlias?.get) {
      if (pole.ugActive) return "UG";
      if (pole.pcoActive) return "PCO";
      return "AERIAL";
    }
    const explicit = String(pole.poleResolution || "").trim().toUpperCase();
    if (MODES.includes(explicit)) return explicit;
    if (pole.ugActive) return "UG";
    if (pole.pcoActive) return "PCO";
    return "AERIAL";
  }

  function flags(value) {
    const mode = normalize(value);
    return { ugActive: mode === "UG", pcoActive: mode === "PCO" };
  }

  function apply(pole, value) {
    if (!pole || typeof pole !== "object") return pole;
    Object.assign(pole, flags(value));
    return pole;
  }

  global.PoleResolution = Object.freeze({ MODES, normalize, forPole, flags, apply });
})(window);
