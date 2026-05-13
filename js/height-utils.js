(function (global) {
  "use strict";

  function normalizeText(value) {
    return String(value ?? "")
      .trim()
      .replace(/[’‘]/g, "'")
      .replace(/[“”]/g, '"')
      .replace(/feet|foot/gi, "ft")
      .replace(/inches|inch/gi, "in")
      .replace(/\s+/g, " ");
  }

  function parseHeight(value) {
    if (value === null || value === undefined || value === "") return null;
    if (typeof value === "number" && Number.isFinite(value)) return Math.round(value * 12);

    let text = normalizeText(value);
    if (!text) return null;
    if (text === "0" || text === "0\"" || text === "0'") return 0;

    let sign = 1;
    if (/^[-−]/.test(text)) {
      sign = -1;
      text = text.replace(/^[-−]\s*/, "");
    }

    const decimalFeetMatch = text.match(/^(\d+(?:\.\d+)?)\s*ft$/i);
    if (decimalFeetMatch) return sign * Math.round(Number(decimalFeetMatch[1]) * 12);

    const onlyDecimal = text.match(/^\d+\.\d+$/);
    if (onlyDecimal) return sign * Math.round(Number(text) * 12);

    const ftInMatch = text.match(/^(?:(\d+)\s*')?\s*(?:(\d+(?:\.\d+)?)\s*(?:"|in)?)?$/i);
    if (ftInMatch && (ftInMatch[1] !== undefined || ftInMatch[2] !== undefined)) {
      const feet = Number(ftInMatch[1] || 0);
      const inches = Number(ftInMatch[2] || 0);
      return sign * Math.round(feet * 12 + inches);
    }

    const inchesOnly = text.match(/^(\d+(?:\.\d+)?)\s*(?:"|in)$/i);
    if (inchesOnly) return sign * Math.round(Number(inchesOnly[1]));

    return null;
  }

  function formatHeight(inches, options = {}) {
    if (inches === null || inches === undefined || Number.isNaN(Number(inches))) return "";
    const rounded = Math.round(Number(inches));
    const sign = rounded < 0 ? "-" : "";
    const abs = Math.abs(rounded);
    const feet = Math.floor(abs / 12);
    const inch = abs % 12;
    if (abs === 0) return "0\"";
    if (feet === 0) return `${sign}${inch}\"`;
    if (inch === 0 && options.hideZeroInches !== false) return `${sign}${feet}'`;
    return `${sign}${feet}'${inch}\"`;
  }

  function addHeights(a, b) {
    const ia = parseHeight(a);
    const ib = parseHeight(b);
    if (ia === null || ib === null) return "";
    return formatHeight(ia + ib);
  }

  function subtractHeights(a, b) {
    const ia = parseHeight(a);
    const ib = parseHeight(b);
    if (ia === null || ib === null) return "";
    return formatHeight(ia - ib);
  }

  function decimalFeetToHeight(value) {
    if (value === null || value === undefined || value === "") return "";
    const n = Number(String(value).replace(/ft/i, "").trim());
    if (!Number.isFinite(n)) return "";
    return formatHeight(Math.round(n * 12));
  }

  function heightToDecimalFeet(value) {
    const inches = parseHeight(value);
    if (inches === null) return "";
    return Number((inches / 12).toFixed(3));
  }

  function inchesToHeight(value) {
    if (value === null || value === undefined || value === "") return "";
    const n = Number(String(value).replace(/in|"/i, "").trim());
    if (!Number.isFinite(n)) return "";
    return formatHeight(Math.round(n));
  }

  function compareHeights(a, b) {
    const ia = parseHeight(a);
    const ib = parseHeight(b);
    if (ia === null || ib === null) return null;
    return ia === ib ? 0 : ia > ib ? 1 : -1;
  }

  function isValidHeight(value, allowBlank = true) {
    if ((value === null || value === undefined || value === "") && allowBlank) return true;
    return parseHeight(value) !== null;
  }

  function diffLabel(from, to) {
    const a = parseHeight(from);
    const b = parseHeight(to);
    if (a === null || b === null) return "";
    return formatHeight(b - a);
  }

  global.HeightUtils = {
    parseHeight,
    formatHeight,
    addHeights,
    subtractHeights,
    decimalFeetToHeight,
    heightToDecimalFeet,
    inchesToHeight,
    compareHeights,
    isValidHeight,
    diffLabel
  };
})(window);
