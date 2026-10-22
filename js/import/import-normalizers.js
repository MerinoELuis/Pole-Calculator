(function (global) {
  "use strict";

  function create(heightUtils) {
    const H = () => heightUtils || global.HeightUtils;

    function rowsToObjects(rows) {
      if (!Array.isArray(rows) || rows.length === 0) return [];
      const headerRow = rows.find(row => Array.isArray(row) && row.some(v => v !== undefined && v !== null && String(v).trim() !== ""));
      if (!headerRow) return [];
      const headerIndex = rows.indexOf(headerRow);
      const headers = headerRow.map(h => String(h ?? "").trim());

      return rows.slice(headerIndex + 1)
        .filter(row => Array.isArray(row) && row.some(v => v !== undefined && v !== null && String(v).trim() !== ""))
        .map(row => {
          const obj = {};
          headers.forEach((header, index) => {
            if (header) obj[header] = row[index] ?? "";
          });
          return obj;
        });
    }

    function normalizeHeaderName(name) {
      return String(name || "")
        .toLowerCase()
        .normalize("NFD")
        .replace(/[\u0300-\u036f]/g, "")
        .replace(/[._\-\s()\[\]/]/g, "")
        .replace(/[^a-z0-9]/g, "");
    }

    function isBlank(value) {
      return value === undefined || value === null || String(value).trim() === "";
    }

    function normalizeNumericProjectIdentifiers(rows) {
      if (!Array.isArray(rows)) return rows;

      const poleIdentifierFields = new Set([
        "id", "poleid", "pole", "frompole", "topole", "otherpole", "remotepole",
        "structurenumber", "collectionid", "linkedcollectionid"
      ]);
      rows.forEach(row => {
        Object.keys(row || {}).forEach(key => {
          const normalizedKey = normalizeHeaderName(key);
          const value = row[key];
          if (normalizedKey === "sequence" || normalizedKey === "seq" || poleIdentifierFields.has(normalizedKey)) {
            const textValue = String(value ?? "").trim();
            if (/^\d{1,2}$/.test(textValue)) row[key] = textValue.padStart(3, "0");
          }
        });
      });
      return rows;
    }

    function sheetSnapshot(sheet, normalizedRows) {
      const rows = Array.isArray(sheet) ? sheet : [];
      const headerRow = rows.find(row => Array.isArray(row) && row.some(value => !isBlank(value)));
      return {
        headers: headerRow ? headerRow.map(value => String(value ?? "").trim()).filter(Boolean) : [],
        rows: Array.isArray(normalizedRows) ? normalizedRows : rowsToObjects(rows)
      };
    }

    function truthyCell(value) {
      return /^(true|yes|1|si|sí)$/i.test(String(value || "").trim());
    }

    function hasBirthmarkValue(value) {
      if (typeof value === "boolean") return value;
      const text = String(value ?? "").trim();
      if (!text) return false;
      return !/^(0|false|no|n|none|null|n\/a)$/i.test(text);
    }

    function findKey(row, names, options = {}) {
      if (!row) return "";
      const keys = Object.keys(row);
      const wanted = names.map(name => ({ raw: name, normalized: normalizeHeaderName(name) })).filter(item => item.normalized);

      for (const item of wanted) {
        const exact = keys.find(k => normalizeHeaderName(k) === item.normalized);
        if (exact) return exact;
      }

      if (options.contains !== false) {
        for (const item of wanted) {
          const partial = keys.find(k => {
            const normalizedKey = normalizeHeaderName(k);
            const minimumLength = Number(options.minPartialLength || 0);
            if (minimumLength > 0
              && (normalizedKey.length < minimumLength || item.normalized.length < minimumLength)) return false;
            return normalizedKey.includes(item.normalized) || item.normalized.includes(normalizedKey);
          });
          if (partial) return partial;
        }
      }

      return "";
    }

    function pick(row, names, options = {}) {
      const key = findKey(row, names, options);
      return key ? row[key] : "";
    }

    function findSheet(workbook, preferredNames) {
      const normalized = preferredNames.map(normalizeHeaderName);
      const name = workbook.SheetNames.find(sheetName => normalized.includes(normalizeHeaderName(sheetName)));
      return name ? workbook.Sheets[name] : null;
    }

    function cleanHeight(value) {
      if (isBlank(value)) return "";
      const parsed = H().parseHeight(value);
      return parsed === null ? String(value).trim() : H().formatHeight(parsed);
    }

    function heightFromMeters(value) {
      if (isBlank(value)) return "";
      const n = Number(value);
      if (!Number.isFinite(n)) return cleanHeight(value);
      return H().formatHeight(Math.round(n * 3.280839895 * 12));
    }

    function pickDisplayValue(row, displayNames) {
      if (!row) return "";
      const keys = Object.keys(row);
      const wanted = displayNames.map(normalizeHeaderName).filter(Boolean);

      const exactDisplay = keys.find(k => wanted.includes(normalizeHeaderName(k))
        && normalizeHeaderName(k).includes("display")
        && !isBlank(row[k]));
      if (exactDisplay) return row[exactDisplay];

      const partialDisplay = keys.find(k => {
        const normalizedKey = normalizeHeaderName(k);
        if (!normalizedKey.includes("display")) return false;
        if (isBlank(row[k])) return false;
        return wanted.some(w => {
          const base = w.replace(/display/g, "");
          return normalizedKey.includes(base) || base.includes(normalizedKey.replace(/display/g, ""));
        });
      });
      return partialDisplay ? row[partialDisplay] : "";
    }

    function heightFromRow(row, displayNames, decimalNames) {
      const display = pickDisplayValue(row, displayNames);
      if (!isBlank(display)) return cleanHeight(display);
      const decimal = pick(row, decimalNames, { contains: true, minPartialLength: 3 });
      return heightFromMeters(decimal);
    }

    function parsePoleHeightFromType(typeValue) {
      const text = String(typeValue || "");
      const parts = text.split(">").map(p => p.trim()).filter(Boolean);
      const last = parts.length ? parts[parts.length - 1] : text;
      const match = last.match(/\d+(?:\.\d+)?/);
      if (!match) return "";
      const feet = Number(match[0]);
      if (!Number.isFinite(feet)) return "";
      return H().formatHeight(Math.round(feet * 12));
    }

    return {
      rowsToObjects,
      normalizeNumericProjectIdentifiers,
      sheetSnapshot,
      normalizeHeaderName,
      isBlank,
      truthyCell,
      hasBirthmarkValue,
      findKey,
      pick,
      findSheet,
      cleanHeight,
      heightFromMeters,
      pickDisplayValue,
      heightFromRow,
      parsePoleHeightFromType
    };
  }

  global.ImportNormalizers = { create };
})(window);
