(function (global) {
  "use strict";

  const CASES = Object.freeze({
    MEDIUM: {
      caseId: "NESC_MEDIUM_250B_GRADE_C",
      name: "NESC Medium (250B) Grade C",
      rule: "250B",
      grade: "C",
      role: "BASE",
      windPressurePsf: 4,
      radialIceInches: 0.25
    },
    LIGHT: {
      caseId: "NESC_23_LIGHT_250B_GRADE_C",
      name: "NESC 23 Light (250B) Grade C",
      rule: "250B",
      grade: "C",
      role: "BASE",
      windPressurePsf: 9,
      radialIceInches: 0
    },
    PRE_2018: {
      caseId: "NESC_23_EXT_WIND_250C_GRADE_C_LE_100_MPH",
      name: "NESC 23 Ext Wind (250C) Grade C (<= 100 mph)",
      rule: "250C",
      grade: "C",
      role: "BASE",
      windSpeedMph: 90,
      radialIceInches: 0
    },
    EXT_WIND: {
      caseId: "NESC_23_EXT_WIND_250C_GRADE_C_GT_100_MPH",
      name: "NESC 23 Ext Wind (250C) Grade C (> 100 mph)",
      rule: "250C",
      grade: "C",
      role: "ADDITIONAL",
      radialIceInches: 0
    },
    EXT_WIND_ICE: {
      caseId: "NESC_23_EXT_WIND_ICE_250D_GRADE_C",
      name: "NESC 23 Ext Wind & Ice (250D) Grade C",
      rule: "250D",
      grade: "C",
      role: "ADDITIONAL",
      windSpeedMph: 30,
      radialIceInches: 0.25
    }
  });

  function text(value) {
    return String(value ?? "").trim();
  }

  function canonicalPoleId(value) {
    return global.AppStore?.canonicalPoleIdentity?.(value)
      || text(value).toUpperCase().replace(/\s+(STEEL|RECOLLECT|T|UG|PCO)(?:\s+(STEEL|RECOLLECT|T|UG|PCO))*$/i, "").trim();
  }

  function parseYear(value) {
    const match = text(value).match(/\b(\d{4})\b/);
    if (!match) return null;
    const year = Number(match[1]);
    return Number.isInteger(year) && year >= 1800 && year <= 2100 ? year : null;
  }

  function yearFromCollection(state, poleId) {
    const pick = global.ExcelImport?.pick;
    if (typeof pick !== "function") return "";
    const aliases = state?.poleIdAliases || {};
    const canonical = canonicalPoleId(poleId);
    const source = state?.excelReviewSource?.collection?.rows || [];
    const row = source.find(item => {
      const rawId = text(pick(item, ["Id", "Pole ID", "PoleId", "PoleName", "Structure Number", "Pole"], { contains: false }));
      return canonicalPoleId(aliases[rawId] || rawId) === canonical;
    });
    return row ? text(pick(row, ["Year Installed"], { contains: false })) : "";
  }

  function collectionPoleIds(state) {
    const aliases = state?.poleIdAliases || {};
    const pick = global.ExcelImport?.pick;
    const sourceRows = state?.excelReviewSource?.collection?.rows || [];
    const ids = sourceRows.map(row => {
      const rawId = typeof pick === "function"
        ? text(pick(row, ["Id", "Pole ID", "PoleId", "PoleName", "Structure Number", "Pole"], { contains: false }))
        : "";
      return canonicalPoleId(aliases[rawId] || rawId);
    }).filter(Boolean);
    // Pole Type Check rows are also built only from Collection and provide a
    // compatibility fallback for older saves that predate raw review snapshots.
    if (!ids.length) {
      (state?.poleClassChecks || []).forEach(row => {
        const poleId = canonicalPoleId(aliases[text(row?.poleId)] || row?.poleId);
        if (poleId) ids.push(poleId);
      });
    }
    return new Set(ids);
  }

  function getYearInfo(state, poleId) {
    const pole = state?.poles?.[poleId];
    const imported = text((state?.poleClassChecks || []).find(row => canonicalPoleId(row?.poleId) === canonicalPoleId(poleId))?.yearInstalled)
      || yearFromCollection(state, poleId);
    const importedYear = parseYear(imported);
    if (importedYear !== null) return { year: importedYear, source: "excel", inputValue: String(importedYear) };
    const manualValue = text(pole?.metadata?.plaExisting?.yearInstalled);
    const manualYear = parseYear(manualValue);
    return {
      year: manualYear,
      source: manualYear === null ? "missing" : "manual",
      inputValue: manualYear === null ? manualValue : String(manualYear)
    };
  }

  function sortedExistingPoles(state) {
    const collectionIds = collectionPoleIds(state);
    return Object.values(state?.poles || {})
      .filter(pole => pole
        && text(pole.poleId)
        && !pole.isGenerated
        && !/^Unknown-/i.test(text(pole.poleId))
        && collectionIds.has(canonicalPoleId(pole.poleId)))
      .sort((left, right) => text(left.sequence).localeCompare(text(right.sequence), undefined, { numeric: true })
        || text(left.poleId).localeCompare(text(right.poleId), undefined, { numeric: true, sensitivity: "base" }));
  }

  function loadCase(base, additions = {}) {
    return { ...base, ...additions };
  }

  function loadingPrincipal(state) {
    return String(state?.settings?.plaExistingLoadingProfile || "MEDIUM").toUpperCase() === "LIGHT" ? "LIGHT" : "MEDIUM";
  }

  function planForPole(state, pole) {
    const yearInfo = getYearInfo(state, pole.poleId);
    const settings = pole.metadata?.plaExisting || {};
    if (yearInfo.year === null) {
      return {
        id: pole.poleId,
        yearInstalled: null,
        yearSource: yearInfo.source,
        yearInputValue: yearInfo.inputValue,
        loadingProfile: null,
        windSpeedMphInput: text(settings.windSpeedMph),
        loadings: [],
        status: "YEAR_REQUIRED"
      };
    }

    if (yearInfo.year < 2018) {
      return {
        id: pole.poleId,
        yearInstalled: yearInfo.year,
        yearSource: yearInfo.source,
        yearInputValue: String(yearInfo.year),
        loadingProfile: "PRE_2018",
        windSpeedMphInput: "",
        loadings: [loadCase(CASES.PRE_2018)],
        status: "READY"
      };
    }

    const loadingProfile = loadingPrincipal(state);
    const rawWindSpeed = text(settings.windSpeedMph);
    const parsedWindSpeed = rawWindSpeed ? Number(rawWindSpeed) : NaN;
    const validWindSpeed = Number.isFinite(parsedWindSpeed) && parsedWindSpeed > 100;
    const loadings = [loadCase(loadingProfile === "LIGHT" ? CASES.LIGHT : CASES.MEDIUM)];
    if (validWindSpeed) loadings.push(loadCase(CASES.EXT_WIND, { windSpeedMph: parsedWindSpeed }));
    if (loadingProfile === "MEDIUM") loadings.push(loadCase(CASES.EXT_WIND_ICE));
    return {
      id: pole.poleId,
      yearInstalled: yearInfo.year,
      yearSource: yearInfo.source,
      yearInputValue: String(yearInfo.year),
      loadingProfile,
      windSpeedMphInput: rawWindSpeed,
      loadings,
      status: validWindSpeed ? "READY" : "WIND_SPEED_REQUIRED"
    };
  }

  function inspect(state) {
    const profile = String(state?.settings?.projectProfile || "").toUpperCase();
    const plans = sortedExistingPoles(state).map(pole => planForPole(state, pole));
    const missingYearPoleIds = plans.filter(plan => plan.status === "YEAR_REQUIRED").map(plan => plan.id);
    const missingWindPoleIds = plans.filter(plan => plan.status === "WIND_SPEED_REQUIRED").map(plan => plan.id);
    const readyPoleCount = plans.filter(plan => plan.status === "READY").length;
    let reason = "";
    if (profile !== "INTEC") reason = "AutoExisting is available only for INTEC.";
    else if (!plans.length) reason = "Import existing pole data before exporting AutoExisting.";
    else if (missingYearPoleIds.length) reason = `Enter Year Installed for: ${missingYearPoleIds.join(", ")}.`;
    else if (missingWindPoleIds.length) reason = `Enter a wind speed greater than 100 mph for: ${missingWindPoleIds.join(", ")}.`;
    else if (!readyPoleCount) reason = "No poles are ready for AutoExisting export.";

    return {
      canExport: !reason,
      reason,
      plans,
      missingYearPoleIds,
      missingWindPoleIds,
      payload: {
        format: "AutoExisting",
        version: 1,
        project: "INTEC",
        job: text(state?.jobName),
        loadingPrincipal: loadingPrincipal(state),
        poles: plans.map(plan => ({
          id: plan.id,
          yearInstalled: plan.yearInstalled,
          loadingProfile: plan.loadingProfile,
          status: plan.status,
          loadings: plan.loadings
        }))
      }
    };
  }

  function exportAutoExistingJson(state = global.AppStore?.getState?.()) {
    const result = inspect(state || {});
    if (!result.canExport) return { ...result, ok: false };
    if (typeof global.ProjectExport?.downloadJson !== "function") {
      return { ...result, canExport: false, ok: false, reason: "JSON download is not available." };
    }
    const filePart = global.AppPersistence?.safeJobFilePart?.(state?.jobName || state?.importedFileName || "pole_job") || "pole_job";
    global.ProjectExport.downloadJson(`${filePart}_AutoExisting.json`, result.payload);
    return { ...result, ok: true };
  }

  global.AutoExistingExport = Object.freeze({
    CASES,
    exportAutoExistingJson,
    getYearInfo,
    inspect,
    loadingPrincipal,
    collectionPoleIds,
    parseYear,
    planForPole,
    sortedExistingPoles
  });
})(window);
