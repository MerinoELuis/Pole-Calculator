(function (global) {
  "use strict";

  // ProjectProfiles keeps customer/project defaults outside of the calculator
  // logic. Add new projects here instead of scattering special cases across
  // app.js, calculations.js, or the importers.
  const PROFILES = {
    INTEC: {
      id: "INTEC",
      label: "INTEC",
      settings: {
        position: "TOP_COMM",
        mrCase: "LOWER",
        proposedOwner: "Wecom",
        primaryPowerCommsClearance: "43\"",
        // INTEC Back Span midspans are unusual, but when the workbook really
        // supplies one it must follow endpoint HOA movements and remain
        // visible for review instead of being treated as an empty reference.
        calculateBackspanMidspan: true,
        borrowMidspanFromPhysicalSpan: false,
        proposeForeSpanWithoutMidspan: false,
        allowLowPowerMidspanAdjustment: true,
        showServiceDrop: true,
        showResagServiceDrop: true,
        hideProposedOwner: false,
        environmentClearances: {},
        streetlightBracketCommClearance: "",
        streetlightDripLoopCommClearance: "",
        powerGuyCommClearance: "",
        streetlightGroundingRequired: false,
        mrTemplate: "INTEC"
      }
    },
    METRONET: {
      id: "METRONET",
      label: "Metronet",
      settings: {
        position: "LOW_COMM",
        mrCase: "UPPER",
        // Metronet uses a separate WI selector. MidAm is the currently
        // supported work issuer and is exported as the Proposed owner.
        metronetWI: "MIDAM",
        proposedOwner: "MidAm",
        primaryPowerCommsClearance: "43\"",
        // MidAm records the measured midspan on some Back Span rows. Its value
        // follows movements made at either physical endpoint of the connection.
        calculateBackspanMidspan: true,
        borrowMidspanFromPhysicalSpan: false,
        proposeForeSpanWithoutMidspan: true,
        allowLowPowerMidspanAdjustment: false,
        showServiceDrop: false,
        showResagServiceDrop: false,
        hideProposedOwner: true,
        polePowerCommsClearance: "40\"",
        clearanceToPower: "40\"",
        commClearance: "12\"",
        streetlightBracketCommClearance: "20\"",
        streetlightDripLoopCommClearance: "12\"",
        powerGuyCommClearance: "",
        streetlightGroundingRequired: true,
        environmentClearances: {
          NONE: "15'6\"",
          STREET: "15'6\"",
          HIGHWAY: "15'6\"",
          PEDESTRIAN: "9'6\"",
          PARALLEL_TO_STREET: "15'6\"",
          OBSTRUCTED_PARALLEL_TO_STREET: "15'6\"",
          UNLIKELY_PARALLEL_TO_STREET: "15'6\"",
          RESIDENTIAL_DRIVEWAY: "15'6\"",
          COMMERCIAL_DRIVEWAY: "15'6\"",
          PARKING_LOT: "15'6\"",
          ALLEY: "15'6\"",
          RAILROAD: "23'6\"",
          RURAL: "9'6\"",
          FARM: "15'6\"",
          WATER_WITHOUT_SAILBOATS: "14'",
          WATER_WITH_SAILBOATS: "Variable"
        },
        mrTemplate: "METRONET"
      }
    },
    CSU: {
      // CSU is a MetroNet WI, not a third top-level project. Keep this
      // profile key for Excel detection and settings, but let the UI expose
      // it through Project=Metronet / WI=CSU.
      id: "METRONET",
      label: "CSU",
      visible: false,
      settings: {
        // CSU is the Colorado Springs MetroNet work instruction. It follows
        // the MetroNet low-communication workflow, but has its own owner,
        // clearances, equipment rules, and ground-clearance table.
        position: "LOW_COMM",
        mrCase: "UPPER",
        metronetWI: "CSU",
        proposedOwner: "MNT",
        calculateBackspanMidspan: true,
        borrowMidspanFromPhysicalSpan: false,
        proposeForeSpanWithoutMidspan: true,
        allowLowPowerMidspanAdjustment: false,
        showServiceDrop: false,
        showResagServiceDrop: false,
        hideProposedOwner: true,
        polePowerCommsClearance: "52\"",
        clearanceToPower: "52\"",
        // The CSU WI does not provide a separate primary-to-comm value. Use
        // the stated low-power-to-top-comm value instead of inventing 43 in.
        primaryPowerCommsClearance: "52\"",
        commClearance: "12\"",
        boltClearance: "4\"",
        midspanCommCommClearance: "12\"",
        transformerCommClearance: "42\"",
        streetlightBracketCommClearance: "16\"",
        streetlightUngroundedBracketCommClearance: "52\"",
        streetlightDripLoopCommClearance: "24\"",
        powerGuyCommClearance: "",
        streetlightGroundingRequired: false,
        environmentClearances: {
          NONE: "15'6\"",
          STREET: "15'6\"",
          HIGHWAY: "15'6\"",
          PEDESTRIAN: "9'6\"",
          PARALLEL_TO_STREET: "15'6\"",
          OBSTRUCTED_PARALLEL_TO_STREET: "15'6\"",
          UNLIKELY_PARALLEL_TO_STREET: "15'6\"",
          RESIDENTIAL_DRIVEWAY: "15'6\"",
          COMMERCIAL_DRIVEWAY: "15'6\"",
          PARKING_LOT: "15'6\"",
          ALLEY: "15'6\"",
          RAILROAD: "23'6\"",
          RURAL: "15'6\"",
          FARM: "15'6\"",
          WATER_WITHOUT_SAILBOATS: "14'",
          WATER_WITH_SAILBOATS: "Variable"
        },
        // Reuse the MetroNet MR wording, while keeping CSU as the WI so Excel
        // review does not apply the MidAm-only checks.
        mrTemplate: "METRONET"
      }
    }
  };

  function normalizeProfileId(value) {
    const id = String(value || "").trim().toUpperCase();
    return PROFILES[id] ? id : "INTEC";
  }

  function getProfile(value) {
    return PROFILES[normalizeProfileId(value)];
  }

  function applyProfileSettings(settings = {}, profileId = "INTEC") {
    const profile = getProfile(profileId);
    return {
      ...settings,
      projectProfile: profile.id,
      ...(profile.settings || {})
    };
  }

  function detectProfile({ fileName = "", owners = [] } = {}) {
    const text = [fileName, ...owners].join(" ").toLowerCase();
    if (/\bcsu\b|colorado\s*springs|\bcocs\d*/.test(text)) return "CSU";
    if (/metronet|proposed\s*mnt|\bmnt\b|utility\s*>\s*midam|\bmidam\b/.test(text)) return "METRONET";
    return "INTEC";
  }

  /**
   * Public project-profile registry and detection helpers.
   * @namespace ProjectProfiles
   */
  global.ProjectProfiles = {
    PROFILES,
    normalizeProfileId,
    getProfile,
    applyProfileSettings,
    detectProfile
  };
})(window);
