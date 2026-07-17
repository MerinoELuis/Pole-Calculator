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
        calculateBackspanMidspan: false,
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
        proposedOwner: "MidAm",
        // MidAm records the measured midspan on some Back Span rows. Those
        // rows remain references in the UI, but their value follows movements
        // made at either physical endpoint of the connection.
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
        powerGuyCommClearance: "3\"",
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
