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
        borrowMidspanFromPhysicalSpan: true,
        proposeForeSpanWithoutMidspan: false,
        hideProposedOwner: false
      }
    },
    METRONET: {
      id: "METRONET",
      label: "Metronet",
      settings: {
        position: "LOW_COMM",
        mrCase: "UPPER",
        proposedOwner: "METRONET",
        borrowMidspanFromPhysicalSpan: false,
        proposeForeSpanWithoutMidspan: true,
        hideProposedOwner: true
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
    if (/metronet|proposed\s*mnt|\bmnt\b/.test(text)) return "METRONET";
    return "INTEC";
  }

  global.ProjectProfiles = {
    PROFILES,
    normalizeProfileId,
    getProfile,
    applyProfileSettings,
    detectProfile
  };
})(window);
