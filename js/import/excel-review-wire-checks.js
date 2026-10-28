(function (global) {
  "use strict";

  global.ExcelReviewWireChecks = {
    create(context) {
      const { add, I, isIntecProject, normalizedInsulator, normalizedText, pick, rowsForPole, text } = context;
      const ALLOWED_COMM_OWNERS = new Set([
        "communication",
        "3j communications",
        "cable one, prescott",
        "cable one, show low",
        "century link communications",
        "commnet",
        "cox communications",
        "cablevision of flagstaff",
        "eagle west llc",
        "mci metro",
        "wecom inc",
        "catv",
        "telco",
        "fiber",
        "cable one"
      ]);
      const COMM_INSULATORS = ["single bolt", "three bolt", "j-hook"];
      const TELCO_DROP_SIZES = ["0.25in Telco", "0.5in Telco", "0.75in Telco", "1in Telco"];
      const CATV_DROP_SIZES = ["0.25in CATV", "0.5in CATV", "0.75in CATV", "1in CATV"];
      const TELCO_BUNDLE_SIZES = [
        [0.5, 0.242], [1, 0.242], [1.5, 0.242], [2, 0.306],
        [2.5, 0.306], [3, 0.386], [3.5, 0.386]
      ];
      const PRIMARY_INSULATORS = ["deadend aps", "horizontal line post", "post 15\""];
      const SECONDARY_INSULATORS = ["pin 7.5in", "deadend aps", "spool 3in", "suspension aps"];
      const NEUTRAL_INSULATORS = [...SECONDARY_INSULATORS, "pin 8.38"];

      function isSelfSupportingFiber(value) {
        return /\bself[\s-]?supporting\s+fiber\b/i.test(text(value));
      }

      function addIntecWireChecks(result, poleId) {
        if (!isIntecProject()) return;
        rowsForPole("spanWires", poleId).forEach((row, index) => {
          const ownerRaw = text(pick(row, ["Owner", "owner"]));
          const owner = normalizedText(ownerRaw).replace(/^communication\s*>\s*/i, "").trim();
          const size = text(pick(row, ["Size", "Size.display", "Wire Size"]));
          const construction = text(pick(row, ["Construction"]));
          const insulator = normalizedInsulator(pick(row, ["Insulator"]));
          const descriptor = `${ownerRaw || "No owner"} / ${size || `row ${index + 2}`}`;
          const power = I().isPowerWire ? I().isPowerWire(row) : /^utility\s*>/i.test(ownerRaw);
          const communication = I().isCommunicationWire ? I().isCommunicationWire(row) : !power;

          if (normalizedText(construction).replace(/\s+/g, "") === "davit") {
            add(result, {
              phase: "HOA", section: "Span.Wire", code: "DAVIT_NOT_ALLOWED", status: "ERROR",
              title: "Construction", message: "Construction DAVIT is not allowed for INTEC.", expected: "Not DAVIT", actual: construction
            });
          }

          if (communication && !ALLOWED_COMM_OWNERS.has(owner)) {
            add(result, {
              phase: "HOA", section: "Span.Wire", code: "UNKNOWN_COMM_OWNER", status: "WARNING",
              title: "Communication Owner", message: `Communication owner is not in the INTEC list: ${ownerRaw || "empty"}.`,
              expected: "Allowed INTEC communication owner", actual: ownerRaw
            });
          }

          const selfSupportingFiber = communication && isSelfSupportingFiber(size);
          const validCommunicationInsulator = COMM_INSULATORS.includes(insulator)
            || (selfSupportingFiber && insulator === "suspension aps");
          if (communication && !validCommunicationInsulator) {
            add(result, {
              phase: "HOA", section: "Span.Wire", code: "INVALID_COMM_INSULATOR", status: "ERROR",
              title: "Communication Insulator", message: `Invalid communication insulator for ${descriptor}.`,
              expected: selfSupportingFiber
                ? `${COMM_INSULATORS.join(", ")}, or suspension aps for Self-Supporting Fiber`
                : COMM_INSULATORS.join(", "),
              actual: text(pick(row, ["Insulator"])) || "Empty"
            });
          }

          if (!power) return;
          const powerType = /primary/i.test(size) ? "PRIMARY" : /secondary/i.test(size) ? "SECONDARY" : /neutral/i.test(size) ? "NEUTRAL" : "";
          if (!powerType) return;
          if (normalizedText(ownerRaw) !== "utility > aps (joint use)") {
            add(result, {
              phase: "HOA", section: "Span.Wire", code: "INVALID_POWER_OWNER", status: "ERROR",
              title: `${powerType} Owner`, message: `${powerType} owner must be UTILITY > APS (Joint use).`,
              expected: "UTILITY > APS (Joint use)", actual: ownerRaw || "Empty"
            });
          }
          const validInsulator = powerType === "PRIMARY"
            ? (insulator.includes("pin 8.38") || PRIMARY_INSULATORS.includes(insulator))
            : powerType === "NEUTRAL"
              ? (insulator.includes("pin 8.38") || NEUTRAL_INSULATORS.includes(insulator))
              : SECONDARY_INSULATORS.includes(insulator);
          if (!validInsulator) {
            const expectedInsulators = powerType === "PRIMARY"
              ? "Pin 8.38, Deadend APS, Horizontal Line Post, or Post 15\""
              : powerType === "NEUTRAL"
                ? `${SECONDARY_INSULATORS.join(", ")}, or Pin 8.38`
                : SECONDARY_INSULATORS.join(", ");
            add(result, {
              phase: "HOA", section: "Span.Wire", code: `INVALID_${powerType}_INSULATOR`, status: "ERROR",
              title: `${powerType} Insulator`, message: `Invalid ${powerType.toLowerCase()} insulator for ${descriptor}.`,
              expected: expectedInsulators,
              actual: text(pick(row, ["Insulator"])) || "Empty"
            });
          }
        });
      }

      function communicationSizePolicy(ownerRaw) {
        const owner = normalizedText(ownerRaw).replace(/^communication\s*>\s*/, "").trim();
        if (/\b(?:century\s*link|centurylink|ctl|telco)\b/.test(owner)) return { label: "Century Link", category: "Telco" };
        if (/\bmci\s+metro\b/.test(owner)) return { label: "MCI Metro", category: "Telco" };
        if (/\bcable\s*one\b/.test(owner) && /\bprescott\b/.test(owner)) return { label: "Cable One, Prescott", category: "CATV" };
        if (/\bcable\s*one\b/.test(owner) && /\bshow\s*low\b/.test(owner)) return { label: "Cable One, Show Low", category: "CATV" };
        if (/\bcox(?:\s+communications?)?\b/.test(owner)) return { label: "Cox Communications", category: "CATV" };
        return null;
      }

      function normalizedCommunicationDropSize(value) {
        const compact = normalizedText(value).replace(/[\s"'′″]/g, "");
        const match = compact.match(/^(\d+(?:\.\d+)?)(?:inches|inch|in)(telco|catv)$/);
        if (!match) return null;
        const inches = Number(match[1]);
        const category = match[2] === "telco" ? "Telco" : "CATV";
        return `${inches}in ${category}`;
      }

      function telcoBundleKey(value) {
        const bundleLabel = text(value).split(">")
          .map(segment => segment.trim())
          .find(segment => /communication\s+bundle/i.test(segment));
        if (!bundleLabel) return "";
        const match = bundleLabel.match(/^\s*(\d+(?:\.\d+)?)\s*(?:(?:in(?:ches?)?)|")?\s*communication\s+bundle\s+msgr\s*:\s*(\d+(?:\.\d+)?)\s*(?:(?:in(?:ches?)?)|")?\s*$/i);
        return match ? `${Number(match[1])}|${Number(match[2])}` : "";
      }

      function addCommunicationSizeChecks(result, poleId) {
        const telcoBundleKeys = new Set(TELCO_BUNDLE_SIZES.map(([bundle, messenger]) => `${bundle}|${messenger}`));
        const expectedBundles = TELCO_BUNDLE_SIZES.map(([bundle, messenger]) => `${bundle}\" Communication Bundle Msgr:${messenger}\"`);

        rowsForPole("spanWires", poleId).forEach(row => {
          const ownerRaw = text(pick(row, ["Owner", "owner"]));
          const policy = communicationSizePolicy(ownerRaw);
          if (!policy) return;
          const size = text(pick(row, ["Size", "Size.display", "Wire Size"]));
          if (!size) return;

          if (policy.category === "Telco" && /communication\s+bundle/i.test(size)) {
            if (!telcoBundleKeys.has(telcoBundleKey(size))) {
              add(result, {
                phase: "HOA", section: "Span.Wire", code: "INVALID_TELCO_BUNDLE_SIZE", status: "ERROR",
                title: "Telco Bundle Size", message: `Invalid Telco bundle size for ${policy.label}.`,
                expected: expectedBundles.join("; "), actual: `${ownerRaw} / ${size}`
              });
            }
            return;
          }

          const normalizedSize = normalizedCommunicationDropSize(size);
          if (!normalizedSize) return;
          const allowedSizes = policy.category === "Telco" ? TELCO_DROP_SIZES : CATV_DROP_SIZES;
          if (!allowedSizes.includes(normalizedSize)) {
            add(result, {
              phase: "HOA", section: "Span.Wire", code: "COMM_SIZE_OWNER_MISMATCH", status: "ERROR",
              title: "Communication Drop Size",
              message: `${policy.label} requires a ${policy.category} drop size; ${size} does not match.`,
              expected: allowedSizes.join(", "), actual: `${ownerRaw} / ${size}`
            });
          }
        });
      }

      return { addIntecWireChecks, addCommunicationSizeChecks };
    }
  };
})(window);
