(function (global) {
  "use strict";

  function create(context = {}) {
    const { STATUS, S, parse, format, text, groupsForPole, idealProposedHeight } = context;

    function splitIssueMessage(message) {
      const source = text(message);
      if (!source || source === "OK") return [];
      return source.split(/\.\s+/).map(value => value.trim()).filter(Boolean).map(value => value.endsWith(".") ? value : `${value}.`);
    }

    function isMidspanIssue(message) {
      return /^(Environment:|Power MS:|Comm-comm MS|Midspan order|No room for proposed above top comm)/i.test(text(message));
    }

    function issueSeverity(message) {
      const values = (text(message).match(/[-−]?\d+\s*'(?:\s*\d+\s*")?|[-−]?\d+\s*"/g) || [])
        .map(value => parse(value.replace("−", "-")))
        .filter(value => value !== null);
      if (!values.length) return 1;
      if (/does not respect/i.test(message) && values.length >= 3) return Math.max(1, values[1] - Math.abs(values[0] - values[2]));
      if (/separation/i.test(message) && /minimum/i.test(message) && values.length >= 2) return Math.max(1, values.at(-1) - values.at(-2));
      if (/(exceeds|<|>)/.test(message) && values.length >= 2) return Math.max(1, Math.abs(values[0] - values[1]));
      return 1;
    }

    function physicalSpanPair(span) {
      return [span?.fromPole, span?.toPole].filter(Boolean).sort().join("|");
    }

    function collectIssues(poleId) {
      const state = S()?.getState?.() || {};
      const poleIssues = new Map();
      const midspanIssues = new Map();
      const add = (target, message) => {
        const clean = text(message);
        if (clean && clean !== "OK") target.set(clean.toLowerCase().replace(/\s+/g, " "), clean);
      };
      const pairs = new Set((S()?.getConnectedSpans?.(poleId) || []).map(physicalSpanPair).filter(Boolean));
      const spanIds = new Set(Object.values(state.spans || {}).filter(span => pairs.has(physicalSpanPair(span))).map(span => span.spanId));
      const rows = Object.values(state.spanComms || {}).filter(row => row.poleId === poleId || spanIds.has(row.spanId));
      rows.forEach(row => {
        if (row.flaggingStatus === "PROBLEM") {
          splitIssueMessage(row.flaggingMessage).forEach(message => {
            if (isMidspanIssue(message)) midspanIssues.set(message.toLowerCase().replace(/\s+/g, " "), message);
            else if (row.poleId === poleId) poleIssues.set(message.toLowerCase().replace(/\s+/g, " "), message);
          });
        }
        if (row.clearanceMSStatus === "PROBLEM" && !/Power MS:/i.test(row.flaggingMessage || "")) {
          add(midspanIssues, row.clearanceMSMessage || "Midspan clearance issue.");
        }
      });
      (S()?.getSpanSidesForPole?.(poleId) || []).forEach(side => {
        if (!side.proposedHOA) return;
        if (side.proposedFlaggingStatus === "PROBLEM") {
          splitIssueMessage(side.proposedFlaggingMessage).forEach(message => add(poleIssues, message));
        }
        if (side.clearanceMSStatus === "PROBLEM") add(midspanIssues, side.clearanceMSMessage || "Proposed midspan clearance issue.");
      });
      const poleMessages = Array.from(poleIssues.values());
      const midspanMessages = Array.from(midspanIssues.values());
      return {
        poleMessages,
        midspanMessages,
        poleViolationCount: poleMessages.length,
        poleViolationInches: poleMessages.reduce((sum, message) => sum + issueSeverity(message), 0),
        midspanViolationCount: midspanMessages.length,
        midspanViolationInches: midspanMessages.reduce((sum, message) => sum + issueSeverity(message), 0)
      };
    }

    function movementMetrics(poleId, mode) {
      const groups = groupsForPole(poleId);
      let movedCommCount = 0;
      let totalMovementInches = 0;
      groups.forEach(group => {
        const movement = Math.abs((group.effectiveInches ?? group.existingInches) - group.existingInches);
        if (movement) {
          movedCommCount += 1;
          totalMovementInches += movement;
        }
      });
      const ideal = idealProposedHeight(groups, mode);
      const proposals = (S()?.getSpanSidesForPole?.(poleId) || []).map(side => parse(side.proposedHOA || "")).filter(value => value !== null);
      return {
        movedCommCount,
        totalMovementInches,
        idealProposedInches: ideal,
        selectedProposedInches: proposals.length
          ? (mode === "LOW_COMM" ? Math.min(...proposals) : Math.max(...proposals))
          : null,
        proposedDistanceFromIdeal: ideal === null || !proposals.length
          ? Number.MAX_SAFE_INTEGER
          : proposals.reduce((sum, value) => sum + Math.abs(value - ideal), 0)
      };
    }

    function categoryForAnalysis(value) {
      if (!value.poleViolationCount && !value.midspanViolationCount) return 0;
      if (!value.poleViolationCount) return 1;
      return 2;
    }

    function rankAnalysis(value) {
      return [
        categoryForAnalysis(value),
        value.poleViolationCount,
        value.poleViolationInches,
        value.midspanViolationCount,
        value.midspanViolationInches,
        value.movedCommCount,
        value.totalMovementInches,
        value.proposedDistanceFromIdeal
      ];
    }

    function compareAnalyses(left, right) {
      if (!left) return 1;
      if (!right) return -1;
      const a = rankAnalysis(left);
      const b = rankAnalysis(right);
      for (let index = 0; index < a.length; index += 1) if (a[index] !== b[index]) return a[index] - b[index];
      return 0;
    }

    function compareTopRecoveryAnalyses(left, right) {
      if (!left) return 1;
      if (!right) return -1;
      const a = rankAnalysis(left);
      const b = rankAnalysis(right);
      for (let index = 0; index <= 4; index += 1) {
        if (a[index] !== b[index]) return a[index] - b[index];
      }
      const leftProposed = Number.isFinite(left.selectedProposedInches) ? left.selectedProposedInches : -1;
      const rightProposed = Number.isFinite(right.selectedProposedInches) ? right.selectedProposedInches : -1;
      if (leftProposed !== rightProposed) return rightProposed - leftProposed;
      for (let index = 5; index < a.length; index += 1) {
        if (a[index] !== b[index]) return a[index] - b[index];
      }
      return 0;
    }

    function statusForAnalysis(value) {
      return [STATUS.SAFE, STATUS.BEST_AVAILABLE, STATUS.CRITICAL][categoryForAnalysis(value)];
    }

    function resultMessage(value, status) {
      if (status === STATUS.SAFE) return "Pole and Midspan clearances satisfied.";
      if (status === STATUS.BEST_AVAILABLE) {
        return `Pole clearances satisfied. Midspan keeps ${value.midspanViolationCount} issue${value.midspanViolationCount === 1 ? "" : "s"} (${format(value.midspanViolationInches)} total shortfall).`;
      }
      if (status === STATUS.CRITICAL) {
        return `No pole-compliant aerial arrangement was found. Best candidate keeps ${value.poleViolationCount} pole issue${value.poleViolationCount === 1 ? "" : "s"} (${format(value.poleViolationInches)} total) and ${value.midspanViolationCount} Midspan issue${value.midspanViolationCount === 1 ? "" : "s"}.`;
      }
      return "Manual review required.";
    }

    function analyzeCurrentState(poleId, mode) {
      return { poleId, mode, ...collectIssues(poleId), ...movementMetrics(poleId, mode) };
    }

    return {
      splitIssueMessage,
      isMidspanIssue,
      issueSeverity,
      physicalSpanPair,
      collectIssues,
      movementMetrics,
      categoryForAnalysis,
      rankAnalysis,
      compareAnalyses,
      compareTopRecoveryAnalyses,
      statusForAnalysis,
      resultMessage,
      analyzeCurrentState
    };
  }

  global.AutoCalculateCandidateAnalysis = { create };
})(window);
