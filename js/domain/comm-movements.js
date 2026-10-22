(function (global) {
  "use strict";

  global.CalculationAutoMovementModule = {
    create(context) {
      const { commOwnerLabel, countsAsTopCommReference, findMidspanSourceComm, findRemoteComm, format, getConnectedSpans, getEffectiveCommHOA, getImportedMidspanInchesForComm, getMidspanCommCommClearance, getMidspanInchesForComm, getPoleBoltBoltClearance, getPoleCommCommClearance, getSettingPosition, H, hasMakeReadyFiberReference, hasOverlashMakeReadyReference, isReciprocalBackSpan, isTerminalBackSpan, normalizedHeightLabelForCalc, normalizeOwnerForMatch, ownerCanonicalOf, parseMidspanValue, recalculateAll, recalculateSpansForPole, S, spanRelationOf, updateExistingHOAChange } = context;

      function updatePoleCommNewHeight(poleId, owner, spanId, newHeight) {
        return updateExistingHOAChange(poleId, owner, spanId, newHeight);
      }

      function cloneStateForAutoCalc() {
        return JSON.parse(JSON.stringify(S().getState()));
      }

      function autoCalcGroupKey(sc) {
        return [
          normalizeOwnerForMatch(commOwnerLabel(sc) || sc.owner),
          normalizedHeightLabelForCalc(sc.existingHOA)
        ].join("|");
      }

      function autoCalcGroupsForPole(poleId) {
        const groups = new Map();
        S().getSpanCommsForPole(poleId).filter(countsAsTopCommReference).forEach(sc => {
          const key = autoCalcGroupKey(sc);
          if (!groups.has(key)) {
            groups.set(key, {
              key,
              owner: commOwnerLabel(sc) || sc.owner,
              ownerToken: normalizeOwnerForMatch(commOwnerLabel(sc) || sc.owner),
              existingInches: H().parseHeight(sc.existingHOA),
              rows: []
            });
          }
          groups.get(key).rows.push(sc);
        });
        return Array.from(groups.values())
          .filter(group => group.existingInches !== null)
          .sort((a, b) => b.existingInches - a.existingInches);
      }

      function autoCalcMaxPole(poleId) {
        const pole = S().getPole(poleId);
        return H().parseHeight(pole?.maxCommHeight || "");
      }

      function autoCalcGapBetweenGroups(upper, lower) {
        if (upper.ownerToken && upper.ownerToken === lower.ownerToken) return getPoleBoltBoltClearance();
        return getPoleCommCommClearance();
      }

      function autoCalcPoleHasCommProblems(groups) {
        return groups.some(group => group.rows.some(row =>
          row.flaggingStatus === "PROBLEM" || row.clearanceMSStatus === "PROBLEM"
        ));
      }

      /**
       * Reports whether one physical span contains enough midspan data to appear
       * in Proposed by Span. Power midspans count even when Span.Wire has no comm
       * rows; this is common on an otherwise valid INTEC Fore Span.
       */
      function spanHasRealMidspan(spanId) {
        const hasPowerMidspan = S().getSpanPowerForSpan(spanId)
          .some(row => H().parseHeight(row.midspan || "") !== null);
        const hasCommMidspan = S().getSpanCommsForSpan(spanId)
          .some(row => getImportedMidspanInchesForComm(row) !== null);
        const hasProposedMidspan = S().getSpanSidesForSpan(spanId)
          .some(side => parseMidspanValue(side.ocalcMS || side.proposedMidspan || side.msProposed || "") !== null);
        return hasPowerMidspan || hasCommMidspan || hasProposedMidspan;
      }

      /**
       * Reports whether a directed span may own a Proposed row at this pole.
       * An imported Fore Span directed from the current pole is automatic. The
       * reciprocal Back Span is also eligible when it is the terminal pole's
       * only outgoing connection, so the incoming span can be attached there.
       */
      function isSpanEligibleForProposed(span, poleId) {
        const terminalBackSpan = isTerminalBackSpan(span, poleId);
        if (isReciprocalBackSpan(span) && !terminalBackSpan) return false;
        const type = spanRelationOf(span).toLowerCase();
        if (/fore\s*span|forespan/.test(type) && span?.fromPole === poleId) return true;
        if (terminalBackSpan) return true;
        // An Overlash Make Ready row can live on the opposite endpoint of an
        // Other/Back relationship. Keep that physical connection available for
        // Proposed so the fiber is placed on the existing job-owner messenger.
        return span?.fromPole === poleId && hasOverlashMakeReadyReference(span);
      }

      function autoCalcProposedSpansForPole(poleId) {
        const seen = new Set();
        const allowNoMidspan = S().getState().settings?.proposeForeSpanWithoutMidspan === true;
        return S().getConnectedSpans(poleId)
          .filter(span => isSpanEligibleForProposed(span, poleId) || S().getSpanSide(span.spanId, poleId)?.isManualProposed)
          .filter(span => !S().getSpanSide(span.spanId, poleId)?.isProposedExcluded)
          .filter(span => !isReciprocalBackSpan(span) || isTerminalBackSpan(span, poleId))
          .filter(span => isTerminalBackSpan(span, poleId)
            || allowNoMidspan
            || spanHasRealMidspan(span.spanId)
            || S().getSpanSide(span.spanId, poleId)?.isManualProposed
            || hasMakeReadyFiberReference(span))
          .filter(span => !S().getSpanSide(span.spanId, poleId)?.isAdditionalProposed)
          .filter(span => {
            const key = `${span.fromPole || ""}->${span.toPole || ""}`;
            if (seen.has(key)) return false;
            seen.add(key);
            return true;
          });
      }

      function autoCalcSetProposed(poleId, spans, proposedInches) {
        const proposed = format(proposedInches);
        spans.forEach(span => {
          const side = S().getSpanSide(span.spanId, poleId) || S().upsertSpanSide({ spanId: span.spanId, poleId });
          S().upsertSpanSide({
            ...side,
            proposedHOA: proposed,
            clearanceFixReadyAt: 0,
            pendingMidspanFinal: "",
            msProposed: "",
            finalMidspan: "",
            clearanceMSStatus: "",
            clearanceMSMessage: "",
            clearanceMSReason: "",
            clearanceMSIssue: false
          });
        });
      }

      function autoCalcBuildStackPlan(groups, topCommTarget) {
        if (topCommTarget === null || !groups.length) return [];
        const plan = [];
        groups.forEach((group, index) => {
          const previous = plan[index - 1];
          const stackTarget = previous ? previous.targetInches - autoCalcGapBetweenGroups(previous.group, group) : topCommTarget;
          const midspanTarget = autoCalcGroupMidspanTarget(group);
          const targetInches = Math.round(Math.min(group.existingInches, stackTarget, midspanTarget ?? Number.POSITIVE_INFINITY));
          plan.push({ group, targetInches });
        });
        return plan;
      }

      function autoCalcPlanMovesComm(plan) {
        return plan.some(item => item.group.rows.some(row => {
          const existing = H().parseHeight(row.existingHOA);
          return existing !== null && existing !== item.targetInches;
        }));
      }

      function autoCalcGroupMidspanTarget(group) {
        // A comm movement changes the related midspan by half of the pole movement.
        // This target is calculated from imported midspan + remote movement only,
        // ignoring the current local auto movement. That lets the second auto pass
        // reduce or remove a local move when the other pole already solved part of
        // the clearance.
        let target = null;
        group.rows.forEach(row => {
          const span = S().getSpan(row.spanId);
          const localExisting = H().parseHeight(row.existingHOA);
          const midspanSource = findMidspanSourceComm(row) || row;
          const importedMidspan = getImportedMidspanInchesForComm(midspanSource);
          const remote = findRemoteComm(row.spanId, row.poleId, ownerCanonicalOf(row), row.wireId || "");
          const remoteExisting = H().parseHeight(remote?.existingHOA || "");
          const remoteCurrent = H().parseHeight(remote ? getEffectiveCommHOA(remote) : "");
          const maxMS = H().parseHeight(span?.midspanMaxCommHeight || "");
          if (localExisting === null || importedMidspan === null || maxMS === null) return;

          let requiredMS = maxMS;
          if (getSettingPosition() === "TOP_COMM" && span?.fromPole === row.poleId) {
            const currentMS = getMidspanInchesForComm(row);
            const spanMidspans = S().getSpanCommsForSpan(row.spanId)
              .filter(countsAsTopCommReference)
              .map(getMidspanInchesForComm)
              .filter(value => value !== null);
            const isTopCommAtMidspan = currentMS !== null && spanMidspans.length && currentMS === Math.max(...spanMidspans);
            if (isTopCommAtMidspan) requiredMS = Math.min(requiredMS, maxMS - getMidspanCommCommClearance());
          }

          const remoteAdjustment = remoteExisting !== null && remoteCurrent !== null ? (remoteCurrent - remoteExisting) / 2 : 0;
          const midspanWithoutLocalMove = importedMidspan + remoteAdjustment;
          if (midspanWithoutLocalMove <= requiredMS) return;
          const localTarget = localExisting + ((requiredMS - midspanWithoutLocalMove) * 2);
          target = target === null ? localTarget : Math.min(target, localTarget);
        });
        return target;
      }

      function autoCalcFinalProposedForPlan(proposedInches, plan, maxPole) {
        if (!plan.length) return proposedInches;
        const topCommAfterMove = plan[0].targetInches;
        return Math.min(proposedInches, maxPole, topCommAfterMove + getPoleCommCommClearance());
      }

      function autoCalcProposedCandidates(groups, maxPole, hasCommProblems, hasSpaceAbove, aboveProposed) {
        // Candidate order matters:
        // 1. If the existing stack is healthy and there is space above, try the
        //    cleanest attachment first: top comm + Pole · Comm-comm.
        // 2. Try the pole maximum so violation cases can use all available room.
        // 3. Try existing HOA heights exactly. Reusing the same bolt elevation is
        //    allowed only if the comm at that elevation is moved away.
        // 4. Try one bolt clearance below the maximum as a last simple fallback.
        const values = [];
        const seen = new Set();
        const add = value => {
          if (value === null || value === undefined || !Number.isFinite(value) || value < 0 || value > maxPole) return;
          const rounded = Math.round(value);
          if (seen.has(rounded)) return;
          seen.add(rounded);
          values.push(rounded);
        };

        if (hasSpaceAbove) add(aboveProposed);
        add(maxPole);
        groups
          .map(group => group.existingInches)
          .filter(value => value !== null && value <= maxPole)
          .sort((a, b) => b - a)
          .forEach(add);
        add(maxPole - getPoleBoltBoltClearance());
        return values;
      }

      function autoCalcAffectedPoleIds(poleId) {
        const ids = new Set([poleId]);
        S().getConnectedSpans(poleId).forEach(span => {
          if (span.fromPole) ids.add(span.fromPole);
          if (span.toPole) ids.add(span.toPole);
        });
        const wireIds = new Set(S().getSpanCommsForPole(poleId).map(row => row.wireId).filter(Boolean));
        Object.values(S().getState().spanComms || {}).forEach(row => {
          if (wireIds.has(row.wireId)) ids.add(row.poleId);
        });
        return Array.from(ids).filter(Boolean);
      }

      function autoCalcClearExistingMarks() {
        Object.values(S().getState().spanComms || {}).forEach(row => {
          if (row.autoCalcStatus === "AUTO" && !row.autoCalcMessage) return;
          if (!row.autoCalcStatus && !row.autoCalcMessage) return;
          S().upsertSpanComm({
            ...row,
            autoCalcStatus: row.autoCalcStatus === "AUTO" ? "AUTO" : "",
            autoCalcMessage: ""
          });
        });
      }

      function autoCalcBadCommKeys(state) {
        return new Set(Object.values(state?.spanComms || {})
          .filter(row => row.flaggingStatus === "PROBLEM" || row.clearanceMSStatus === "PROBLEM")
          .map(row => S().keyForSpanComm(row.spanId, row.poleId, row.owner, row.wireId || "")));
      }

      function autoCalcBadSpanSideKeys(state) {
        return new Set(Object.values(state?.spanSides || {})
          .filter(side => side.proposedHOA && (side.proposedFlaggingStatus === "PROBLEM" || side.clearanceMSStatus === "PROBLEM"))
          .map(side => S().keyForSpanSide(side.spanId, side.poleId)));
      }

      function autoCalcResultIsSafe(affectedPoleIds, beforeState = null, primaryPoleId = "") {
        const beforeBadComms = autoCalcBadCommKeys(beforeState);
        const beforeBadSides = autoCalcBadSpanSideKeys(beforeState);
        const affected = new Set(affectedPoleIds);
        const rows = Object.values(S().getState().spanComms || {}).filter(row => affected.has(row.poleId));
        const badComm = rows.find(row => {
          if (row.flaggingStatus !== "PROBLEM" && row.clearanceMSStatus !== "PROBLEM") return false;
          const key = S().keyForSpanComm(row.spanId, row.poleId, row.owner, row.wireId || "");
          return row.poleId === primaryPoleId || !beforeBadComms.has(key);
        });
        if (badComm) return { ok: false, reason: badComm.flaggingMessage || badComm.clearanceMSMessage || "Comm clearance issue." };

        const badProposed = Object.values(S().getState().spanSides || {}).find(side =>
          affected.has(side.poleId)
          && side.proposedHOA
          && (side.proposedFlaggingStatus === "PROBLEM" || side.clearanceMSStatus === "PROBLEM")
          && (side.poleId === primaryPoleId || !beforeBadSides.has(S().keyForSpanSide(side.spanId, side.poleId)))
        );
        if (badProposed) return { ok: false, reason: badProposed.proposedFlaggingMessage || badProposed.clearanceMSMessage || "Proposed clearance issue." };
        return { ok: true, reason: "" };
      }

      function autoCalcApplyPlan(plan) {
        plan.forEach(item => {
          const target = format(item.targetInches);
          item.group.rows.forEach(row => {
            const existing = H().parseHeight(row.existingHOA);
            S().upsertSpanComm({
              ...row,
              existingHOAChange: existing === item.targetInches ? "" : target,
              autoCalcStatus: existing === item.targetInches ? "" : "AUTO",
              autoCalcMessage: ""
            });
          });
        });
      }

      function autoCalcTryCandidate(beforePole, poleId, proposedInches, requireMoveBecauseOfProblems = false, hasManualChanges = false) {
        // Each candidate is tested against a clean copy of the pole state. That
        // lets the solver try a proposed height, calculate any required comm stack,
        // run the normal validations, and throw that attempt away if it fails.
        S().setState(beforePole);
        const groups = autoCalcGroupsForPole(poleId);
        const proposedSpans = autoCalcProposedSpansForPole(poleId);
        if (!proposedSpans.length) return { ok: false, reason: "No proposed spans." };

        autoCalcSetProposed(poleId, proposedSpans, proposedInches);
        const topExisting = groups.length ? Math.max(...groups.map(group => group.existingInches)) : null;
        const topCommTarget = groups.length ? proposedInches - getPoleCommCommClearance() : null;
        const hasAutoChanges = groups.some(group => group.rows.some(row => row.autoCalcStatus === "AUTO"));
        const mustMoveComms = groups.length && (requireMoveBecauseOfProblems || hasAutoChanges || (topExisting !== null && topExisting > topCommTarget));
        const plan = mustMoveComms ? autoCalcBuildStackPlan(groups, topCommTarget) : [];

        if (!plan || plan.some(item => item.targetInches < 0)) return { ok: false, reason: "No safe pole stack found." };
        const maxPole = autoCalcMaxPole(poleId);
        const finalProposed = maxPole === null ? proposedInches : autoCalcFinalProposedForPlan(proposedInches, plan, maxPole);
        // Do not overwrite user-entered HOA changes. A candidate can still pass if
        // it only fills Proposed and the existing manual comm moves already work.
        if (hasManualChanges && autoCalcPlanMovesComm(plan)) return { ok: false, reason: "Existing HOA Change already set." };
        autoCalcSetProposed(poleId, proposedSpans, finalProposed);
        autoCalcApplyPlan(plan);
        const affected = autoCalcAffectedPoleIds(poleId);
        affected.forEach(recalculateSpansForPole);
        const safe = autoCalcResultIsSafe(affected, beforePole, poleId);
        return safe.ok ? { ok: true, reason: "" } : safe;
      }

      function autoCalcMovementSignature() {
        const state = S().getState();
        const commMoves = Object.values(state.spanComms || {})
          .map(row => ({
            key: S().keyForSpanComm(row.spanId, row.poleId, row.owner, row.wireId || ""),
            existingHOAChange: row.existingHOAChange || "",
            autoCalcStatus: row.autoCalcStatus || ""
          }))
          .sort((a, b) => a.key.localeCompare(b.key));
        const proposedMoves = Object.values(state.spanSides || {})
          .map(side => ({
            key: S().keyForSpanSide(side.spanId, side.poleId),
            proposedHOA: side.proposedHOA || "",
            nextPoleProposed: side.proposedHOAChange || "",
            nextPoleProposedAuto: Boolean(side.nextPoleProposedAuto)
          }))
          .sort((a, b) => a.key.localeCompare(b.key));

        return JSON.stringify({ commMoves, proposedMoves });
      }

      /**
       * Repeatedly proposes safe Top Comm movements until state converges or a
       * repeated/max-pass safety stop is reached.
       * @returns {{applied:number, manual:number, skipped:number, passes:number, converged:boolean, stoppedByRepeat:boolean, maxPassesReached:boolean, disabled?:boolean}}
       */
      function autoCalculateMovements() {
        const summary = { applied: 0, manual: 0, skipped: 0, passes: 0, converged: false, stoppedByRepeat: false, maxPassesReached: false };
        if (getSettingPosition() !== "TOP_COMM") {
          return { ...summary, disabled: true };
        }
        // Old auto/manual labels were only helper metadata. Clear them every run so
        // the UI only shows real clearance flagging, not solver bookkeeping.
        autoCalcClearExistingMarks();
        const poleIds = Object.keys(S().getState().poles || {});
        const appliedPoles = new Set();
        const manualPoles = new Set();
        const skippedPoles = new Set();
        const seenSignatures = new Set();
        const maxPasses = Math.max(8, poleIds.length * 2 + 4);

        let previousSignature = autoCalcMovementSignature();
        seenSignatures.add(previousSignature);

        for (let pass = 0; pass < maxPasses; pass += 1) {
          recalculateAll();
          poleIds.forEach(poleId => {
            const beforePole = cloneStateForAutoCalc();
            const groups = autoCalcGroupsForPole(poleId);
            const proposedSpans = autoCalcProposedSpansForPole(poleId);
            const maxPole = autoCalcMaxPole(poleId);
            if (!proposedSpans.length) {
              skippedPoles.add(poleId);
              return;
            }
            if (maxPole === null) {
              manualPoles.add(poleId);
              return;
            }

            const hasManualChanges = groups.some(group => group.rows.some(row => (row.otherHOA || row.existingHOAChange) && row.autoCalcStatus !== "AUTO"));
            const hasCommProblems = autoCalcPoleHasCommProblems(groups);
            const topExisting = groups.length ? Math.max(...groups.map(group => group.existingInches)) : null;
            const aboveProposed = topExisting === null ? maxPole : topExisting + getPoleCommCommClearance();
            const hasSpaceAbove = aboveProposed <= maxPole;

            const candidates = autoCalcProposedCandidates(groups, maxPole, hasCommProblems, hasSpaceAbove, aboveProposed);
            const solved = candidates.some(candidate => autoCalcTryCandidate(beforePole, poleId, candidate, hasCommProblems, hasManualChanges).ok);
            if (solved) {
              appliedPoles.add(poleId);
              manualPoles.delete(poleId);
              return;
            }

            S().setState(beforePole);
            recalculateSpansForPole(poleId);
            manualPoles.add(poleId);
          });

          recalculateAll();
          summary.passes = pass + 1;
          const nextSignature = autoCalcMovementSignature();
          if (nextSignature === previousSignature) {
            summary.converged = true;
            break;
          }
          if (seenSignatures.has(nextSignature)) {
            summary.stoppedByRepeat = true;
            break;
          }
          seenSignatures.add(nextSignature);
          previousSignature = nextSignature;
        }

        recalculateAll();
        if (!summary.converged && !summary.stoppedByRepeat && summary.passes >= maxPasses) {
          summary.maxPassesReached = true;
        }
        summary.applied = appliedPoles.size;
        summary.manual = manualPoles.size;
        summary.skipped = skippedPoles.size;
        return summary;
      }

      return { updatePoleCommNewHeight, cloneStateForAutoCalc, autoCalcGroupKey, autoCalcGroupsForPole, autoCalcMaxPole, autoCalcGapBetweenGroups, autoCalcPoleHasCommProblems, spanHasRealMidspan, isSpanEligibleForProposed, autoCalcProposedSpansForPole, autoCalcSetProposed, autoCalcBuildStackPlan, autoCalcPlanMovesComm, autoCalcGroupMidspanTarget, autoCalcFinalProposedForPlan, autoCalcProposedCandidates, autoCalcAffectedPoleIds, autoCalcClearExistingMarks, autoCalcBadCommKeys, autoCalcBadSpanSideKeys, autoCalcResultIsSafe, autoCalcApplyPlan, autoCalcTryCandidate, autoCalcMovementSignature, autoCalculateMovements };
    }
  };
})(window);