(function (global) {
  "use strict";

  function create({
    MAX_CANDIDATES,
    S,
    idealProposedHeight,
    gap,
    stackGap,
    commClearance,
    boltClearance
  }) {
    function minimumTopCommHeight(groups, state = S()?.getState?.()) {
      const ordered = [...groups].sort((a, b) => b.existingInches - a.existingInches);
      if (!ordered.length) return null;
      const hasRequiredFloor = ordered.some(group =>
        (group.locked && group.lockedInches !== null)
        || Number.isFinite(group.minimumInches)
      );
      if (!hasRequiredFloor) return null;
      const floors = ordered.map(group => {
        if (group.locked && group.lockedInches !== null) return group.lockedInches;
        return Number.isFinite(group.minimumInches) ? group.minimumInches : 0;
      });
      for (let index = floors.length - 2; index >= 0; index -= 1) {
        floors[index] = Math.max(
          floors[index],
          floors[index + 1] + gap(ordered[index], ordered[index + 1], state)
        );
      }
      return floors[0];
    }

    function addCandidate(set, value, maxPole) {
      if (!Number.isFinite(value)) return;
      const rounded = Math.round(value);
      if (rounded >= 0 && rounded <= maxPole) set.add(rounded);
    }

    function isLegalExistingBoltTarget(target, existingBoltPoints, state, movingServiceDrop = false) {
      if (movingServiceDrop) return true;
      const requiredClearance = boltClearance(state);
      return existingBoltPoints.every(point => {
        const existing = typeof point === "number" ? point : point?.height;
        if (!Number.isFinite(existing) || point?.serviceDrop) return true;
        const difference = Math.abs(target - existing);
        return difference === 0 || difference >= requiredClearance;
      });
    }

    function highestLegalTarget(ceiling, floor, existingBoltPoints, state, movingServiceDrop = false) {
      const roundedCeiling = Math.max(0, Math.floor(ceiling));
      const roundedFloor = Math.max(0, Math.ceil(floor));
      for (let target = roundedCeiling; target >= roundedFloor; target -= 1) {
        if (isLegalExistingBoltTarget(target, existingBoltPoints, state, movingServiceDrop)) return target;
      }
      for (let target = Math.min(roundedCeiling, roundedFloor - 1); target >= 0; target -= 1) {
        if (isLegalExistingBoltTarget(target, existingBoltPoints, state, movingServiceDrop)) return target;
      }
      return roundedCeiling;
    }

    function candidateHeights({ groups, maxPole, mode, currentProposed = [], manualReference = null, preferMaximum = false, state = S()?.getState?.() }) {
      if (!Number.isFinite(maxPole) || maxPole < 0) return [];
      if (manualReference !== null) return [Math.round(manualReference)];
      const values = new Set();
      const requiredValues = new Set();
      const progressive = [];
      const progressiveSeen = new Set();
      const ideal = idealProposedHeight(groups, mode, state);
      const comm = commClearance(state);
      const bolt = boltClearance(state);
      const anchors = [ideal, maxPole, maxPole - bolt, ...currentProposed];
      const addProgressive = value => {
        if (!Number.isFinite(value)) return;
        const rounded = Math.round(value);
        if (rounded < 0 || rounded > maxPole || progressiveSeen.has(rounded)) return;
        progressiveSeen.add(rounded);
        progressive.push(rounded);
      };
      const minimumTop = mode === "TOP_COMM" ? minimumTopCommHeight(groups, state) : null;
      const highestLegalProposed = mode === "TOP_COMM"
        ? highestLegalTarget(
          maxPole,
          0,
          groups
            .filter(group => Number.isFinite(group.existingInches))
            .map(group => ({ height: group.existingInches, serviceDrop: Boolean(group.serviceDrop) })),
          state
        )
        : null;
      if (minimumTop !== null) {
        const requiredProposed = Math.round(minimumTop + comm);
        anchors.push(requiredProposed);
        if (requiredProposed >= 0 && requiredProposed <= maxPole) requiredValues.add(requiredProposed);
      }
      groups.forEach(group => {
        const existing = group.existingInches;
        const current = group.effectiveInches ?? existing;
        anchors.push(
          existing,
          current,
          existing + comm,
          existing - comm,
          existing + bolt,
          existing - bolt,
          group.minimumInches,
          group.maximumInches,
          group.minimumInches === null ? null : group.minimumInches + comm,
          group.maximumInches === null ? null : group.maximumInches + comm
        );
      });
      anchors
        .filter(Number.isFinite)
        .forEach(value => [-1, 0, 1].forEach(offset => addCandidate(values, value + offset, maxPole)));
      if (ideal !== null) {
        for (let offset = -18; offset <= 18; offset += 1) addCandidate(values, ideal + offset, maxPole);
        const start = mode === "LOW_COMM" ? ideal - 48 : Math.min(ideal, maxPole) - 60;
        const end = mode === "LOW_COMM" ? ideal + 48 : maxPole;
        for (let value = start; value <= end; value += 6) addCandidate(values, value, maxPole);
      }
      for (let offset = 0; offset <= 12; offset += 1) addCandidate(values, maxPole - offset, maxPole);
      const preferred = ideal ?? maxPole;
      if (mode === "TOP_COMM") {
        if (preferMaximum) {
          addProgressive(maxPole);
          addProgressive(highestLegalProposed);
        }
        addProgressive(ideal);
        if (!preferMaximum && (ideal === null || ideal > maxPole)) {
          addProgressive(maxPole);
          addProgressive(highestLegalProposed);
        }
        Array.from(requiredValues).forEach(addProgressive);
        addProgressive(maxPole);
        addProgressive(highestLegalProposed);
      }
      const orderedCandidates = [
        ...progressive,
        ...Array.from(values)
          .filter(value => !progressiveSeen.has(value))
          .sort((a, b) => Math.abs(a - preferred) - Math.abs(b - preferred) || (mode === "LOW_COMM" ? a - b : b - a))
      ]
        .slice(0, MAX_CANDIDATES);
      Object.defineProperty(orderedCandidates, "progressiveCount", {
        value: Math.min(progressive.length, orderedCandidates.length),
        enumerable: false
      });
      return orderedCandidates;
    }

    function buildStackPlan(groups, proposedInches, mode, maxPole, state = S()?.getState?.(), options = {}) {
      const ordered = mode === "LOW_COMM"
        ? [...groups].sort((a, b) => a.existingInches - b.existingInches)
        : [...groups].sort((a, b) => b.existingInches - a.existingInches);
      const existingBoltPoints = ordered
        .filter(group => Number.isFinite(group.existingInches))
        .map(group => ({ height: group.existingInches, serviceDrop: Boolean(group.serviceDrop) }));
      const topCommFloors = mode === "TOP_COMM"
        ? ordered.map(group => {
          if (group.locked && group.lockedInches !== null) return group.lockedInches;
          return Number.isFinite(group.minimumInches) ? group.minimumInches : 0;
        })
        : [];
      for (let index = topCommFloors.length - 2; index >= 0; index -= 1) {
        topCommFloors[index] = Math.max(
          topCommFloors[index],
          topCommFloors[index + 1] + (options.preserveExistingGaps
            ? Math.max(stackGap(ordered[index], ordered[index + 1], state, options), Math.abs(ordered[index].existingInches - ordered[index + 1].existingInches))
            : stackGap(ordered[index], ordered[index + 1], state, options))
        );
      }
      const plan = [];
      ordered.forEach((group, index) => {
        const previous = plan[index - 1];
        let target;
        if (group.locked && group.lockedInches !== null) {
          target = group.lockedInches;
        } else if (mode === "LOW_COMM") {
          const minimum = previous
            ? previous.targetInches + stackGap(previous.group, group, state, options)
            : proposedInches + commClearance(state);
          let preferred = group.existingInches;
          if (Number.isFinite(group.minimumInches)) preferred = Math.max(preferred, group.minimumInches);
          if (Number.isFinite(group.maximumInches)) preferred = Math.min(preferred, group.maximumInches);
          target = Math.min(maxPole, Math.max(preferred, minimum));
        } else {
          const requiredGap = previous
            ? (options.preserveExistingGaps
              ? Math.max(stackGap(previous.group, group, state, options), Math.abs(previous.group.existingInches - group.existingInches))
              : stackGap(previous.group, group, state, options))
            : 0;
          const maximum = previous
            ? previous.targetInches - requiredGap
            : proposedInches - commClearance(state);
          const ceiling = Number.isFinite(group.maximumInches)
            ? Math.min(maximum, group.maximumInches, maxPole)
            : Math.min(maximum, maxPole);
          const floor = topCommFloors[index] || 0;
          let preferred = group.existingInches;
          if (options.followProposed || options.preferHighest) {
            target = highestLegalTarget(ceiling, floor, existingBoltPoints, state, group.serviceDrop);
          } else if (ceiling >= floor) {
            target = Math.max(floor, Math.min(preferred, ceiling));
          } else {
            // This Proposed candidate cannot fit the required stack. Keep the
            // candidate inside its upper envelope and let validation reject it.
            target = Math.max(0, ceiling);
          }
        }
        if (options.onlyLower && !group.locked && target > group.existingInches) {
          // TOP COMM pole-first recovery may lower automatic communications,
          // but it must never raise one merely to satisfy a stack floor.
          target = group.existingInches;
        }
        plan.push({ group, targetInches: Math.round(target) });
      });
      return plan;
    }

    return { candidateHeights, buildStackPlan };
  }

  global.AutoCalculateCandidateGeneration = { create };
})(window);
