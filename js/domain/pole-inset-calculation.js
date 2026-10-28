(function (global) {
  "use strict";

  global.PoleInsetCalculationModule = {
    create(context) {
      const { effectivePowerMidspanInches, format, getPrimaryPowerCommsClearance, H, isForeSpan, isPrimaryPowerRow, samePhysicalSpan, S, spanDirectionForPole, spanRelationOf } = context;

      function getPoleInsetSpanOptions(poleId) {
        const sideBySpanId = new Map(S().getSpanSidesForPole(poleId).map(side => [side.spanId, side]));
        const selectedSpanId = String(S().getPole(poleId)?.poleInsetSpanId || "").trim();
        const proposedSides = S().getConnectedSpans(poleId)
          .map(span => ({ side: sideBySpanId.get(span.spanId) || { spanId: span.spanId, poleId }, span }))
          .filter(item => item.span && !item.side.isProposedExcluded)
          .sort((a, b) => {
            const aSelected = a.span.spanId === selectedSpanId ? 0 : 1;
            const bSelected = b.span.spanId === selectedSpanId ? 0 : 1;
            if (aSelected !== bSelected) return aSelected - bSelected;
            const aFore = isForeSpan(a.span) ? 0 : 1;
            const bFore = isForeSpan(b.span) ? 0 : 1;
            if (aFore !== bFore) return aFore - bFore;
            const aAdditional = a.side.isAdditionalProposed ? 1 : 0;
            const bAdditional = b.side.isAdditionalProposed ? 1 : 0;
            if (aAdditional !== bAdditional) return aAdditional - bAdditional;
            const aForward = a.span.fromPole === poleId ? 0 : 1;
            const bForward = b.span.fromPole === poleId ? 0 : 1;
            return aForward - bForward;
          });

        // Reciprocal Excel rows describe one connected pole pair; show one selectable option.
        const uniqueByConnection = new Map();
        proposedSides.forEach(item => {
          const otherPole = item.span.fromPole === poleId ? item.span.toPole : item.span.fromPole;
          const connectionKey = [poleId, otherPole].filter(Boolean).sort().join("||");
          if (!uniqueByConnection.has(connectionKey)) uniqueByConnection.set(connectionKey, item);
        });

        return Array.from(uniqueByConnection.values()).map(({ span }) => {
          const direction = spanDirectionForPole(span, poleId);
          const otherPole = span.fromPole === poleId ? span.toPole : span.fromPole;
          const endpoints = [poleId, otherPole].filter(Boolean).join(" → ");
          return {
            spanId: span.spanId,
            direction,
            type: spanRelationOf(span),
            label: [endpoints, direction || "No direction"].filter(Boolean).join(" · ")
          };
        });
      }

      function getPoleInsetDetails(poleId) {
        const pole = S().getPole(poleId);
        if (!pole) {
          return {
            maxHeightAtMidspan: "",
            selectedHeight: "",
            endDropToPole: "",
            endDropToNextPole: "",
            spanId: "",
            direction: "",
            spanOptions: []
          };
        }

        const spanOptions = getPoleInsetSpanOptions(poleId);
        const selectedSpanId = spanOptions.some(option => option.spanId === pole.poleInsetSpanId)
          ? pole.poleInsetSpanId
          : spanOptions[0]?.spanId || "";
        const selectedSide = selectedSpanId ? S().getSpanSide(selectedSpanId, poleId) : null;
        const selectedSpan = selectedSpanId ? S().getSpan(selectedSpanId) : null;
        const settings = S().getState().settings || {};
        const secondaryClearance = H().parseHeight(
          settings.polePowerCommsClearance || settings.clearanceToPower || "40\""
        ) ?? 40;
        const primaryClearance = getPrimaryPowerCommsClearance();
        const selectedSpanIds = new Set(
          [selectedSpan?.spanId, selectedSpan?.sourceSpanId]
            .map(value => String(value || "").trim())
            .filter(Boolean)
        );
        // Fore/Back records of one physical connection share this inset's power reference.
        if (selectedSpan) {
          S().getConnectedSpans(poleId)
            .filter(span => samePhysicalSpan(span, selectedSpan))
            .forEach(span => {
              if (span?.spanId) selectedSpanIds.add(span.spanId);
              if (span?.sourceSpanId) selectedSpanIds.add(span.sourceSpanId);
            });
        }
        const selectedPowerRows = Array.from(selectedSpanIds)
          .flatMap(spanId => S().getSpanPowerForSpan(spanId));
        const selectedPowerCandidates = selectedPowerRows
          .map(row => {
            const midspan = effectivePowerMidspanInches(row);
            if (midspan === null) return null;
            const clearance = isPrimaryPowerRow(row) ? primaryClearance : secondaryClearance;
            return clearance === null ? null : midspan - clearance;
          })
          .filter(value => value !== null);
        const selectedSpanRecords = Array.from(selectedSpanIds)
          .map(spanId => S().getSpan(spanId))
          .filter(Boolean);
        if (!selectedPowerCandidates.length) {
          selectedSpanRecords.forEach(span => {
            const lowPowerMidspan = H().parseHeight(span.sourceMidspanLowPower || span.midspanLowPower || "");
            if (lowPowerMidspan !== null) selectedPowerCandidates.push(lowPowerMidspan - secondaryClearance);
            const storedMax = H().parseHeight(span.midspanMaxCommHeight || "");
            if (storedMax !== null) selectedPowerCandidates.push(storedMax);
          });
        }

        const selectedHeight = H().parseHeight(pole.poleInsetHeight || "");
        const localProposed = H().parseHeight(selectedSide?.proposedHOA || "");
        let nextPoleProposed = H().parseHeight(selectedSide?.proposedHOAChange || "");
        if (nextPoleProposed === null && localProposed !== null) {
          const existingEndDrop = H().parseHeight(selectedSide?.endDrop || "");
          if (existingEndDrop !== null) nextPoleProposed = localProposed + existingEndDrop;
        }

        return {
          maxHeightAtMidspan: selectedPowerCandidates.length ? format(Math.min(...selectedPowerCandidates)) : "",
          selectedHeight: pole.poleInsetHeight || "",
          endDropToPole: selectedHeight !== null && localProposed !== null
            ? format(selectedHeight - localProposed)
            : "",
          endDropToNextPole: selectedHeight !== null && nextPoleProposed !== null
            ? format(selectedHeight - nextPoleProposed)
            : "",
          spanId: selectedSpan?.spanId || "",
          direction: selectedSpan ? spanDirectionForPole(selectedSpan, poleId) : "",
          spanOptions
        };
      }

      return { getPoleInsetSpanOptions, getPoleInsetDetails };
    }
  };
})(window);
