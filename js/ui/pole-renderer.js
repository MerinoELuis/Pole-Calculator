(function (global) {
  "use strict";

  global.PoleWorkspaceRenderer = {
    create(context) {
      const { autoCalculateRunning, bindLocalActions, els, escapeHtml, H, markDirty, ownerCanonicalOf, ownerImportedOf, poleResolutionOf, recordUndoSnapshot, render, S, scrollToPole, selectPole, setPoleIndexOpen, SPAN_COLOR_CLASS_COUNT, spanRelationOf, toast, updatePoleIndexToggleVisibility, wireEditableEvents } = context;
      const excelReviewRenderer = global.ExcelReviewRenderer.create({
        els,
        escapeHtml,
        markDirty,
        recordUndoSnapshot,
        S,
        selectPole,
        toast
      });

      function spanLabel(span) {
        if (!span) return "";
        return `${span.fromPole} → ${span.toPole || "Unknown"}`;
      }

      function shortSpanLabel(span) {
        if (!span) return "";
        return `${span.fromPole || "?"} → ${span.toPole || "Unknown"}`;
      }

      function spanLengthDisplay(span) {
        if (!span) return "";
        return span.lengthDisplay || "";
      }

      function isBackspan(span) {
        return /back\s*span|backspan/i.test(spanRelationOf(span));
      }

      function poleLink(poleId) {
        if (!poleId) return "";
        return `<button class="link-button" type="button" data-scroll-to-pole="${escapeHtml(poleId)}">${escapeHtml(poleId)}</button>`;
      }

      function displayMidspan(sc) {
        return global.Calculations.displayMidspanForComm(sc);
      }

      function displayDecimalFeetInput(primary, fallback = "") {
        if (primary) return primary;
        const parsed = H.parseHeight(fallback);
        if (parsed === null) return fallback || "";
        return String(Number((parsed / 12).toFixed(3)));
      }

      function connectedSpansSorted(poleId) {
        return S.getConnectedSpans(poleId).sort((a, b) =>
          String(a.spanIndex || a.spanId).localeCompare(String(b.spanIndex || b.spanId), undefined, { numeric: true })
        );
      }

      function spanColorClass(poleId, spanId) {
        const spans = connectedSpansSorted(poleId);
        const index = Math.max(0, spans.findIndex(span => span.spanId === spanId));
        return `span-color-${index % SPAN_COLOR_CLASS_COUNT}`;
      }

      function spanRowClasses(poleId, spanId) {
        return `span-color-row ${spanColorClass(poleId, spanId)}`;
      }

      function spanChip(poleId, spanId) {
        return `<span class="span-color-chip ${spanColorClass(poleId, spanId)}" aria-hidden="true"></span>`;
      }

      function spanColorDot(poleId, spanId) {
        return `<span class="span-color-dot ${spanColorClass(poleId, spanId)}" aria-hidden="true"></span>`;
      }

      function commOwnerLabel(sc) {
        const label = global.Calculations.commOwnerLabel
          ? global.Calculations.commOwnerLabel(sc)
          : String(sc?.ownerDisplay || sc?.owner || ownerCanonicalOf(sc) || ownerImportedOf(sc) || "").trim();
        return label || "No owner";
      }

      function normalizedHeightLabel(value) {
        const parsed = H.parseHeight(value || "");
        return parsed === null ? String(value || "").trim() : H.formatHeight(parsed);
      }

      function commGroupKey(sc) {
        return [
          commOwnerLabel(sc).toLowerCase(),
          normalizedHeightLabel(sc?.existingHOA || "").toLowerCase()
        ].join("|");
      }

      function spanHasRealMidspan(spanId) {
        return global.Calculations.spanHasRealMidspan(spanId);
      }

      function isSpanEligibleForProposed(span, poleId) {
        if (global.Calculations.isSpanEligibleForProposed) {
          return global.Calculations.isSpanEligibleForProposed(span, poleId);
        }
        const type = spanRelationOf(span).toLowerCase();
        return /fore\s*span|forespan/.test(type) && span?.fromPole === poleId;
      }

      function proposedSpansForPole(poleId) {
        const seen = new Set();
        const allowNoMidspan = S.getState().settings?.proposeForeSpanWithoutMidspan === true;
        return connectedSpansSorted(poleId)
          .filter(span => isSpanEligibleForProposed(span, poleId) || S.getSpanSide(span.spanId, poleId)?.isManualProposed)
          // The terminal Back Span is the visible attachment endpoint for the
          // preceding pole's span. Keep it in the table even though it is also a
          // reciprocal Back Span and may not have imported Midspan data.
          .filter(span => !global.Calculations.isReciprocalBackSpan?.(span)
            || global.Calculations.isTerminalBackSpan?.(span, poleId))
          .filter(span => !S.getSpanSide(span.spanId, poleId)?.isProposedExcluded)
          .filter(span => global.Calculations.isTerminalBackSpan?.(span, poleId)
            || allowNoMidspan
            || span.ugActive === true
            || spanHasRealMidspan(span.spanId)
            || S.getSpanSide(span.spanId, poleId)?.isManualProposed
            || global.Calculations.hasMakeReadyFiberReference?.(span))
          .filter(span => {
            const side = S.getSpanSide(span.spanId, poleId);
            const key = side?.isAdditionalProposed
              ? `${span.fromPole || ""}->${span.toPole || ""}->${span.spanId}`
              : `${span.fromPole || ""}->${span.toPole || ""}`;
            if (seen.has(key)) return false;
            seen.add(key);
            return true;
          });
      }

      function groupedCommsForPole(poleId) {
        const groups = new Map();
        S.getSpanCommsForPole(poleId).forEach(sc => {
          const key = commGroupKey(sc);
          if (!groups.has(key)) {
            groups.set(key, {
              key,
              owner: commOwnerLabel(sc),
              existingHOA: normalizedHeightLabel(sc.existingHOA || ""),
              existingHOAChange: sc.existingHOAChange || "",
              otherHOA: sc.otherHOA || "",
              otherHOAActive: Boolean(sc.otherHOAActive),
              pofActive: Boolean(sc.pofActive),
              pofEligible: isSelfSupportingComm(sc),
              isPof: false,
              rows: []
            });
          }
          const group = groups.get(key);
          if (!group.existingHOAChange && sc.existingHOAChange) group.existingHOAChange = sc.existingHOAChange;
          if (!group.otherHOA && sc.otherHOA) group.otherHOA = sc.otherHOA;
          group.otherHOAActive = group.otherHOAActive || Boolean(sc.otherHOAActive);
          group.pofActive = group.pofActive || Boolean(sc.pofActive);
          group.pofEligible = group.pofEligible || isSelfSupportingComm(sc);
          if (global.Calculations.isPofComm && global.Calculations.isPofComm(sc)) group.isPof = true;
          group.rows.push(sc);
        });
        return Array.from(groups.values()).sort((a, b) => {
          // The table represents the original physical pole stack, so sorting stays
          // based on Existing HOA instead of HOA Change.
          const aHeight = H.parseHeight(a.existingHOA || "");
          const bHeight = H.parseHeight(b.existingHOA || "");
          if (aHeight !== null && bHeight !== null && aHeight !== bHeight) return bHeight - aHeight;
          if (aHeight !== null && bHeight === null) return -1;
          if (aHeight === null && bHeight !== null) return 1;
          return `${a.owner}${a.existingHOA}`.localeCompare(`${b.owner}${b.existingHOA}`, undefined, { numeric: true });
        });
      }

      function commMidspanEntries(group, poleId) {
        // One comm can touch more than one span. The Span, Max Height at MS,
        // Midspan and Other Pole HOA columns all come from this list so every
        // stacked value stays visually aligned row by row.
        const seen = new Set();
        const entries = [];
        const rows = [...group.rows].sort((a, b) =>
          `${a.spanId}${a.wireIndex || ""}`.localeCompare(`${b.spanId}${b.wireIndex || ""}`, undefined, { numeric: true })
        );
        rows.forEach(sc => {
          const span = S.getSpan(sc.spanId);
          const spanType = spanRelationOf(span).toLowerCase();
          const ownMidspan = H.parseHeight(sc.midspan || sc.ocalcMS || "") !== null;
          const isBackspan = /back\s*span|backspan/.test(spanType);
          const isOtherReference = /other/.test(spanType) && !ownMidspan;
          const calculatedBackspan = Boolean(global.Calculations.isCalculatedBackspanComm?.(sc));
          const isDerivedPhysicalMidspan = Boolean(global.Calculations.isDerivedPhysicalMidspanComm?.(sc));
          // A Back Span is REF only while it lacks its own measured midspan. When
          // INTEC/MidAm imports one on this exact row, it participates in endpoint
          // movement calculations and should no longer look like a passive REF.
          const isReferenceSpan = isDerivedPhysicalMidspan
            || (isBackspan && !calculatedBackspan)
            || isOtherReference;
          const midspan = displayMidspan(sc);
          const originalMidspan = sc.midspan || sc.ocalcMS || "";
          const originalMidspanTitle = originalMidspan
            ? `Original Midspan before movements: ${originalMidspan}`
            : "No original Midspan was imported";
          const hasMidspan = H.parseHeight(midspan || "") !== null;
          const hasStoredMidspan = H.parseHeight(sc.midspan || sc.ocalcMS || "") !== null;
          // Keep different wires on the same physical span as separate editable
          // rows. They may share an imported midspan today but still need
          // independent HOA changes and midspan calculations after a move.
          const key = `${sc.spanId}|${sc.wireId || ""}|${midspan || ""}|${hasMidspan ? "ms" : "ref"}`;
          if (seen.has(key)) return;
          seen.add(key);
          const remote = global.Calculations.findRemoteComm(sc.spanId, sc.poleId, ownerCanonicalOf(sc), sc.wireId || "");
          const midspanLocked = Boolean(sc.otherHOA || sc.existingHOAChange || remote?.otherHOA || remote?.existingHOAChange);
          // Backspans are reference-only here, but Fore/Other rows without an
          // imported midspan must stay editable so the user can create the MS.
          const canEditMidspan = !isBackspan && !isDerivedPhysicalMidspan && (!midspanLocked || !ownMidspan);
          const rawMidspan = H.parseHeight(sc.calculatedMidspan || sc.midspan || sc.ocalcMS || "");
          const resagTarget = H.parseHeight("15'6\"");
          const canResagServiceDrop = Boolean(
            !isReferenceSpan
            && sc.serviceDrop
            && (sc.resagServiceDrop || (rawMidspan !== null && resagTarget !== null && rawMidspan < resagTarget))
          );
          entries.push({
            otherHOAActive: Boolean(sc.otherHOAActive),
            spanHtml: `<div class="comm-span-row" data-span-id="${escapeHtml(sc.spanId)}">
              ${span ? spanColorDot(poleId, span.spanId) : ""}
              <span>${span ? `${poleLink(span.fromPole)} → ${poleLink(span.toPole)}` : escapeHtml(sc.spanId || "")}</span>
              ${isReferenceSpan ? `<em>REF</em>` : ""}
              <button class="inline-icon-action danger-action" type="button"
                data-delete-comm-span
                data-pole="${escapeHtml(sc.poleId)}"
                data-span="${escapeHtml(sc.spanId)}"
                data-owner="${escapeHtml(sc.owner)}"
                data-wire-id="${escapeHtml(sc.wireId || "")}"
                title="Delete only this span"
                aria-label="Delete only this span">&#10005;</button>
            </div>`,
            serviceDropHtml: `<div class="comm-midspan-value">
              <input type="checkbox"
                data-scope="spanComm"
                data-pole="${escapeHtml(sc.poleId)}"
                data-span="${escapeHtml(sc.spanId)}"
                data-owner="${escapeHtml(sc.owner)}"
                data-wire-id="${escapeHtml(sc.wireId || "")}"
                data-field="serviceDrop"
                ${sc.serviceDrop ? "checked" : ""}>
            </div>`,
            otherHoaHtml: `<div class="comm-midspan-value">
              <input class="input height-input remote-height-input" data-scope="spanComm" data-pole="${escapeHtml(sc.poleId)}" data-span="${escapeHtml(sc.spanId)}" data-owner="${escapeHtml(sc.owner)}" data-wire-id="${escapeHtml(sc.wireId || "")}" data-field="otherHOA" value="${escapeHtml(sc.otherHOA || "")}" placeholder="" title="Override HOA used for this comm movement" aria-label="Other HOA for ${escapeHtml(commOwnerLabel(sc))}">
            </div>`,
            hoaChangeHtml: `<div class="comm-midspan-value">
              <input class="input height-input remote-height-input" data-scope="spanComm" data-pole="${escapeHtml(sc.poleId)}" data-span="${escapeHtml(sc.spanId)}" data-owner="${escapeHtml(sc.owner)}" data-wire-id="${escapeHtml(sc.wireId || "")}" data-field="existingHOAChange" value="${escapeHtml(sc.existingHOAChange || "")}">
            </div>`,
            downGuyHtml: `<div class="comm-midspan-value">
              <input type="checkbox"
                data-scope="spanComm"
                data-pole="${escapeHtml(sc.poleId)}"
                data-span="${escapeHtml(sc.spanId)}"
                data-owner="${escapeHtml(sc.owner)}"
                data-wire-id="${escapeHtml(sc.wireId || "")}"
                data-field="downGuy"
                ${sc.downGuy ? "checked" : ""}>
            </div>`,
            resagHtml: `<div class="comm-midspan-value">
              <input type="checkbox"
                data-scope="spanComm"
                data-pole="${escapeHtml(sc.poleId)}"
                data-span="${escapeHtml(sc.spanId)}"
                data-owner="${escapeHtml(sc.owner)}"
                data-wire-id="${escapeHtml(sc.wireId || "")}"
                data-field="resagServiceDrop"
                title="Re-sag this service drop to 15'6&quot; at midspan"
                aria-label="Re-sag this service drop to 15 feet 6 inches at midspan"
                ${sc.resagServiceDrop ? "checked" : ""}
                ${canResagServiceDrop ? "" : "disabled"}>
            </div>`,
            midspanHtml: `<div class="comm-midspan-value colored-midspan ${span ? spanColorClass(poleId, span.spanId) : ""}" data-span-id="${escapeHtml(sc.spanId)}">${canEditMidspan
              ? `<input class="input height-input remote-height-input midspan-highlight-input" data-scope="spanComm" data-pole="${escapeHtml(sc.poleId)}" data-span="${escapeHtml(sc.spanId)}" data-owner="${escapeHtml(sc.owner)}" data-wire-id="${escapeHtml(sc.wireId || "")}" data-field="midspan" value="${escapeHtml(midspan)}" placeholder="" title="${escapeHtml(originalMidspanTitle)}">`
              : `<strong class="midspan-highlight-display" title="${escapeHtml(originalMidspanTitle)}">${escapeHtml(midspan)}</strong>`}
              ${hasStoredMidspan ? `<button class="inline-icon-action danger-action" type="button"
                data-clear-comm-midspan
                data-pole="${escapeHtml(sc.poleId)}"
                data-span="${escapeHtml(sc.spanId)}"
                data-owner="${escapeHtml(sc.owner)}"
                data-wire-id="${escapeHtml(sc.wireId || "")}"
                title="Clear only this midspan"
                aria-label="Clear only this midspan">&#10005;</button>` : ""}
              ${!hasStoredMidspan ? `<span class="inline-icon-action ghost-action" aria-hidden="true"></span>` : ""}
            </div>`,
            maxHeightAtMSHtml: `<div class="comm-midspan-value"><strong>${(!isReferenceSpan || calculatedBackspan) && hasMidspan ? escapeHtml(span?.midspanMaxCommHeight || "") : ""}</strong></div>`,
            remoteHtml: (isReferenceSpan && !calculatedBackspan) || !hasMidspan || !remote
              ? `<div class="comm-midspan-value"><strong></strong></div>`
              : `<div class="comm-midspan-value"><input class="input height-input remote-height-input" data-scope="spanComm" data-pole="${escapeHtml(remote.poleId)}" data-span="${escapeHtml(remote.spanId)}" data-owner="${escapeHtml(remote.owner)}" data-wire-id="${escapeHtml(remote.wireId || "")}" data-field="existingHOAChange" value="${escapeHtml(remote.existingHOAChange || remote.existingHOA || "")}"></div>`
          });
        });
        return entries;
      }

      function renderCommSpanRefs(group, poleId) {
        const entries = commMidspanEntries(group, poleId);
        return `<div class="comm-span-list">${entries.map(entry => entry.spanHtml).join("")}</div>`;
      }

      function renderCommServiceDropValues(group, poleId) {
        const entries = commMidspanEntries(group, poleId);
        return `<div class="comm-midspan-list">${entries.map(entry => entry.serviceDropHtml).join("")}</div>`;
      }

      function renderCommDownGuyValues(group, poleId) {
        const entries = commMidspanEntries(group, poleId);
        return `<div class="comm-midspan-list">${entries.map(entry => entry.downGuyHtml).join("")}</div>`;
      }

      function renderCommTransferValues(group, poleId) {
        // Transfer belongs to the physical comm at this pole, not to each span
        // touching it. Existing files may still have the flag on only one span,
        // so any checked relation makes the single grouped control active.
        const checked = group.rows.some(row => row.transferToNewPole);
        return `<div class="comm-group-checkbox">
          <input type="checkbox"
            data-scope="commGroup"
            data-pole="${escapeHtml(poleId)}"
            data-group-key="${escapeHtml(group.key)}"
            data-field="transferToNewPole"
            title="Replace raise/lower MR with one transfer to the new pole"
            aria-label="Transfer ${escapeHtml(group.owner)} to the new pole"
            ${checked ? "checked" : ""}>
        </div>`;
      }

      function renderCommResagValues(group, poleId) {
        const entries = commMidspanEntries(group, poleId);
        return `<div class="comm-midspan-list">${entries.map(entry => entry.resagHtml).join("")}</div>`;
      }

      function renderCommMidspanValues(group, poleId) {
        const entries = commMidspanEntries(group, poleId);
        return `<div class="comm-midspan-list comm-midspan-values" data-midspan-list>${entries.map(entry => entry.midspanHtml).join("")}</div>`;
      }

      function renderCommMaxHeightAtMSValues(group, poleId) {
        const entries = commMidspanEntries(group, poleId);
        return `<div class="comm-midspan-list">${entries.map(entry => entry.maxHeightAtMSHtml).join("")}</div>`;
      }

      function renderCommRemoteValues(group) {
        const entries = commMidspanEntries(group, group.rows[0]?.poleId || "");
        return `<div class="comm-midspan-list">${entries.map(entry => entry.remoteHtml).join("")}</div>`;
      }

      function renderCommFlagging(group) {
        const messages = Array.from(new Set(group.rows
          .map(row => String(row.flaggingMessage || "").trim())
          .filter(message => message && message !== "OK")));
        const hasProblem = group.rows.some(row => row.flaggingStatus === "PROBLEM");
        const hasMissing = group.rows.some(row => row.flaggingStatus === "MISSING" || row.flaggingStatus === "MISSING_POWER");
        if (!messages.length && !hasProblem && !hasMissing) return `<span class="badge changed">OK</span>`;
        const badgeClass = hasProblem ? "danger" : "warning";
        const label = hasProblem ? "Clearance Issue" : "Missing Data";
        return `<div class="flagging-cell">
          <span class="badge ${badgeClass}">${label}</span>
          <div class="flagging-message">${messages.map(escapeHtml).join(" ")}</div>
        </div>`;
      }

      function poleIdsForSpan(spanId) {
        const span = S.getSpan(spanId);
        return span ? [span.fromPole, span.toPole].filter(Boolean) : [];
      }

      function affectedPoleIdsForElement(el, scope) {
        const ids = new Set();
        if (el?.dataset?.pole) ids.add(el.dataset.pole);
        if (el?.dataset?.span) poleIdsForSpan(el.dataset.span).forEach(id => ids.add(id));
        if (scope === "pole" && el?.dataset?.pole) {
          S.getConnectedSpans(el.dataset.pole).forEach(span => {
            if (span.fromPole) ids.add(span.fromPole);
            if (span.toPole) ids.add(span.toPole);
          });
        }
        if (scope === "spanSide" && el?.dataset?.pole && el?.dataset?.field === "proposedHOA") {
          // A proposed value on the current pole can feed Next Pole Proposed for
          // spans arriving from previous poles, even when the spanId is different.
          S.getConnectedSpans(el.dataset.pole).forEach(span => {
            if (span.fromPole) ids.add(span.fromPole);
            if (span.toPole) ids.add(span.toPole);
          });
        }
        if (scope === "spanPower" && el?.dataset?.pole && ["attachmentHeight", "attachmentHeightChange"].includes(el?.dataset?.field)) {
          const selectedPower = S.getSpanPowerForPole(el.dataset.pole)
            .find(row => row.key === el.dataset.powerKey);
          const sourceHeight = H.parseHeight(selectedPower?.attachmentHeight || "");
          if (sourceHeight !== null) {
            S.getSpanPowerForPole(el.dataset.pole)
              .filter(row => H.parseHeight(row.attachmentHeight || "") === sourceHeight)
              .forEach(row => poleIdsForSpan(row.spanId).forEach(id => ids.add(id)));
          }
        }
        return Array.from(ids).filter(Boolean);
      }

      function bindScrollLinks(root) {
        if (!root) return;
        root.querySelectorAll("[data-scroll-to-pole]").forEach(btn => {
          btn.addEventListener("click", () => selectPole(btn.dataset.scrollToPole));
        });
      }

      function replacePoleCard(poleId) {
        if (!els.polesOverview || !poleId || !S.getPole(poleId)) return false;
        const oldCard = els.polesOverview.querySelector(`[data-pole-card="${CSS.escape(poleId)}"]`);
        if (!oldCard || !oldCard.isConnected || !oldCard.parentNode) return false;
        // Replacing a card after a checkbox change used to reset every wide table
        // to its left edge. Keep each table viewport so controls on the right stay
        // under the pointer while the recalculated card is rendered.
        const tableScrollPositions = Array.from(oldCard.querySelectorAll(".table-wrap"))
          .map(wrapper => ({ left: wrapper.scrollLeft, top: wrapper.scrollTop }));
        const powerSection = oldCard.querySelector(".power-section");
        const powerSectionOpen = powerSection ? powerSection.open : false;
        const equipmentSection = oldCard.querySelector(".equipment-section");
        const equipmentSectionOpen = equipmentSection ? equipmentSection.open : false;
        const template = document.createElement("template");
        template.innerHTML = renderPoleWorkspace(poleId).trim();
        const nextCard = template.content.firstElementChild;
        if (!nextCard) return false;
        const parent = oldCard.parentNode;
        if (!parent || !parent.contains(oldCard)) return false;
        parent.replaceChild(nextCard, oldCard);
        Array.from(nextCard.querySelectorAll(".table-wrap")).forEach((wrapper, index) => {
          const position = tableScrollPositions[index];
          if (!position) return;
          wrapper.scrollLeft = position.left;
          wrapper.scrollTop = position.top;
        });
        const nextPowerSection = nextCard.querySelector(".power-section");
        if (nextPowerSection) nextPowerSection.open = powerSectionOpen;
        const nextEquipmentSection = nextCard.querySelector(".equipment-section");
        if (nextEquipmentSection) nextEquipmentSection.open = equipmentSectionOpen;
        wireEditableEvents(nextCard);
        bindScrollLinks(nextCard);
        bindLocalActions(nextCard);
        return true;
      }

      function renderAffectedPoles(poleIds) {
        const ids = Array.from(new Set((poleIds || []).filter(Boolean)));
        if (!ids.length || (S.getState().ui.filter || "all") !== "all") {
          render();
          return;
        }

        global.Calculations.recalculateAll();
        renderSummary();
        renderPoleLists();

        let replacedAny = false;
        ids.forEach(id => {
          replacedAny = replacePoleCard(id) || replacedAny;
        });

        if (!replacedAny) renderAllPolesWorkspace();
        updateAutoCalculateButtonState();
      }

      function noteList(lines) {
        const clean = (Array.isArray(lines) ? lines : [lines])
          .map(line => String(line || "").trim())
          .filter(Boolean);
        if (!clean.length) return "";
        return `<div class="auto-notes">${clean.map(line => `<div>${escapeHtml(line)}</div>`).join("")}</div>`;
      }

      function commMidspanNote(row, span) {
        if (!span || !span.midspanLowPower) return "Missing Power MS to calculate Max MS Comm.";
        const midspan = typeof row === "string" ? row : displayMidspan(row);
        if (!midspan) return span.midspanMaxCommHeight ? `Missing midspan. Max ${span.midspanMaxCommHeight}.` : "Missing midspan.";
        return row?.clearanceMSMessage || `Max ${span.midspanMaxCommHeight} · Low Power MS ${span.midspanLowPower}.`;
      }

      function poleClearanceNote(height, pole, missingLabel = "Missing Data") {
        if (!pole || !pole.lowPower) return missingLabel;
        const max = pole.maxCommHeight || "";
        if (!height) return max ? `Missing data. Max ${max}.` : "Missing data.";
        const aboveMax = max && H.compareHeights(height, max) === 1;
        return aboveMax
          ? `Clearance issue. Max ${max} · Low Power ${pole.lowPower}.`
          : `OK. Max ${max} · Low Power ${pole.lowPower}.`;
      }

      function spanSideClearanceNote(side) {
        if (!side.msProposed && !side.finalMidspan) return "Missing O-CALC MS or automatic Proposed MS inputs.";
        const settings = S.getState().settings || {};
        if (side.clearanceMSReason === "LOW_POWER" && side.clearanceMSIssue && settings.allowLowPowerMidspanAdjustment !== false) return `Ensure min 30" to low power at midspan.`;
        if (side.clearanceMSStatus === "PENDING") {
          return `${side.clearanceMSMessage || "Clearance issue."} Adjusting to ${side.pendingMidspanFinal || ""}...`;
        }
        return side.clearanceMSMessage || "";
      }

      function renderMidspanClearanceStatus(row, span) {
        const status = typeof row === "string" ? "" : row?.clearanceMSStatus || "";
        const midspan = typeof row === "string" ? row : displayMidspan(row);
        if (!span || !span.midspanLowPower) return `<span class="badge warning">Missing Power MS</span>`;
        if (!midspan) return `<span class="badge warning">Missing Data</span>`;
        if (status === "PROBLEM") return `<span class="badge danger">Clearance Issue</span>`;
        if (status === "MISSING_POWER") return `<span class="badge warning">Missing Power MS</span>`;
        return `<span class="badge changed">OK</span>`;
      }

      function renderCombinedSpanFlagging(side) {
        const msStatus = String(side?.clearanceMSStatus || "").toUpperCase();
        const msIssue = msStatus === "PENDING" || msStatus === "PROBLEM";
        const proposedIssue = String(side?.proposedFlaggingStatus || "").toUpperCase() === "PROBLEM";
        const messages = [
          msIssue ? side.clearanceMSMessage : "",
          proposedIssue ? side.proposedFlaggingMessage : ""
        ].map(value => String(value || "").trim()).filter(Boolean);

        if (msIssue || proposedIssue) {
          return `<div class="flagging-cell">
            <span class="badge danger">Clearance Issue</span>
            ${messages.length ? `<div class="flagging-message">${escapeHtml(Array.from(new Set(messages)).join(" "))}</div>` : ""}
          </div>`;
        }

        const hasMS = Boolean(side?.msProposed || side?.finalMidspan);
        const hasProposed = Boolean(side?.proposedHOA);
        if (!hasMS || !hasProposed) return "";
        if (msStatus === "ADJUSTED" || msStatus === "ADJUSTMENT_NEEDED") {
          return `<span class="badge warning">Adjusted</span>`;
        }
        return `<span class="badge changed">OK</span>`;
      }

      function renderCommHoaChangeValues(group, poleId) {
        const entries = commMidspanEntries(group, poleId);
        if (!entries.some(entry => entry.otherHOAActive)) {
          return `<div class="comm-group-height-input"><input class="input height-input" data-scope="commGroup" data-pole="${escapeHtml(poleId)}" data-group-key="${escapeHtml(group.key)}" data-field="existingHOAChange" value="${escapeHtml(group.existingHOAChange || "")}"></div>`;
        }
        return `<div class="comm-midspan-list">${entries.map(entry => entry.otherHOAActive
          ? entry.hoaChangeHtml
          : `<div class="comm-midspan-value"><span class="muted">${escapeHtml(group.existingHOAChange || "")}</span></div>`).join("")}</div>`;
      }

      function renderCommOtherHoaValues(group, poleId) {
        const entries = commMidspanEntries(group, poleId);
        return `<div class="comm-midspan-list">${entries.map(entry => entry.otherHoaHtml).join("")}</div>`;
      }

      function renderClearanceStatus(height, pole, missingLabel = "Missing Data") {
        if (!pole || !pole.lowPower) return `<span class="badge warning">${missingLabel}</span>`;
        const max = pole.maxCommHeight || "";
        if (!height) return `<span class="badge warning">Missing Data</span>`;
        const aboveMax = max && H.compareHeights(height, max) === 1;
        if (aboveMax) return `<span class="badge danger">Clearance Issue</span>`;
        return `<span class="badge changed">OK</span>`;
      }

      function renderClearanceSettings() {
        if (!els.clearanceSettings) return;
        const settings = S.getState().settings || {};
        const profiles = global.ProjectProfiles?.PROFILES || {};
        const selectedProfile = settings.projectProfile || "INTEC";
        const projectOptions = Object.values(profiles)
          .filter(profile => profile.visible !== false)
          .filter((profile, index, all) => all.findIndex(item => item.id === profile.id) === index)
          .map(profile =>
          `<option value="${escapeHtml(profile.id)}" ${selectedProfile === profile.id ? "selected" : ""}>${escapeHtml(profile.label)}</option>`
        ).join("");
        const isComedWI = String(settings.metronetWI || "").toUpperCase() === "COMED";
        const poleClearanceRows = [
          ["polePowerCommsClearance", "Pole · Power-comms", settings.polePowerCommsClearance || settings.clearanceToPower || "40\""],
          ["commClearance", "Pole · Comm-comm", settings.commClearance || "12\""],
          ["boltClearance", "Pole · Bolt-bolt", settings.boltClearance || "4\""]
        ];
        const midspanClearanceRows = [
          ["midspanPowerCommClearance", "Midspan · Power-comm", settings.midspanPowerCommClearance || "30\""],
          ["midspanCommCommClearance", "Midspan · Comm-comm", settings.midspanCommCommClearance || "4\""]
        ];
        const streetlightClearanceRows = [];
        if (!isComedWI) {
          poleClearanceRows.splice(1, 0, ["primaryPowerCommsClearance", "Pole · Primary-comm", settings.primaryPowerCommsClearance || "43\""]);
          midspanClearanceRows.splice(1, 0, ["midspanPrimaryPowerCommClearance", "Midspan · Primary-comm", settings.midspanPrimaryPowerCommClearance || "33\""]);
        }
        if (selectedProfile === "INTEC") {
          streetlightClearanceRows.push(
            ["streetlightBracketCommClearance", "Streetlight bracket-comm", settings.streetlightBracketCommClearance || "40\""],
            ["streetlightGroundedBracketCommClearance", "Streetlight bracket-comm (Grounding)", settings.streetlightGroundedBracketCommClearance || "12\""]
          );
        }
        if ((selectedProfile === "METRONET" && String(settings.proposedOwner || "MidAm").toUpperCase() === "MIDAM")
          || (selectedProfile === "METRONET" && isComedWI)) {
          streetlightClearanceRows.push(
            ["streetlightBracketCommClearance", "Streetlight bracket-comm", settings.streetlightBracketCommClearance || (isComedWI ? '4"' : '20"')],
            ["streetlightDripLoopCommClearance", "Streetlight drip loop-comm", settings.streetlightDripLoopCommClearance || "12\""]
          );
        }
        const rawPosition = String(settings.position || "TOP_COMM").toUpperCase();
        const position = rawPosition === "LOW_COMM" ? "LOW_COMM"
          : rawPosition === "OVERLASH" ? "OVERLASH" : "TOP_COMM";
        const proposedOwner = settings.proposedOwner || "Wecom";
        const metronetWI = String(settings.metronetWI || (String(proposedOwner).toUpperCase() === "MNT" ? "CSU" : String(proposedOwner).toUpperCase() === "POWER" ? "COMED" : "MIDAM")).toUpperCase();
        const fiberSizes = settings.fiberSizes && typeof settings.fiberSizes === "object" ? settings.fiberSizes : {};
        const detectedFibers = new Set(Object.keys(fiberSizes));
        (S.getState().makeReadyReferences || []).forEach(ref => {
          const source = `${ref.attachmentFiber || ""} ${ref.attachmentSizeRaw || ""}`;
          const match = source.match(/\b(\d+)\s*CT\b/i);
          if (match) detectedFibers.add(`${match[1]}CT Fiber`);
        });
        const fiberTypes = Array.from(detectedFibers).sort((a, b) => {
          const aCount = Number((a.match(/\d+/) || [0])[0]);
          const bCount = Number((b.match(/\d+/) || [0])[0]);
          return aCount - bCount || a.localeCompare(b);
        });
        const profileSettings = global.ProjectProfiles?.getProfileSettings?.(selectedProfile) || {};
        const proposedOwners = selectedProfile === "OLSSON_OPPD"
          ? ["Cox", "Verizon", "Uniti"]
          : (Array.isArray(profileSettings.proposedOwners) && profileSettings.proposedOwners.length
            ? profileSettings.proposedOwners
            : ["Wecom", "MidAm", "CenturyLink", "Cable One", "Cox", "Fatbeam", "Vexus", "MCI Metro"]);
        const ownerOptions = proposedOwners.map(owner =>
          `<option value="${escapeHtml(owner)}" ${proposedOwner === owner ? "selected" : ""}>${escapeHtml(owner)}</option>`
        ).join("");
        const renderRow = ([field, label, value]) => `
            <label class="clearance-row">
              <span>${escapeHtml(label)}</span>
              <input class="input height-input" data-scope="settings" data-field="${escapeHtml(field)}" value="${escapeHtml(value)}" />
            </label>
          `;
        els.clearanceSettings.innerHTML = `
          <div class="settings-section">
            <h3>Clearance</h3>
            <div class="clearance-settings-groups">
              <div class="settings-grid clearance-grid">${poleClearanceRows.map(renderRow).join("")}</div>
              <div class="settings-grid clearance-grid">${midspanClearanceRows.map(renderRow).join("")}</div>
              ${streetlightClearanceRows.length ? `<div class="settings-grid clearance-grid">${streetlightClearanceRows.map(renderRow).join("")}</div>` : ""}
            </div>
          </div>
          <div class="settings-section settings-section-adjustments">
            <h3>Adjustments</h3>
            <div class="settings-grid adjustments-grid">
              <label class="clearance-row position-row">
                <span>Project</span>
                <select class="input position-select" data-scope="settings" data-field="projectProfile">
                  ${projectOptions}
                </select>
              </label>
              <label class="clearance-row position-row">
                <span class="position-label">Position</span>
                <select class="input position-select" data-scope="settings" data-field="position">
                  <option value="TOP_COMM" ${position === "TOP_COMM" ? "selected" : ""}>Top Comm</option>
                  <option value="LOW_COMM" ${position === "LOW_COMM" ? "selected" : ""}>Low Comm</option>
                  <option value="OVERLASH" ${position === "OVERLASH" ? "selected" : ""}>Overlash</option>
                </select>
              </label>
              <label class="clearance-row position-row">
                <span>Make Ready</span>
                <select class="input position-select" data-scope="settings" data-field="mrCase">
                  <option value="LOWER" ${(settings.mrCase || "LOWER") === "LOWER" ? "selected" : ""}>Lowercase</option>
                  <option value="UPPER" ${settings.mrCase === "UPPER" ? "selected" : ""}>Uppercase</option>
                </select>
              </label>
              ${selectedProfile === "METRONET" ? `<label class="clearance-row position-row">
                <span>WI</span>
                <select class="input position-select" data-scope="settings" data-field="metronetWI">
                  <option value="MIDAM" ${metronetWI === "MIDAM" ? "selected" : ""}>MidAm</option>
                  <option value="CSU" ${metronetWI === "CSU" ? "selected" : ""}>CSU</option>
                  <option value="COMED" ${metronetWI === "COMED" ? "selected" : ""}>ComEd</option>
                </select>
              </label>` : settings.hideProposedOwner ? "" : `<label class="clearance-row position-row">
                <span>Proposed Owner</span>
                <select class="input position-select" data-scope="settings" data-field="proposedOwner">
                  ${ownerOptions}
                </select>
              </label>`}
            </div>
          </div>
          ${selectedProfile === "INTEC" && fiberTypes.length ? `<div class="settings-section settings-section-fiber">
            <h3>Fiber</h3>
            <div class="settings-grid fiber-grid">
              <label class="clearance-row">
                <span>Messenger Size</span>
                <input class="input" data-scope="attachmentSettings" data-field="attachmentMessengerSize" value="${escapeHtml(settings.attachmentMessengerSize || "")}" />
              </label>
              ${fiberTypes.map(fiber => `<label class="clearance-row">
                <span>${escapeHtml(fiber)} Size</span>
                <input class="input" data-scope="attachmentSettings" data-field="fiberSize" data-fiber="${escapeHtml(fiber)}" value="${escapeHtml(fiberSizes[fiber] || "")}" />
              </label>`).join("")}
            </div>
          </div>` : ""}
        `;
        wireEditableEvents(els.clearanceSettings);
      }

      function updateAutoCalculateButtonState() {
        if (!els.autoCalculateBtn) {
          updateFollowProposedButtonState();
          return;
        }
        const state = S.getState();
        const hasPoleData = Object.keys(state.poles || {}).length > 0;
        const hasProposedFiber = global.CompactAutoProposed?.hasProposedFiber?.(state) === true;
        const supportedExportProject = global.CompactAutoProposed?.isExportProjectSupported?.(state) === true;
        const position = String(state.settings?.position || "TOP_COMM").toUpperCase();
        const isLowComm = position === "LOW_COMM";
        const isOverlash = position === "OVERLASH";
        const disableAuto = !hasPoleData || isLowComm || isOverlash || context.autoCalculateRunning;
        els.autoCalculateBtn.disabled = disableAuto;
        els.autoCalculateBtn.classList.toggle("btn-disabled", disableAuto);
        els.autoCalculateBtn.classList.toggle("btn-primary", !disableAuto);
        els.autoCalculateBtn.textContent = context.autoCalculateRunning
          ? "Auto Proposed · Processing"
          : !hasPoleData
          ? "Auto Proposed · Import Data First"
          : isLowComm
            ? "Auto Proposed · Top Comm Required"
            : isOverlash
              ? "Auto Proposed · Manual Overlash"
            : "Auto Proposed";

        if (els.exportProposedJsonBtn) {
          const canExportProposed = hasPoleData && hasProposedFiber && supportedExportProject;
          els.exportProposedJsonBtn.disabled = !canExportProposed;
          els.exportProposedJsonBtn.classList.toggle("btn-disabled", !canExportProposed);
          els.exportProposedJsonBtn.classList.toggle("btn-success", canExportProposed);
          els.exportProposedJsonBtn.title = !hasPoleData
            ? "Import pole data first."
            : !supportedExportProject
              ? "AutoProposed export is currently available for INTEC only."
            : !hasProposedFiber
              ? "Add at least one proposed fiber before exporting."
              : "Export the proposed fiber data.";
        }
        if (els.exportDebugJsonBtn) {
          const buildSource = String(global.AppDeploymentVersion?.source || "development").toLowerCase();
          const localOnly = buildSource !== "github-pages";
          els.exportDebugJsonBtn.disabled = !hasPoleData || !localOnly;
          els.exportDebugJsonBtn.classList.toggle("btn-disabled", !hasPoleData || !localOnly);
          els.exportDebugJsonBtn.classList.toggle("hidden", !localOnly);
        }
        updateFollowProposedButtonState();
      }

      function updateFollowProposedButtonState() {
        const button = els.followProposedBtn;
        if (!button) return;
        const state = S.getState();
        const hasPoleData = Object.keys(state.poles || {}).length > 0;
        const position = String(state.settings?.position || "TOP_COMM").toUpperCase();
        const active = state.settings?.autoMoveCommsToProposed === true
          || String(state.settings?.autoMoveCommsToProposed).toLowerCase() === "true";
        const disabled = !hasPoleData || position !== "TOP_COMM" || context.autoCalculateRunning;
        button.disabled = disabled;
        button.setAttribute("aria-pressed", String(active));
        button.classList.toggle("btn-disabled", disabled);
        button.classList.toggle("btn-primary", active && !disabled);
        button.textContent = active ? "Follow Proposed: ON" : "Follow Proposed: OFF";
        button.title = position !== "TOP_COMM"
          ? "Follow Proposed is available only in Top Comm mode."
          : "When enabled, changing Proposed recalculates existing comm movements.";
      }

      function renderPoleClassResults() {
        if (!els.poleClassResults) return;
        const rows = S.getState().poleClassChecks || [];
        renderPoleClassTabWarning();
        if (!rows.length) {
          els.poleClassResults.innerHTML = `<div class="detail-placeholder">Import an Excel file with Collection data to review Height / Class.</div>`;
          return;
        }
        const heightIssueCount = rows.filter(row => poleClassSeverity(row) === "critical").length;
        const classIssueCount = rows.filter(row => poleClassSeverity(row) === "warning").length;
        const issueCount = rows.filter(row => row.status !== "OK").length;
        els.poleClassResults.innerHTML = `
          <div class="pole-class-summary">
            <span class="badge changed">OK ${rows.length - issueCount}</span>
            <span class="badge danger">Height ${heightIssueCount}</span>
            <span class="badge warning">Class ${classIssueCount}</span>
          </div>
          <div class="table-wrap"><table class="pole-class-table">
            <thead><tr>
              <th>Pole ID</th><th>Birthmark</th><th>Tip</th><th>Imported Circ.</th><th>Diameter</th><th>Circumference</th><th>Imported Type</th><th>Height</th><th>Class</th><th>Expected Type</th><th>Status</th>
            </tr></thead>
            <tbody>${rows.map((row, index) => {
              const severity = poleClassSeverity(row);
              const birthmarkActive = row.birthmarkActive === undefined
                ? Boolean(row.birthmarkExists)
                : Boolean(row.birthmarkActive);
              const birthmarkLabel = birthmarkActive ? "active" : "inactive";
              return `
              <tr class="${poleClassRowClass(severity)}">
                <td><strong>${poleLink(row.poleId)}</strong></td>
                <td><input class="pole-class-birthmark-toggle" type="checkbox" data-scope="poleClassCheck" data-index="${index}" data-field="birthmarkActive" ${birthmarkActive ? "checked" : ""} aria-label="Birthmark ${birthmarkLabel}" title="${row.birthmarkExists ? `Birthmark imported; click to ${birthmarkActive ? "disable" : "enable"} its use in the class review.` : `No Birthmark imported; click to ${birthmarkActive ? "disable" : "enable"} manual Birthmark authority.`}"></td>
                <td>${escapeHtml(row.tip || "")}</td>
                <td>${escapeHtml(row.importedCircumference || "")}</td>
                <td><input class="input pole-class-diameter-input" data-scope="poleClassCheck" data-index="${index}" data-field="manualDiameter" value="${escapeHtml(row.manualDiameter || "")}" /></td>
                <td>${escapeHtml(row.circumference || "")}</td>
                <td>${escapeHtml(row.importedType || "")}</td>
                <td>${escapeHtml(row.calculatedHeight || "")}</td>
                <td>${escapeHtml(row.calculatedClass || "")}</td>
                <td>${escapeHtml(row.expectedType || "")}</td>
                <td>${renderPoleClassStatus(row, severity)}</td>
              </tr>
            `; }).join("")}</tbody>
          </table></div>`;
        els.poleClassResults.insertAdjacentHTML("beforeend", renderAnsiReferenceTable());
        wireEditableEvents(els.poleClassResults);
        bindScrollLinks(els.poleClassResults);
      }

      function poleClassSeverity(row) {
        const status = String(row?.status || "");
        if (status === "OK") return "ok";
        if (/Height mismatch|Missing Tip|No (?:ANSI|reference) height row/i.test(status)) return "critical";
        return "warning";
      }

      // Height discrepancies are operationally more important than class-only
      // differences. Keep this lookup independent from comm/proposed flagging so
      // the warning remains visible even when a pole is marked UG or PCO.
      function hasCriticalPoleHeightIssue(poleId) {
        const normalizedPoleId = String(poleId || "").trim().toUpperCase();
        return (S.getState().poleClassChecks || []).some(row => (
          String(row?.poleId || "").trim().toUpperCase() === normalizedPoleId
          && poleClassSeverity(row) === "critical"
        ));
      }

      function renderPoleClassTabWarning() {
        if (!els.poleClassTabWarning || !els.poleClassTabWarningCount) return;
        const criticalCount = (S.getState().poleClassChecks || [])
          .filter(row => poleClassSeverity(row) === "critical")
          .length;
        els.poleClassTabWarningCount.textContent = String(criticalCount);
        els.poleClassTabWarning.classList.toggle("hidden", criticalCount === 0);
        els.poleClassTabWarning.title = criticalCount === 1
          ? "1 pole has a critical height issue"
          : `${criticalCount} poles have critical height issues`;
      }

      function poleClassRowClass(severity) {
        if (severity === "ok") return "changed-row";
        if (severity === "critical") return "pole-class-critical-row";
        return "pole-class-warning-row";
      }

      function renderPoleClassStatus(row, severity) {
        if (severity === "ok") return `<span class="badge changed">OK</span>`;
        const label = severity === "critical" ? "Height Critical" : "Review Class";
        const badgeClass = severity === "critical" ? "danger" : "warning";
        return `<div class="flagging-cell">
          <span class="badge ${badgeClass}">${label}</span>
          <div class="flagging-message">${escapeHtml(row.status)}</div>
        </div>`;
      }

      function renderAnsiReferenceTable() {
        const classes = global.ExcelImport?.ANSI_POLE_CLASSES || [];
        const isOlsson = String(S.getState().settings?.projectProfile || "").toUpperCase() === "OLSSON_OPPD";
        const oppdTypes = isOlsson ? (global.ExcelImport?.OPPD_POLE_TYPES || []) : [];
        const table = global.ExcelImport?.ANSI_CLASS_TABLE || {};
        const groundline = global.ExcelImport?.ANSI_APPROX_GROUNDLINE_DISTANCE || {};
        const heights = Object.keys(table).map(Number).sort((a, b) => a - b);
        return `<details class="reference-table-panel" open>
          <summary>Pole Height and Class Reference Table</summary>
          ${oppdTypes.length ? `<p class="reference-note"><strong>OPPD permitted pole types:</strong> ${oppdTypes.map(escapeHtml).join(", ")}. Expected Type rounds to the available conservative class (for example 35-3 becomes 35-4).</p>` : ""}
          <div class="table-wrap"><table class="ansi-reference-table">
            <thead>
              <tr><th>Length of Pole (ft)</th><th>Approx. Groundline Distance from Butt (ft)</th>${classes.map(item => `<th>${escapeHtml(item)}</th>`).join("")}</tr>
            </thead>
            <tbody>${heights.map(height => `
              <tr>
                <td><strong>${height}</strong></td>
                <td>${escapeHtml(groundline[height] ?? "-")}</td>
                ${(table[height] || []).map(value => `<td>${value === null || value === undefined ? "-" : escapeHtml(value)}</td>`).join("")}
              </tr>
            `).join("")}</tbody>
          </table></div>
        </details>`;
      }

      function renderActiveView() {
        const activeView = S.getState().ui.activeView || "calculator";
        const isPoleClassView = activeView === "poleClass";
        document.querySelectorAll("[data-view-tab]").forEach(btn => {
          btn.classList.toggle("active", btn.dataset.viewTab === activeView);
        });
        document.querySelectorAll("[data-view-panel]").forEach(panel => {
          panel.classList.toggle("hidden", panel.dataset.viewPanel !== activeView);
        });
        document.body.classList.toggle("pole-class-view", isPoleClassView);
        document.body.classList.toggle("excel-review-view", activeView === "excelReview");
        if (activeView !== "calculator") setPoleIndexOpen(false);
        updatePoleIndexToggleVisibility();
      }

      function renderEnvironmentOptions(selected) {
        const current = selected || "NONE";
        return (S.ENVIRONMENT_OPTIONS || []).map(option => (
          `<option value="${escapeHtml(option.value)}" ${option.value === current ? "selected" : ""}>${escapeHtml(option.label)}</option>`
        )).join("");
      }

      function poleSummary(poleId) {
        const state = S.getState();
        const pole = S.getPole(poleId) || S.createPole({ poleId, isGenerated: /^Unknown-/i.test(poleId) });
        const spans = S.getConnectedSpans(poleId);
        const warnings = state.warnings.filter(w => w.poleId === poleId);
        const midspanCount = S.getSpanCommsForPole(poleId).filter(sc => sc.calculatedMidspan || sc.midspan || sc.ocalcMS).length;
        return {
          pole,
          spans,
          warnings,
          midspanCount,
          hasChanges: S.poleHasChanges(poleId)
        };
      }

      function isSelfSupportingComm(sc) {
        return /self[\s-]*supporting\s*fiber/i.test(`${sc?.size || ""} ${ownerImportedOf(sc)} ${sc?.ownerDisplay || sc?.owner || ""}`);
      }

      function poleHasProposed(poleId) {
        return S.getSpanSidesForPole(poleId).some(side => String(side?.proposedHOA || "").trim() !== "");
      }

      function poleFlaggingSummary(poleId) {
        const pole = S.getPole(poleId);
        const commIssues = S.getSpanCommsForPole(poleId)
          .filter(sc => sc.flaggingStatus === "PROBLEM")
          .length;
        const proposedIssues = S.getSpanSidesForPole(poleId)
          // The proposed table combines pole and midspan clearance in one
          // Clearance Issue badge. Count either source once in the pole index.
          .filter(side => side.proposedFlaggingStatus === "PROBLEM" || side.clearanceMSStatus === "PROBLEM")
          .length;
        const standaloneIssue = pole?.standaloneProposedHOA && global.Calculations.evaluateSpanSideFlagging({
          spanId: "",
          poleId,
          proposedHOA: pole.standaloneProposedHOA
        }).status === "PROBLEM" ? 1 : 0;
        const calculationIssueCount = commIssues + proposedIssues + standaloneIssue;
        const heightCritical = hasCriticalPoleHeightIssue(poleId);
        return {
          calculationIssueCount,
          heightCritical,
          issueCount: calculationIssueCount + (heightCritical ? 1 : 0),
          resolution: poleResolutionOf(pole) === "UG" ? "UG" : poleResolutionOf(pole) === "PCO" ? "PCO" : ""
        };
      }

      function renderSummary() {
        const state = S.getState();
        if (els.jobNameInput && document.activeElement !== els.jobNameInput) {
          els.jobNameInput.value = state.jobName || S.jobNameFromFileName(state.importedFileName);
        }
        if (els.projectMeta) {
          els.projectMeta.textContent = state.importedFileName
            ? `${state.importedFileName} · ${new Date(state.importedAt || Date.now()).toLocaleString()}`
            : "No file imported";
        }
      }

      function filteredPoles({ includeHidden = true, includeUnknown = true } = {}) {
        const state = S.getState();
        const search = (state.ui.search || "").toLowerCase();
        const filter = state.ui.filter || "all";
        const hidden = new Set(state.ui.hiddenPoleIds || []);
        return Object.keys(state.poles).filter(poleId => {
          // Generated endpoint placeholders (and imported poles whose identifier
          // contains "Unknown") are useful for span geometry, but they should not
          // clutter either user-facing pole index. Keep them in state so the
          // calculations remain intact and only filter them at the index render.
          // Keep real imported IDs such as P6-Unknown addressable. Only generated
          // endpoint placeholders beginning with Unknown- stay out of the index.
          if (!includeUnknown && /^unknown(?:-|\b)/i.test(String(poleId))) return false;
          if (!includeHidden && hidden.has(poleId)) return false;
          const summary = poleSummary(poleId);
          const pole = summary.pole;
          const comms = Array.isArray(pole.comms) ? pole.comms : [];
          const flagging = poleFlaggingSummary(poleId);
          const hasComms = comms.length > 0 || S.getSpanCommsForPole(poleId).length > 0;
          const text = `${poleId} ${pole.sequence || ""} ${pole.poleHeight || ""} ${pole.lowPower || ""} ${pole.maxCommHeight || ""} ${comms.map(c => `${c.owner || ""} ${c.existingHOA || ""}`).join(" ")}`.toLowerCase();
          if (search && !text.includes(search)) return false;
          if (filter === "warnings" && summary.warnings.length === 0 && flagging.issueCount === 0) return false;
          if (filter === "changed" && !summary.hasChanges) return false;
          if (filter === "ug" && poleResolutionOf(pole) !== "UG") return false;
          if (filter === "pco" && poleResolutionOf(pole) !== "PCO") return false;
          if (filter === "height" && !flagging.heightCritical) return false;
          if (filter === "ready" && flagging.issueCount > 0) return false;
          if (filter === "comms" && !hasComms) return false;
          return true;
        }).sort((a, b) => a.localeCompare(b, undefined, { numeric: true }));
      }

      function renderPoleListItem(poleId) {
        const state = S.getState();
        const { pole, warnings, hasChanges } = poleSummary(poleId);
        const hasProposed = poleHasProposed(poleId);
        const flagging = poleFlaggingSummary(poleId);
        const hasMR = (state.mr || []).some(item => item?.poleId === poleId && String(item.text || "").trim());
        const hasExistingViolations = Boolean(global.Calculations?.getMidamPreMoveViolations?.(poleId)?.length);
        const hidden = (state.ui.hiddenPoleIds || []).includes(poleId);
        const active = state.selectedPoleId === poleId ? " active" : "";
        return `<button class="pole-index-link${active}${hidden ? " hidden-pole" : ""}" data-pole-select="${escapeHtml(poleId)}" type="button" title="${hidden ? "Show this pole" : "Open this pole"}">
          <span>${escapeHtml(poleId)}</span>
          ${hidden ? `<span class="mini-dot hidden-status">Hidden</span>` : ""}
          ${pole.isGenerated ? `<span class="mini-dot warning">Gen</span>` : ""}
          ${flagging.resolution ? `<span class="mini-dot ${flagging.resolution.toLowerCase()}">${flagging.resolution}</span>` : ""}
          ${!flagging.resolution && flagging.calculationIssueCount ? `<span class="mini-dot danger">Flag ${flagging.calculationIssueCount}</span>` : ""}
          ${flagging.heightCritical ? `<span class="mini-dot danger" title="Critical pole height issue">&#9888; Height</span>` : ""}
          ${hasExistingViolations ? `<span class="mini-dot ev" title="Existing violations">EV</span>` : ""}
          ${hasMR ? `<span class="mini-dot warning" title="Make Ready exists for this pole">MR</span>` : ""}
          ${hasProposed ? `<span class="mini-dot changed">Proposed</span>` : ""}
        </button>`;
      }

      function renderPoleLists() {
        // Keep generated Unknown endpoints addressable in the index. They remain
        // hidden from the workspace through hiddenPoleIds and carry a Hidden badge.
        const poleIds = filteredPoles({ includeUnknown: true });
        els.poleSearchInput.value = S.getState().ui.search || "";
        els.warningFilterSelect.value = S.getState().ui.filter || "all";
        els.poleSearchInputTop.value = S.getState().ui.search || "";
        els.warningFilterSelectTop.value = S.getState().ui.filter || "all";
        [els.polesList, els.polesListTop].forEach(list => {
          list.innerHTML = poleIds.map(id => renderPoleListItem(id)).join("") || `<div class="detail-placeholder">No poles match that filter.</div>`;
          list.querySelectorAll("[data-pole-select]").forEach(btn => {
            btn.addEventListener("click", () => selectPole(btn.dataset.poleSelect));
          });
        });
      }

      function renderMRText(poleId) {
        const item = S.getState().mr.find(mr => mr.poleId === poleId);
        if (!item) return `<p class="muted">No Make Ready generated yet.</p>`;
        return `<pre class="mr-output">${escapeHtml(item.text)}</pre>`;
      }

      function formatExistingViolationsText(violations) {
        return violations.map(item => {
          const scope = item.scope === "pole" ? "ON POLE" : "MIDSPAN";
          const message = String(item.message || "")
            .trim()
            .replace(/^on[- ]pole\s+violation:\s*/i, "")
            .replace(/^midspan\s+violation:\s*/i, "")
            .replace(/;\s*separation\s+.*$/i, "")
            .replace(/;\s*environment\s+clearance\s+.*$/i, "")
            .toUpperCase();
          return `${scope} VIOLATION: ${message}`;
        }).join("\n");
      }

      function renderMidamPreMoveViolations(poleId) {
        const settings = S.getState().settings || {};
        const isMidAm = String(settings.projectProfile || "").toUpperCase() === "METRONET"
          && String(settings.proposedOwner || "MidAm").toUpperCase() === "MIDAM";
        if (!isMidAm) return "";
        const violations = global.Calculations?.getMidamPreMoveViolations?.(poleId) || [];
        if (!violations.length) return "";
        const text = formatExistingViolationsText(violations);
        return `<section class="subsection wide midam-violations-section" id="existing-violations-${escapeHtml(poleId)}">
          <div class="subsection-title-row">
            <h4>Existing Violations</h4>
            <button class="mini-btn" type="button" data-copy-existing-violations data-pole="${escapeHtml(poleId)}">Copy</button>
          </div>
          <pre class="existing-violations-output">${escapeHtml(text)}</pre>
        </section>`;
      }

      function renderPoleActions(poleId) {
        const pole = S.getPole(poleId);
        const resolution = poleResolutionOf(pole);
        const projectProfile = String(S.getState().settings?.projectProfile || "INTEC").toUpperCase();
        const isIntec = projectProfile === "INTEC";
        const isMetronet = projectProfile === "METRONET";
        const isOlsson = projectProfile === "OLSSON_OPPD";
        const powerEquipment = Array.isArray(pole?.metadata?.powerEquipment) ? pole.metadata.powerEquipment : [];
        const riserEquipment = powerEquipment.filter(row => String(row.category || row.type || "").toUpperCase().includes("RISER"));
        const secondaryDripLoopActive = Boolean(pole?.secondaryDripLoopActive);
        const showUGReason = resolution === "UG" && (isIntec || isMetronet || isOlsson);
        const ugTemplate = showUGReason
          ? global.MRLogic.getEditableUGTemplate(pole)
          : "";
        const ugSpanOptions = isIntec && resolution === "UG"
          ? (global.MRLogic.getUGSpanOptions?.(poleId) || [])
          : [];
        const savedUGSpanId = String(pole?.ugSpanId || "ALL").trim() || "ALL";
        const selectedUGSpanId = savedUGSpanId.toUpperCase() === "ALL"
          || ugSpanOptions.some(option => option.spanId === savedUGSpanId)
          ? savedUGSpanId
          : "ALL";
        const ugSpanTemplate = isIntec && selectedUGSpanId.toUpperCase() !== "ALL"
          ? (global.MRLogic.getEditableUGSpanTemplate?.(pole) || "")
          : "";
        const showPCOEditor = resolution === "PCO";
        const pcoTemplate = showPCOEditor
          ? global.MRLogic.getEditablePCOTemplate(pole)
          : "";
        const makeReadyText = S.getState().mr.find(item => item.poleId === poleId)?.text || "";
        const riserAvailable = (isIntec || isMetronet || isOlsson) && Boolean(global.MRLogic.isRiserAvailable?.(poleId));
        const riserEnabled = riserAvailable && Boolean(global.MRLogic.isRiserEnabled?.(poleId));
        // Pole Inset and UG/PCO are mutually exclusive actions.  Use the
        // effective state here as a guard as well as in the toggle handlers so
        // stale/imported state cannot leave the Pole Inset controls visible after
        // the pole is moved to UG or PCO.
        const poleInsetEnabled = !isMetronet && !isOlsson
          && resolution === "AERIAL"
          && Boolean(pole?.poleInsetActive);
        const poleInsetReason = String(pole?.poleInsetReason || "FAILING_CLEARANCES").toUpperCase() === "OVERLOADED"
          ? "OVERLOADED"
          : "FAILING_CLEARANCES";
        const poleInsetDetails = poleInsetEnabled
          ? (global.Calculations.getPoleInsetDetails?.(poleId) || {})
          : {};
        const showRiserDirection = (isIntec || isMetronet || isOlsson) && (riserEnabled || /\bPl riser\b/i.test(makeReadyText));
        const riserDirection = String(
          global.MRLogic.getResolvedRiserDirection?.(poleId)
            || pole?.ugRiserDirection
            || global.MRLogic.getImportedRiserDirection?.(poleId)
            || ""
        ).toUpperCase();
        const compassDirections = ["", "N", "NE", "E", "SE", "S", "SW", "W", "NW"];
        const directionOptions = compassDirections
          .map(direction => `<option value="${direction}" ${direction === riserDirection ? "selected" : ""}>${direction || "Select direction"}</option>`)
          .join("");
        const ancAvailable = (isIntec || isMetronet) && resolution === "AERIAL";
        const ohgAvailable = (isMetronet || isOlsson) && resolution === "AERIAL";
        const ancEnabled = ancAvailable && Boolean(global.MRLogic.isANCEnabled?.(poleId));
        const ohgEnabled = ohgAvailable && Boolean(pole?.ohgActive);
        const ancDistance = global.MRLogic.getEffectiveANCDistance?.(poleId)
          || pole?.ancDistance
          || "";
        const ancDGHOA = global.MRLogic.getEffectiveANCDGHOA?.(poleId)
          || pole?.dgHoa
          || "";
        const anchorDirection = String(global.MRLogic.getEffectiveANCDirection?.(poleId)
          || pole?.ancDirection
          || "").toUpperCase();
        const ohgDirection = String(pole?.ohgDirection || "").toUpperCase();
        const anchorDirectionOptions = compassDirections
          .map(direction => `<option value="${direction}" ${direction === anchorDirection ? "selected" : ""}>${direction || "Select direction"}</option>`)
          .join("");
        const ohgDirectionOptions = ["", "N", "E", "S", "W"]
          .map(direction => `<option value="${direction}" ${direction === ohgDirection ? "selected" : ""}>${direction || "Select direction"}</option>`)
          .join("");
        return `<div class="pole-action-buttons">
          <button class="mini-btn ${resolution === "UG" ? "active-action" : ""}" type="button" data-toggle-ug data-pole="${escapeHtml(poleId)}">UG</button>
          <button class="mini-btn ${resolution === "PCO" ? "active-action" : ""}" type="button" data-toggle-pco data-pole="${escapeHtml(poleId)}">PCO</button>
          ${!isMetronet && !isOlsson ? `<button class="mini-btn ${poleInsetEnabled ? "active-action" : ""}" type="button" data-toggle-pole-inset data-pole="${escapeHtml(poleId)}" ${resolution !== "AERIAL" ? "disabled" : ""} title="${resolution !== "AERIAL" ? "Pole Inset is disabled while this pole is UG or PCO" : "Add or remove the Pole Inset Make Ready"}">Pole Inset</button>` : ""}
          ${isOlsson ? `<button class="mini-btn ${pole?.redTagActive ? "active-action" : ""}" type="button" data-toggle-red-tag data-pole="${escapeHtml(poleId)}" title="Add or remove the Olsson OPPD Red Tag Make Ready">Red Tag</button>` : ""}
          <button class="mini-btn ${riserEnabled ? "active-action" : ""}" type="button" data-toggle-riser data-pole="${escapeHtml(poleId)}" ${riserAvailable ? "" : "disabled"} title="${riserAvailable ? "Add or remove the pole Riser Make Ready" : "Riser is disabled while this pole is UG or PCO"}">Riser</button>
          ${isOlsson ? `<button class="mini-btn ${pole?.lessThan12CommClearanceActive ? "active-action" : ""}" type="button" data-toggle-less-than-12-comm data-pole="${escapeHtml(poleId)}" title="Add or remove LESS THAN 12&quot; COMM CLEARANCE REQUESTED.">&lt;12&quot; Comm</button>` : ""}
          ${isOlsson ? `<button class="mini-btn ${secondaryDripLoopActive ? "active-action" : ""}" type="button" data-toggle-secondary-drip-loop data-pole="${escapeHtml(poleId)}" title="Add or remove the secondary drip-loop Make Ready">Secure Secondary Drip Loop</button>` : ""}
          ${(isIntec || isMetronet) ? `<button class="mini-btn ${ancEnabled ? "active-action" : ""}" type="button" data-toggle-anc data-pole="${escapeHtml(poleId)}" ${ancAvailable ? "" : "disabled"} title="${ancAvailable ? `Add or remove the ${isMetronet ? "MidAm" : "INTEC"} ANC Make Ready` : "ANC is disabled while this pole is UG or PCO"}">ANC</button>` : ""}
          ${(isMetronet || isOlsson) ? `<button class="mini-btn ${ohgEnabled ? "active-action" : ""}" type="button" data-toggle-ohg data-pole="${escapeHtml(poleId)}" ${ohgAvailable ? "" : "disabled"} title="${ohgAvailable ? `Add or remove the ${isOlsson ? "Olsson OPPD" : "MidAm"} OHG Make Ready` : "OHG is disabled while this pole is UG or PCO"}">OHG</button>` : ""}
        </div>
        ${poleInsetEnabled ? `<div class="pole-action-fields action-detail-grid pole-inset-controls">
          <label class="pole-action-field">
            <span>Pole Inset Reason</span>
            <select class="input" data-scope="pole" data-pole="${escapeHtml(poleId)}" data-field="poleInsetReason">
              <option value="OVERLOADED" ${poleInsetReason === "OVERLOADED" ? "selected" : ""}>Overloaded</option>
              <option value="FAILING_CLEARANCES" ${poleInsetReason === "FAILING_CLEARANCES" ? "selected" : ""}>Failing clearances</option>
            </select>
          </label>
          <label class="pole-action-field">
            <span>Pole Inset Span</span>
            <select class="input" data-scope="pole" data-pole="${escapeHtml(poleId)}" data-field="poleInsetSpanId">
              ${(poleInsetDetails.spanOptions || []).length
                ? poleInsetDetails.spanOptions.map(option => `<option value="${escapeHtml(option.spanId)}" ${option.spanId === poleInsetDetails.spanId ? "selected" : ""}>${escapeHtml(option.label)}</option>`).join("")
                : '<option value="" selected disabled>No connected span available</option>'}
            </select>
          </label>
        </div>
        <div class="pole-action-fields action-detail-grid pole-inset-details">
          <div class="pole-action-field pole-action-readout"><span>Max Height at Midspan</span><strong>${escapeHtml(poleInsetDetails.maxHeightAtMidspan || "—")}</strong></div>
          <label class="pole-action-field"><span>Pole Inset HOA</span><input class="input height-input" data-scope="pole" data-pole="${escapeHtml(poleId)}" data-field="poleInsetHeight" value="${escapeHtml(poleInsetDetails.selectedHeight || "")}" placeholder="Enter height"></label>
          <div class="pole-action-field pole-action-readout"><span>End Drop to Pole</span><strong>${escapeHtml(poleInsetDetails.endDropToPole || "—")}</strong></div>
          <div class="pole-action-field pole-action-readout"><span>End Drop to Next Pole</span><strong>${escapeHtml(poleInsetDetails.endDropToNextPole || "—")}</strong></div>
        </div>` : ""}
        ${isOlsson && pole?.redTagActive ? `<div class="pole-action-fields action-detail-grid red-tag-fields">
          <label class="pole-action-field"><span>Red Tag Minimum Height</span><input class="input height-input" data-scope="pole" data-pole="${escapeHtml(poleId)}" data-field="redTagMinHeight" value="${escapeHtml(pole?.redTagMinHeight || "")}"></label>
          <label class="pole-action-field"><span>Red Tag Minimum Class</span><input class="input" data-scope="pole" data-pole="${escapeHtml(poleId)}" data-field="redTagClass" value="${escapeHtml(pole?.redTagClass || "")}"></label>
          <label class="pole-action-field"><span>Red Tag Reason</span><input class="input" data-scope="pole" data-pole="${escapeHtml(poleId)}" data-field="redTagReason" value="${escapeHtml(pole?.redTagReason || "")}" placeholder="CLEARANCES / LOAD FAILURE"></label>
        </div>` : ""}
        ${showUGReason ? `<div class="pole-action-fields action-detail-grid ug-controls">
          ${isIntec ? `<label class="pole-action-field">
            <span>UG Span</span>
            <select class="input" data-scope="pole" data-pole="${escapeHtml(poleId)}" data-field="ugSpanId">
              <option value="ALL" ${selectedUGSpanId.toUpperCase() === "ALL" ? "selected" : ""}>ALL</option>
              ${ugSpanOptions.map(option => `<option value="${escapeHtml(option.spanId)}" ${option.spanId === selectedUGSpanId ? "selected" : ""}>${escapeHtml(option.label)}</option>`).join("")}
            </select>
          </label>
          ` : ""}
        </div>
        ${isIntec && selectedUGSpanId.toUpperCase() !== "ALL" ? `<div class="pole-action-field ug-span-preview">
          <span>UG Make Ready</span>
          <textarea class="input pole-mr-editor" data-scope="pole" data-pole="${escapeHtml(poleId)}" data-field="ugSpanMRText">${escapeHtml(ugSpanTemplate)}</textarea>
        </div>` : `<label class="pole-action-field">
          <span>UG Make Ready</span>
          <textarea class="input pole-mr-editor" data-scope="pole" data-pole="${escapeHtml(poleId)}" data-field="ugMRText">${escapeHtml(ugTemplate)}</textarea>
        </label>`}` : ""}
        ${showPCOEditor ? `<label class="pole-action-field">
          <span>PCO Make Ready</span>
          <textarea class="input pole-mr-editor" data-scope="pole" data-pole="${escapeHtml(poleId)}" data-field="pcoMRText">${escapeHtml(pcoTemplate)}</textarea>
        </label>` : ""}
        ${showRiserDirection ? `<label class="pole-action-field riser-direction-field">
          <span>Riser Direction</span>
          <select class="input" data-scope="pole" data-pole="${escapeHtml(poleId)}" data-field="ugRiserDirection">${directionOptions}</select>
        </label>` : ""}
        ${isOlsson && secondaryDripLoopActive ? `<label class="pole-action-field">
          <span>Secondary Drip Loop HOA</span>
          <input class="input height-input" data-scope="pole" data-pole="${escapeHtml(poleId)}" data-field="secondaryDripLoopHoa" value="${escapeHtml(pole?.secondaryDripLoopHoa || "")}" placeholder="25'8\"">
        </label>` : ""}
        ${ancEnabled ? `<div class="pole-action-fields action-detail-grid">
          ${isIntec ? "" : `<label class="pole-action-field"><span>ANC Size</span><input class="input" data-scope="pole" data-pole="${escapeHtml(poleId)}" data-field="ancSize" value="${escapeHtml(pole?.ancSize || "8\"")}"></label>`}
          <label class="pole-action-field"><span>ANC Distance</span><input class="input height-input" data-scope="pole" data-pole="${escapeHtml(poleId)}" data-field="ancDistance" value="${escapeHtml(ancDistance)}" title="Defaults to one-third of Proposed, rounded up to the next whole foot; edit to override"></label>
          <label class="pole-action-field"><span>ANC Direction</span><select class="input" data-scope="pole" data-pole="${escapeHtml(poleId)}" data-field="ancDirection">${anchorDirectionOptions}</select></label>
          <label class="pole-action-field"><span>DG HOA</span><input class="input height-input" data-scope="pole" data-pole="${escapeHtml(poleId)}" data-field="dgHoa" value="${escapeHtml(ancDGHOA)}" title="Defaults to Proposed when ANC is active; edit to override"></label>
        </div>` : ""}
        ${ohgEnabled ? `<div class="pole-action-fields action-detail-grid">
          <label class="pole-action-field"><span>OHG Direction</span><select class="input" data-scope="pole" data-pole="${escapeHtml(poleId)}" data-field="ohgDirection">${ohgDirectionOptions}</select></label>
          <label class="pole-action-field"><span>OHG HOA</span><input class="input height-input" data-scope="pole" data-pole="${escapeHtml(poleId)}" data-field="ohgHoa" value="${escapeHtml(pole?.ohgHoa || "")}"></label>
        </div>` : ""}`;
      }

      function renderPoleEditableHeader(poleId) {
        const { pole } = poleSummary(poleId);
        const hasProposed = poleHasProposed(poleId);
        const flagging = poleFlaggingSummary(poleId);
        const hasMR = (S.getState().mr || []).some(item => item?.poleId === poleId && String(item.text || "").trim());
        const hasExistingViolations = Boolean(global.Calculations?.getMidamPreMoveViolations?.(poleId)?.length);
        return `<div class="pole-workspace-header">
          <div class="pole-heading-block">
            <div class="pole-heading-row">
              <h3 class="inline-pole-id" id="pole-${escapeHtml(poleId)}" data-inline-pole-edit data-pole="${escapeHtml(poleId)}" title="Double-click to edit pole ID">${escapeHtml(poleId)}</h3>
              <div class="pole-header-actions">
                <button class="mini-btn pole-header-icon-btn" type="button" data-hide-pole data-pole="${escapeHtml(poleId)}" title="Hide this pole from the workspace" aria-label="Hide ${escapeHtml(poleId)}"><svg viewBox="0 0 24 24" aria-hidden="true" focusable="false"><path d="M3 3l18 18M10.6 10.6a2 2 0 002.8 2.8M9.9 5.2A10.8 10.8 0 0112 5c5.5 0 9 7 9 7a15 15 0 01-3.1 3.9M6.2 6.2C3.9 7.8 3 12 3 12s3.5 7 9 7a10.8 10.8 0 004.1-.8" /></svg></button>
                <button class="mini-btn danger-action pole-header-icon-btn" type="button" data-delete-pole data-pole="${escapeHtml(poleId)}" title="Delete this pole and its connected data" aria-label="Delete ${escapeHtml(poleId)}"><svg viewBox="0 0 24 24" aria-hidden="true" focusable="false"><path d="M4 7h16M9 7V4h6v3m-9 0l1 13h10l1-13M10 11v5m4-5v5" /></svg></button>
              </div>
            </div>
            <div class="pole-meta">
              ${pole.isGenerated ? `<span class="badge warning">Editable generated other pole</span>` : ""}
              ${flagging.resolution ? `<span class="badge ${flagging.resolution.toLowerCase()}">${flagging.resolution}</span>` : ""}
              ${!flagging.resolution && flagging.calculationIssueCount ? `<span class="badge danger">Flagging ${flagging.calculationIssueCount}</span>` : ""}
              ${flagging.heightCritical ? `<span class="badge danger" title="Critical pole height issue">&#9888; Height Critical</span>` : ""}
              ${hasExistingViolations ? `<span class="badge ev" title="Existing violations">EV</span>` : ""}
              ${hasMR ? `<span class="badge warning" title="Make Ready exists for this pole">MR</span>` : ""}
              ${hasProposed ? `<span class="badge changed">Proposed</span>` : ""}
            </div>
          </div>
          <div class="pole-kpis two-up">
            <label>Low Power on Pole<input class="input height-input" data-scope="pole" data-pole="${escapeHtml(poleId)}" data-field="lowPower" value="${escapeHtml(pole.lowPower || "")}"></label>
            <label>Max Height on Pole<input class="input height-input muted-input" value="${escapeHtml(pole.maxCommHeight || "")}" readonly></label>
          </div>
        </div>`;
      }

      function renderSpanProposedTable(poleId) {
        const spans = proposedSpansForPole(poleId);
        const pole = S.getPole(poleId);
        const settings = S.getState().settings || {};
        const isSimplifiedMetronetProposed = String(settings.projectProfile || "").toUpperCase() === "METRONET"
          && (String(settings.metronetWI || "").toUpperCase() === "CSU"
            || String(settings.metronetWI || "").toUpperCase() === "MIDAM"
            || String(settings.metronetWI || "").toUpperCase() === "COMED"
            || String(settings.proposedOwner || "").toUpperCase() === "MNT");
        const showEndDrop = !isSimplifiedMetronetProposed;
        const showNextPoleProposed = !isSimplifiedMetronetProposed;
        const showOcalcMS = !isSimplifiedMetronetProposed;
        const compactBackspanTable = spans.length > 0
          && !pole?.standaloneProposedHOA
          && spans.every(span => isBackspan(span.sourceSpanId ? S.getSpan(span.sourceSpanId) || span : span));
        const showStandalone = !spans.length || Boolean(pole?.standaloneProposedHOA);
        const standaloneFlagging = global.Calculations.evaluateSpanSideFlagging({
          spanId: "",
          poleId,
          proposedHOA: pole?.standaloneProposedHOA || ""
        });
        return `<div class="table-wrap"><table class="span-proposed-table wide-table${compactBackspanTable ? " backspan-proposed-only-table" : ""}">
          <thead><tr>
            <th>Span</th><th>Proposed</th>${compactBackspanTable ? "" : `${showEndDrop ? "<th>End Drop</th>" : ""}${showNextPoleProposed ? "<th>Next Pole Proposed</th>" : ""}<th>Slack</th>${showOcalcMS ? "<th>O-CALC MS</th>" : ""}<th>MS Proposed</th><th>Max Height at MS</th><th>Adjusted Final MS</th><th>Flagging</th><th>Environment</th><th>Environment Clearance</th>`}<th>Actions</th>
          </tr></thead>
          <tbody>${spans.map(span => {
            const side = S.getSpanSide(span.spanId, poleId) || S.upsertSpanSide({ spanId: span.spanId, poleId });
            const physicalSpan = span.sourceSpanId ? S.getSpan(span.sourceSpanId) || span : span;
            const spanUG = physicalSpan.ugActive === true;
            const spanUGReason = physicalSpan.ugReason || "CLEARANCE VIOLATION";
            const aboveMax = side.proposedHOA && H.compareHeights(side.proposedHOA, side.maxCommHeight || pole?.maxCommHeight) === 1;
            const midspanIssue = side.clearanceMSStatus === "PENDING" || side.clearanceMSStatus === "PROBLEM";
            const boltIssue = global.Calculations.evaluateProposedPoleClearance(side);
            const proposedFlaggingIssue = side.proposedFlaggingStatus === "PROBLEM";
            const rowClasses = [
              spanRowClasses(poleId, physicalSpan.spanId),
              side.proposedHOA || side.isSlack || side.ocalcMS || side.msProposed || side.finalMidspan || side.proposedMidspan || side.endDrop || spanUG ? "changed-row" : "",
              aboveMax || midspanIssue || !boltIssue.ok || proposedFlaggingIssue ? "warning-row" : ""
            ].filter(Boolean).join(" ");
            return `<tr class="${rowClasses}">
              <td class="span-cell">
                <strong class="span-main-line">${spanChip(poleId, physicalSpan.spanId)}${poleLink(span.fromPole)} → ${poleLink(span.toPole)}</strong>
                ${spanLengthDisplay(physicalSpan) ? `<span class="span-distance-line">${escapeHtml(spanLengthDisplay(physicalSpan))}</span>` : ""}
              </td>
              <td><input class="input height-input" data-scope="spanSide" data-pole="${escapeHtml(poleId)}" data-span="${escapeHtml(span.spanId)}" data-field="proposedHOA" value="${escapeHtml(side.proposedHOA || "")}"></td>
              ${compactBackspanTable ? "" : `${showEndDrop ? `<td><span class="calculated-value">${escapeHtml(side.endDrop || "")}</span></td>` : ""}
              ${showNextPoleProposed ? `<td><input class="input height-input" data-scope="spanSide" data-pole="${escapeHtml(poleId)}" data-span="${escapeHtml(span.spanId)}" data-field="proposedHOAChange" value="${escapeHtml(side.proposedHOAChange || "")}"></td>` : ""}
              <td><input type="checkbox" data-scope="spanSide" data-pole="${escapeHtml(poleId)}" data-span="${escapeHtml(span.spanId)}" data-field="isSlack" ${side.isSlack ? "checked" : ""} aria-label="Proposed span is slack" title="Mark this Proposed span as Slack"></td>
              ${showOcalcMS ? `<td><input class="input decimal-height-input" data-scope="spanSide" data-pole="${escapeHtml(poleId)}" data-span="${escapeHtml(span.spanId)}" data-field="ocalcMS" value="${escapeHtml(displayDecimalFeetInput(side.ocalcMS, side.proposedMidspan))}"></td>` : ""}
              <td><span class="calculated-value">${escapeHtml(side.msProposed || "")}</span></td>
              <td>${escapeHtml(physicalSpan.midspanMaxCommHeight || "")}</td>
              <td><span class="calculated-value midspan-highlight-display ${spanColorClass(poleId, physicalSpan.spanId)}">${escapeHtml(side.finalMidspan || "")}</span></td>
              <td>${renderCombinedSpanFlagging(side)}</td>
              <td><select class="input environment-input" data-scope="span" data-span="${escapeHtml(physicalSpan.spanId)}" data-field="environment">${renderEnvironmentOptions(physicalSpan.environment)}</select></td>
              <td><input class="input" data-scope="span" data-span="${escapeHtml(physicalSpan.spanId)}" data-field="environmentClearance" value="${escapeHtml(physicalSpan.environmentClearance || "")}"></td>`}
              <td><button class="icon-action danger-action" type="button" data-delete-proposed-span data-pole="${escapeHtml(poleId)}" data-span="${escapeHtml(span.spanId)}" title="Delete proposed span" aria-label="Delete proposed span">&#10005;</button></td>
            </tr>`;
          }).join("")}${showStandalone ? `<tr class="standalone-proposed-row ${pole?.standaloneProposedHOA ? "changed-row" : ""} ${standaloneFlagging.status === "PROBLEM" ? "warning-row" : ""}">
              <td class="span-cell"></td>
              <td><input class="input height-input" data-scope="pole" data-pole="${escapeHtml(poleId)}" data-field="standaloneProposedHOA" value="${escapeHtml(pole?.standaloneProposedHOA || "")}" aria-label="Proposed attachment on terminal pole"></td>
              ${compactBackspanTable ? "" : `${showEndDrop ? "<td></td>" : ""}${showNextPoleProposed ? "<td></td>" : ""}${showOcalcMS ? "<td></td>" : ""}<td></td><td></td><td></td><td></td>`}
              <td>${standaloneFlagging.status === "PROBLEM"
                ? `<div class="flagging-cell"><span class="badge danger">Clearance Issue</span><div class="flagging-message">${escapeHtml(standaloneFlagging.message || "")}</div></div>`
                : (pole?.standaloneProposedHOA ? `<span class="badge changed">OK</span>` : "")}</td>
               ${compactBackspanTable ? "" : "<td></td><td></td><td></td>"}
            </tr>` : ""}</tbody>
        </table></div>`;
      }

      function renderManualProposedSpanStarter(poleId, optionsConfig = {}) {
        const options = Object.keys(S.getState().poles || {})
          .filter(id => id !== poleId)
          .sort((a, b) => a.localeCompare(b, undefined, { numeric: true }))
          .map(id => `<option value="${escapeHtml(id)}">${escapeHtml(id)}</option>`)
          .join("");
        if (!options) return `<p class="muted">No other pole is available for a proposed span.</p>`;
        return `<div class="manual-proposed-starter ${optionsConfig.compact ? "compact-starter" : ""}">
          ${optionsConfig.compact ? "" : `<span class="muted">No spans with midspan are available to propose from this pole.</span>`}
          <label>
            <span>Propose to</span>
            <select class="input" data-manual-proposed-target="${escapeHtml(poleId)}">
              <option value="">Select pole</option>
              ${options}
            </select>
          </label>
          <button class="mini-btn" type="button" data-add-proposed-span data-pole="${escapeHtml(poleId)}">Add Proposed Span</button>
        </div>`;
      }

      function renderCommMovementTable(poleId) {
        const groups = groupedCommsForPole(poleId);
        if (!groups.length) return `<p class="muted">No comms imported from Span.Wire for this pole.</p>`;
        const settings = S.getState().settings || {};
        const isCsuMetronet = String(settings.projectProfile || "").toUpperCase() === "METRONET"
          && (String(settings.metronetWI || "").toUpperCase() === "CSU"
            || String(settings.proposedOwner || "").toUpperCase() === "MNT");
        const showServiceDrop = settings.showServiceDrop !== false;
        const showTransferToNewPole = !isCsuMetronet;
        const showResagServiceDrop = String(settings.projectProfile || "INTEC").toUpperCase() === "INTEC"
          && settings.showResagServiceDrop !== false;
        const showPof = String(settings.projectProfile || "INTEC").toUpperCase() === "INTEC";
        return `<div class="table-wrap"><table class="comm-movement-table">
          <thead><tr>
            <th>Owner/Comm</th><th>Existing HOA</th><th>HOA Change</th><th>Other Pole HOA</th><th>Span</th><th>Max Height at MS</th><th>Midspan</th><th>Flagging</th>${showServiceDrop ? "<th>Service Drop</th>" : ""}${showResagServiceDrop ? "<th>Re-sag Service Drop</th>" : ""}${showTransferToNewPole ? "<th>Transfer to New Pole</th>" : ""}<th>DG</th><th>Other HOA</th>${showPof ? "<th>POF</th>" : ""}<th>Actions</th>
          </tr></thead>
          <tbody>${groups.map(group => {
            const pole = S.getPole(poleId);
            const effective = group.rows.map(row => row.otherHOA || row.existingHOAChange || row.existingHOA).filter(Boolean).sort((a, b) => (H.parseHeight(b) || -Infinity) - (H.parseHeight(a) || -Infinity))[0] || group.existingHOA;
            const aboveMax = effective && H.compareHeights(effective, pole?.maxCommHeight) === 1;
            const changed = Boolean(group.existingHOAChange || group.rows.some(row => (
              row.otherHOA || row.mr || row.transferToNewPole || row.resagServiceDrop
            )));
            const flaggingIssue = group.rows.some(row => row.flaggingStatus === "PROBLEM");
            const flaggingMissing = group.rows.some(row => row.flaggingStatus === "MISSING" || row.flaggingStatus === "MISSING_POWER");
            const rowClasses = [
              changed ? "changed-row" : "",
              aboveMax || flaggingIssue || flaggingMissing ? "warning-row" : ""
            ].filter(Boolean).join(" ");
            return `<tr class="${rowClasses}">
              <td class="inline-edit-cell" data-inline-comm-edit="owner" data-pole="${escapeHtml(poleId)}" data-group-key="${escapeHtml(group.key)}" title="Double-click to edit owner/comm"><span class="badge owner">${escapeHtml(group.owner)}</span>${group.isPof ? ` <span class="badge pof">POF</span>` : ""}</td>
              <td class="inline-edit-cell" data-inline-comm-edit="existingHOA" data-pole="${escapeHtml(poleId)}" data-group-key="${escapeHtml(group.key)}" title="Double-click to edit existing HOA">${escapeHtml(group.existingHOA || "")}</td>
              <td>${renderCommHoaChangeValues(group, poleId)}</td>
              <td>${renderCommRemoteValues(group)}</td>
              <td>${renderCommSpanRefs(group, poleId)}</td>
              <td>${renderCommMaxHeightAtMSValues(group, poleId)}</td>
              <td>${renderCommMidspanValues(group, poleId)}</td>
              <td>${renderCommFlagging(group)}</td>
              ${showServiceDrop ? `<td>${renderCommServiceDropValues(group, poleId)}</td>` : ""}
              ${showResagServiceDrop ? `<td>${renderCommResagValues(group, poleId)}</td>` : ""}
              ${showTransferToNewPole ? `<td>${renderCommTransferValues(group, poleId)}</td>` : ""}
              <td>${renderCommDownGuyValues(group, poleId)}</td>
              <td>${renderCommOtherHoaValues(group, poleId)}</td>
              ${showPof ? `<td>${group.pofEligible ? `<label class="equipment-action-control" title="Activar POF manualmente después de Re-sag"><input type="checkbox" data-scope="commGroup" data-pole="${escapeHtml(poleId)}" data-group-key="${escapeHtml(group.key)}" data-field="pofActive" ${group.pofActive ? "checked" : ""}><span>POF</span></label>` : `<span class="muted">&mdash;</span>`}</td>` : ""}
              <td><div class="row-actions">
                <button class="icon-action" type="button" data-edit-comm-spans data-pole="${escapeHtml(poleId)}" data-group-key="${escapeHtml(group.key)}" title="Edit comm spans" aria-label="Edit comm spans">&#8644;</button>
                <button class="icon-action danger-action" type="button" data-delete-comm data-pole="${escapeHtml(poleId)}" data-group-key="${escapeHtml(group.key)}" title="Delete full comm" aria-label="Delete full comm">&#10005;</button>
              </div></td>
            </tr>`;
          }).join("")}</tbody>
        </table></div>`;
      }

      function renderPowerTable(poleId) {
        const powerGroups = [];
        const heightGroups = new Map();
        S.getSpanPowerForPole(poleId).forEach(row => {
          const sourceHeight = H.parseHeight(row.attachmentHeight || "");
          if (sourceHeight === null) {
            powerGroups.push({ rows: [row] });
            return;
          }
          const key = String(sourceHeight);
          let group = heightGroups.get(key);
          if (!group) {
            group = { rows: [] };
            heightGroups.set(key, group);
            powerGroups.push(group);
          }
          group.rows.push(row);
        });
        powerGroups.sort((a, b) => {
          const aHeight = H.parseHeight(a.rows[0]?.attachmentHeight || "");
          const bHeight = H.parseHeight(b.rows[0]?.attachmentHeight || "");
          if (aHeight === null && bHeight === null) return 0;
          if (aHeight === null) return 1;
          if (bHeight === null) return -1;
          return bHeight - aHeight;
        });
        if (!powerGroups.length) return `<p class="muted">No power wires were imported for this pole.</p>`;
        return `<div class="table-wrap"><table class="power-table">
          <thead><tr><th>Span</th><th>Type</th><th>Attachment Height</th><th>HOA Change</th><th>Midspan</th><th>Actions</th></tr></thead>
          <tbody>${powerGroups.map(group => {
            const groupRows = [...group.rows].sort((a, b) =>
              String(a.spanId || "").localeCompare(String(b.spanId || ""), undefined, { numeric: true })
            );
            const representative = groupRows[0];
            const sharedHoaChange = groupRows.find(row => row.attachmentHeightChange)?.attachmentHeightChange || "";
            const spanEntries = groupRows.map(row => {
              const span = S.getSpan(row.spanId);
              return `<div class="power-group-entry">${span ? `${poleLink(span.fromPole)} → ${poleLink(span.toPole)}` : ""}</div>`;
            }).join("");
            const typeEntries = groupRows.map(row =>
              `<div class="power-group-entry"><span class="badge warning">${escapeHtml(row.label || "Power")}</span></div>`
            ).join("");
            const midspanEntries = groupRows.map(row => {
              const originalMidspan = String(row.midspan || "");
              const calculatedMidspan = String(row.calculatedMidspan || originalMidspan);
              const originalTitle = originalMidspan
                ? `Original Midspan before movements: ${originalMidspan}`
                : "No original Midspan was imported";
              return `<div class="power-group-entry"><input class="input height-input" data-scope="spanPower" data-power-key="${escapeHtml(row.key || "")}" data-pole="${escapeHtml(poleId)}" data-span="${escapeHtml(row.spanId)}" data-field="midspan" value="${escapeHtml(calculatedMidspan)}" title="${escapeHtml(originalTitle)}" aria-label="Power Midspan for ${escapeHtml(row.spanId)}"></div>`;
            }).join("");
            const actionEntries = groupRows.map(row => {
              const span = S.getSpan(row.spanId);
              const label = span ? `${span.fromPole} to ${span.toPole}` : "unlinked span";
              return `<button class="icon-action danger-action" type="button" data-delete-span-power data-power-key="${escapeHtml(row.key || "")}" data-pole="${escapeHtml(poleId)}" data-span="${escapeHtml(row.spanId)}" title="Delete ${escapeHtml(row.label || "power wire")} on ${escapeHtml(label)}" aria-label="Delete ${escapeHtml(row.label || "power wire")} on ${escapeHtml(label)}">&#10005;</button>`;
            }).join("");
            return `<tr>
              <td class="span-cell"><div class="power-group-stack">${spanEntries}</div></td>
              <td><div class="power-group-stack">${typeEntries}</div></td>
              <td><input class="input height-input" data-scope="spanPower" data-power-key="${escapeHtml(representative.key || "")}" data-pole="${escapeHtml(poleId)}" data-span="${escapeHtml(representative.spanId)}" data-field="attachmentHeight" value="${escapeHtml(representative.attachmentHeight || "")}" aria-label="Power attachment height shared by wires at this height"></td>
              <td><input class="input height-input" data-scope="spanPower" data-power-key="${escapeHtml(representative.key || "")}" data-pole="${escapeHtml(poleId)}" data-span="${escapeHtml(representative.spanId)}" data-field="attachmentHeightChange" value="${escapeHtml(sharedHoaChange)}" aria-label="Power HOA Change" title="Target attachment height shared by power wires at this pole and height"></td>
              <td><div class="power-group-stack">${midspanEntries}</div></td>
              <td><div class="power-group-stack">${actionEntries}</div></td>
            </tr>`;
          }).join("")}</tbody>
        </table></div>`;
      }

      function renderCollapseEyeIcon() {
        return `<span class="collapse-hint" aria-hidden="true">
          <svg class="collapse-show-icon" viewBox="0 0 24 24" focusable="false"><path d="M2.5 12s3.5-7 9.5-7 9.5 7 9.5 7-3.5 7-9.5 7-9.5-7-9.5-7Z"/><circle cx="12" cy="12" r="3"/></svg>
          <svg class="collapse-hide-icon" viewBox="0 0 24 24" focusable="false"><path d="M3 3l18 18M10.6 10.6a2 2 0 002.8 2.8M9.9 5.2A10.8 10.8 0 0112 5c5.5 0 9 7 9 7a15 15 0 01-3.1 3.9M6.2 6.2C3.9 7.8 3 12 3 12s3.5 7 9 7a10.8 10.8 0 004.1-.8"/></svg>
        </span>`;
      }

      function renderEquipmentTable(poleId) {
        const pole = S.getPole(poleId);
        const rows = Array.isArray(pole?.metadata?.powerEquipment) ? pole.metadata.powerEquipment : [];
        if (!rows.length) return `<p class="muted">No Power Equipment was imported for this pole.</p>`;
        const labels = {
          STREETLIGHT: "Streetlight",
          TRANSFORMER: "Transformer",
          RISER: "Power Riser"
        };
        const actionLabels = {
          STREETLIGHT: "Ground",
          TRANSFORMER: "Redress",
          RISER: "Raise"
        };
        return `<div class="table-wrap"><table class="power-table equipment-table">
          <thead><tr><th>Equipment</th><th>Owner</th><th>Attachment Height</th><th>Bottom Height</th><th>Drip Loop Height</th><th>Orientation</th><th>Action</th><th>New HOA</th><th>Max Comm Height</th><th>Actions</th></tr></thead>
          <tbody>${rows.map((row, index) => {
            const category = String(row.category || "").toUpperCase();
            const needsHeight = category === "TRANSFORMER" || category === "RISER";
            const settings = S.getState().settings || {};
            const requiredMidAmGround = category === "STREETLIGHT"
              && String(settings.projectProfile || "").toUpperCase() === "METRONET"
              && String(settings.proposedOwner || "MidAm").toUpperCase() === "MIDAM";
            const isIntecStreetlight = category === "STREETLIGHT"
              && String(settings.projectProfile || "INTEC").toUpperCase() === "INTEC";
            const isMidAmStreetlight = requiredMidAmGround;
            const isRaiseableStreetlight = isIntecStreetlight || isMidAmStreetlight;
            const isRiser = category === "RISER";
            const isTransformer = category === "TRANSFORMER";
            const actionControls = isRaiseableStreetlight
              ? `<div class="equipment-action-stack">
                  <label class="equipment-action-control">
                    <input type="checkbox" data-scope="equipment" data-pole="${escapeHtml(poleId)}" data-equipment-index="${index}" data-field="actionActive" ${row.actionActive || requiredMidAmGround ? "checked" : ""} ${requiredMidAmGround ? "disabled title=\"Required by MidAm\"" : ""}>
                    <span>Ground</span>
                  </label>
                  <label class="equipment-action-control">
                    <input type="checkbox" data-scope="equipment" data-pole="${escapeHtml(poleId)}" data-equipment-index="${index}" data-field="raiseActive" ${row.raiseActive ? "checked" : ""}>
                    <span>Raise</span>
                  </label>
                </div>`
              : isRiser
                ? `<div class="equipment-action-stack">
                    <label class="equipment-action-control">
                      <input type="checkbox" data-scope="equipment" data-pole="${escapeHtml(poleId)}" data-equipment-index="${index}" data-field="actionActive" ${row.actionActive ? "checked" : ""}>
                      <span>Raise</span>
                    </label>
                    <label class="equipment-action-control">
                      <input type="checkbox" data-scope="equipment" data-pole="${escapeHtml(poleId)}" data-equipment-index="${index}" data-field="secureActive" ${row.secureActive ? "checked" : ""}>
                      <span>Secure</span>
                    </label>
                  </div>`
                : `<label class="equipment-action-control">
                  <input type="checkbox" data-scope="equipment" data-pole="${escapeHtml(poleId)}" data-equipment-index="${index}" data-field="actionActive" ${row.actionActive || requiredMidAmGround ? "checked" : ""} ${requiredMidAmGround ? "disabled title=\"Required by MidAm\"" : ""}>
                  <span>${escapeHtml(isTransformer ? "Secure" : (actionLabels[category] || "Apply"))}</span>
                </label>`;
            const heightField = isRaiseableStreetlight ? "raiseHeight" : "actionHeight";
            const heightEnabled = isRaiseableStreetlight ? row.raiseActive : row.actionActive;
            return `<tr>
            <td><span class="badge warning" title="${escapeHtml(row.type || "")}">${escapeHtml(labels[row.category] || row.category || "Equipment")}</span></td>
            <td class="inline-edit-cell" data-inline-equipment-edit="owner" data-pole="${escapeHtml(poleId)}" data-equipment-index="${index}" title="Double-click to edit equipment owner">${escapeHtml(row.owner || "")}</td>
            <td class="inline-edit-cell" data-inline-equipment-edit="attachmentHeight" data-pole="${escapeHtml(poleId)}" data-equipment-index="${index}" title="Double-click to edit attachment height">${escapeHtml(row.attachmentHeight || "")}</td>
            <td>${escapeHtml(row.bottomHeight || "")}</td>
            <td>${escapeHtml(row.dripLoopHeight || "")}</td>
            <td>${escapeHtml(row.orientation || "")}</td>
            <td>${actionControls}</td>
            <td>${needsHeight || isRaiseableStreetlight ? `<input class="input height-input equipment-action-height" data-scope="equipment" data-pole="${escapeHtml(poleId)}" data-equipment-index="${index}" data-field="${heightField}" value="${escapeHtml(row[heightField] || "")}" ${heightEnabled ? "" : "disabled"} ${isRaiseableStreetlight ? `title="Maximum raise: 1 foot above ${escapeHtml(row.attachmentHeight || "the imported HOA")}"` : ""}>` : `<span class="muted">&mdash;</span>`}</td>
            <td><strong>${escapeHtml(global.Calculations.getPowerEquipmentCeiling(row) || "")}</strong></td>
            <td><button class="icon-action danger-action" type="button" data-delete-equipment data-pole="${escapeHtml(poleId)}" data-equipment-index="${index}" title="Delete equipment" aria-label="Delete equipment">&#10005;</button></td>
          </tr>`;
          }).join("")}</tbody>
        </table></div>`;
      }

      function renderPoleWorkspace(poleId) {
        const pole = S.getPole(poleId);
        const hasCommMovements = S.getSpanCommsForPole(poleId).some(row => Boolean(row.existingHOAChange || row.otherHOA));
        const commMovementsActive = hasCommMovements && pole?.commMovementsActive !== false;
        const resolution = poleResolutionOf(pole);
        return `<article class="pole-workspace-card ${resolution === "UG" ? "ug-active" : ""} ${resolution === "PCO" ? "pco-active" : ""}" data-pole-card="${escapeHtml(poleId)}">
          ${renderPoleEditableHeader(poleId)}
          <div class="workspace-grid">
            <section class="subsection wide" id="spans-${escapeHtml(poleId)}">
              <div class="subsection-title-row">
                <h4>Proposed by Span</h4>
                ${renderManualProposedSpanStarter(poleId, { compact: true })}
              </div>
              ${renderSpanProposedTable(poleId)}
            </section>
            <section class="subsection wide" id="comms-${escapeHtml(poleId)}">
              <div class="subsection-title-row">
                <h4>Existing Comm Movements</h4>
                <div class="pole-action-buttons">
                  <button class="mini-btn ${commMovementsActive ? "active-action" : ""}" type="button" data-toggle-comm-movements data-pole="${escapeHtml(poleId)}" ${hasCommMovements ? "" : "disabled"} title="${hasCommMovements ? (commMovementsActive ? "Ignore comm HOA changes in calculations" : "Use comm HOA changes in calculations") : "Enter a comm HOA Change first"}">Comm Moves ${commMovementsActive ? "ON" : "OFF"}</button>
                  <button class="mini-btn" type="button" data-add-comm data-pole="${escapeHtml(poleId)}">Add Comm</button>
                  <button class="mini-btn danger-action" type="button" data-clear-comm-movements data-pole="${escapeHtml(poleId)}" ${hasCommMovements ? "" : "disabled"} title="Clear all comm movement heights">Clear Moves</button>
                </div>
              </div>
              <p class="muted">Move each existing comm to a new height. When the other pole on the same span changes, the calculated Midspan updates.</p>
              ${renderCommMovementTable(poleId)}
            </section>
            <details class="subsection wide power-section">
              <summary class="subsection-collapse-summary">
                <h4>Imported Power / Clearance</h4>
                ${renderCollapseEyeIcon()}
              </summary>
              ${renderPowerTable(poleId)}
            </details>
            <details class="subsection wide equipment-section" open>
              <summary class="subsection-collapse-summary">
                <h4>Equipment</h4>
                ${renderCollapseEyeIcon()}
              </summary>
              ${renderEquipmentTable(poleId)}
            </details>
            <section class="subsection make-ready-section" id="warnings-${escapeHtml(poleId)}">
              <div class="subsection-title-row">
                <h4>Make Ready</h4>
                <button class="mini-btn" type="button" data-copy-mr data-pole="${escapeHtml(poleId)}">Copy</button>
              </div>
              ${renderMRText(poleId)}
            </section>
            <section class="subsection pole-actions-section">
              <h4>Pole Actions</h4>
              ${renderPoleActions(poleId)}
            </section>
            ${renderMidamPreMoveViolations(poleId)}
          </div>
        </article>`;
      }

      function renderAllPolesWorkspace() {
        const poleIds = filteredPoles({ includeHidden: false });
        const hiddenCount = (S.getState().ui.hiddenPoleIds || []).filter(id => S.getPole(id)).length;
        els.polesOverview.innerHTML = poleIds.map(id => {
          try {
            return renderPoleWorkspace(id);
          } catch (error) {
            console.error(`Error rendering pole ${id}`, error);
            return `<article class="pole-workspace-card warning-row" data-pole-card="${escapeHtml(id)}">
              <div class="pole-workspace-header">
                <div>
                  <h3>${escapeHtml(id)}</h3>
                  <span class="badge danger">Data Error</span>
                </div>
              </div>
              <p class="muted">This pole was imported, but it has incomplete or unexpected data. Review the Excel file or import again.</p>
            </article>`;
          }
        }).join("") || `<div class="detail-placeholder">${hiddenCount
          ? `${hiddenCount} pole${hiddenCount === 1 ? " is" : "s are"} hidden. Select a Hidden pole in the index to show it again.`
          : "No poles to show."}</div>`;
        wireEditableEvents(els.polesOverview);
        bindScrollLinks(els.polesOverview);
        bindLocalActions(els.polesOverview);
      }

      return { spanLabel, shortSpanLabel, spanLengthDisplay, poleLink, displayMidspan, displayDecimalFeetInput, connectedSpansSorted, spanColorClass, spanRowClasses, spanChip, spanColorDot, commOwnerLabel, normalizedHeightLabel, commGroupKey, spanHasRealMidspan, isSpanEligibleForProposed, proposedSpansForPole, groupedCommsForPole, commMidspanEntries, renderCommSpanRefs, renderCommServiceDropValues, renderCommDownGuyValues, renderCommTransferValues, renderCommResagValues, renderCommMidspanValues, renderCommMaxHeightAtMSValues, renderCommRemoteValues, renderCommFlagging, poleIdsForSpan, affectedPoleIdsForElement, bindScrollLinks, replacePoleCard, renderAffectedPoles, noteList, commMidspanNote, poleClearanceNote, spanSideClearanceNote, renderMidspanClearanceStatus, renderCombinedSpanFlagging, renderCommHoaChangeValues, renderCommOtherHoaValues, renderClearanceStatus, renderClearanceSettings, updateAutoCalculateButtonState, updateFollowProposedButtonState, renderPoleClassResults, poleClassSeverity, ...excelReviewRenderer, hasCriticalPoleHeightIssue, renderPoleClassTabWarning, poleClassRowClass, renderPoleClassStatus, renderAnsiReferenceTable, renderActiveView, renderEnvironmentOptions, poleSummary, isSelfSupportingComm, poleHasProposed, poleFlaggingSummary, renderSummary, filteredPoles, renderPoleListItem, renderPoleLists, renderMRText, formatExistingViolationsText, renderMidamPreMoveViolations, renderPoleActions, renderPoleEditableHeader, renderSpanProposedTable, renderManualProposedSpanStarter, renderCommMovementTable, renderPowerTable, renderEquipmentTable, renderPoleWorkspace, renderAllPolesWorkspace };
    }
  };
})(window);
