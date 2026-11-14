(function (global) {
  "use strict";

  function create(context = {}) {
    const { S, H, getCalculations, openAppDialog, escapeHtml } = context;
    const calculations = () => typeof getCalculations === "function" ? getCalculations() : context.Calculations;
    const ownerCanonicalOf = row => S.getOwnerCanonical(row);
    const spanRelationOf = span => S.getSpanRelation(span);

    function commPairingRowLabel(row) {
      const owner = S.getOwnerDisplay(row) || row.owner || "Comm";
      const rawCable = row.size || row.cableType || row.wireType || row.type || "Type not provided";
      const cableParts = String(rawCable).split(">").map(part => part.trim()).filter(Boolean);
      const cable = cableParts.length > 2 ? cableParts.slice(0, -1).join(" > ") : rawCable;
      return owner + " · " + cable;
    }

    function commPairingRowKey(row) {
      return S.keyForSpanComm(row.spanId, row.poleId, ownerCanonicalOf(row), row.wireId || "");
    }

    function commPairingIdentityKey(identity) {
      if (!identity) return "";
      return S.keyForSpanComm(identity.spanId, identity.poleId, identity.ownerCanonical, identity.wireId || "");
    }

    function commPairingSpanGroups(poleId) {
      const grouped = new Map();
      S.getConnectedSpans(poleId)
        .filter(span => span.fromPole === poleId && !/back\s*span|backspan/i.test(spanRelationOf(span)))
        .forEach(span => {
          const otherPoleId = S.getOtherPoleId(span, poleId);
          if (!otherPoleId) return;
          if (!grouped.has(otherPoleId)) grouped.set(otherPoleId, []);
          grouped.get(otherPoleId).push(span);
        });
      return Array.from(grouped, ([otherPoleId, spans]) => {
        const fore = spans.find(span => /fore\s*span|forespan/i.test(spanRelationOf(span)));
        const preferred = fore || spans[0];
        const relation = spanRelationOf(preferred) || "Span";
        return {
          otherPoleId,
          spanId: preferred.spanId,
          label: relation + " · " + preferred.fromPole + " → " + (preferred.toPole || otherPoleId)
        };
      }).sort((a, b) => {
        const aFore = /fore\s*span|forespan/i.test(a.label);
        const bFore = /fore\s*span|forespan/i.test(b.label);
        return Number(bFore) - Number(aFore) || a.label.localeCompare(b.label, undefined, { numeric: true });
      });
    }

    function commPairingRowsForSpan(poleId, spanId, side = "local") {
      const pair = calculations().getCommPairingRows(spanId, poleId);
      const rows = side === "remote" ? pair.remoteRows : pair.localRows;
      return rows.slice().sort((a, b) => commPairingRowLabel(a).localeCompare(commPairingRowLabel(b))
        || (H.parseHeight(b.existingHOA) ?? -Infinity) - (H.parseHeight(a.existingHOA) ?? -Infinity)
        || String(a.wireId || "").localeCompare(String(b.wireId || "")));
    }

    function commPairingPendingSelections(poleId, spanGroup, pendingByOtherPole) {
      const pending = pendingByOtherPole.get(spanGroup.otherPoleId);
      const selections = new Map();
      if (!pending?.pairs) return selections;
      const localRows = commPairingRowsForSpan(poleId, spanGroup.spanId, "local");
      const remoteRows = commPairingRowsForSpan(poleId, spanGroup.spanId, "remote");
      pending.pairs.forEach(pair => {
        const local = localRows.find(row => commPairingRowKey(row) === commPairingIdentityKey(pair.local));
        if (!local) return;
        const remoteIndex = pair.remote
          ? remoteRows.findIndex(row => commPairingRowKey(row) === commPairingIdentityKey(pair.remote))
          : -1;
        selections.set(commPairingRowKey(local), remoteIndex < 0 ? "" : String(remoteIndex));
      });
      return selections;
    }

    function commMatchIdentityForUI(row) {
      return {
        spanId: row.spanId,
        poleId: row.poleId,
        ownerCanonical: ownerCanonicalOf(row),
        wireId: row.wireId || ""
      };
    }

    function commMatchRemoteSelection(local, remoteRows) {
      const match = calculations().getRemoteCommMatchResult(
        local.spanId, local.poleId, ownerCanonicalOf(local), local.wireId || "", local
      );
      if (!match.row || match.status === "AMBIGUOUS") return "";
      const index = remoteRows.findIndex(row => commPairingRowKey(row) === commPairingRowKey(match.row));
      return index < 0 ? "" : String(index);
    }

    function renderCommPairingTable(poleId, spanGroup, pendingPairs) {
      const otherPoleId = spanGroup.otherPoleId;
      const localRows = commPairingRowsForSpan(poleId, spanGroup.spanId, "local");
      const remoteRows = commPairingRowsForSpan(poleId, spanGroup.spanId, "remote");
      if (!localRows.length || !remoteRows.length) {
        return `<div class="comm-match-empty">No communications are available on both ends of this span.</div>`;
      }
      const showLocalMidspan = localRows.some(row => calculations().getCommMatchMidspanDisplay(row));
      const showRemoteMidspan = remoteRows.some(row => calculations().getCommMatchMidspanDisplay(row));
      const columnWidths = showLocalMidspan
        ? (showRemoteMidspan ? ["24%", "10%", "10%", "46%", "10%"] : ["28%", "12%", "12%", "48%"])
        : (showRemoteMidspan ? ["30%", "12%", "46%", "12%"] : ["34%", "14%", "52%"]);
      const rowsHtml = localRows.map((local, localIndex) => {
        const localKey = commPairingRowKey(local);
        const saved = pendingPairs.get(localKey);
        const defaultSelection = commMatchRemoteSelection(local, remoteRows);
        const selected = saved === undefined ? defaultSelection : saved;
        const compatible = remoteRows.map((remote, remoteIndex) => ({ remote, remoteIndex }))
          .filter(({ remote }) => calculations().canMatchCommRows(local, remote));
        const options = [
          `<option value="">Do not pair</option>`,
          ...compatible.map(({ remote, remoteIndex }) => {
            const remoteMidspan = calculations().getCommMatchMidspanDisplay(remote);
            const label = commPairingRowLabel(remote) + " · On pole " + (remote.existingHOA || "—")
              + (remoteMidspan ? " · Midspan " + remoteMidspan : "");
            return `<option value="${remoteIndex}" ${String(remoteIndex) === selected ? "selected" : ""}>${escapeHtml(label)}</option>`;
          })
        ].join("");
        return `<tr>
          <td class="comm-match-cable-label">${escapeHtml(commPairingRowLabel(local))}</td>
          <td>${escapeHtml(local.existingHOA || "—")}</td>
          ${showLocalMidspan ? `<td>${escapeHtml(calculations().getCommMatchMidspanDisplay(local) || "—")}</td>` : ""}
          <td><select class="input comm-match-select" name="match-local-${localIndex}" data-match-local="${localIndex}" aria-label="Match ${escapeHtml(commPairingRowLabel(local))} at ${escapeHtml(poleId)}">
            ${options}
          </select></td>
          ${showRemoteMidspan ? `<td data-match-remote-midspan="${localIndex}">${escapeHtml(selected === "" ? "—" : calculations().getCommMatchMidspanDisplay(remoteRows[Number(selected)]) || "—")}</td>` : ""}
        </tr>`;
      }).join("");
      return `<div class="comm-match-table-wrap"><table class="comm-match-table${showLocalMidspan ? " has-local-midspan" : ""}">
        <colgroup>${columnWidths.map(width => `<col style="width:${width}">`).join("")}</colgroup>
        <thead><tr><th colspan="${showLocalMidspan ? 3 : 2}">${escapeHtml(poleId)}</th><th colspan="${showRemoteMidspan ? 2 : 1}">${escapeHtml(otherPoleId)}</th></tr>
          <tr><th>Cable</th><th>On pole</th>${showLocalMidspan ? "<th>Midspan</th>" : ""}<th>Match cable</th>${showRemoteMidspan ? "<th>Midspan</th>" : ""}</tr></thead>
        <tbody>${rowsHtml}</tbody>
      </table></div>`;
    }

    async function open(poleId, warningIndex = null) {
      const warnings = calculations().getCommPairingWarnings(poleId);
      const warning = warningIndex === null || warningIndex === undefined || warningIndex === ""
        ? null
        : warnings[Number(warningIndex)];
      const spanGroups = commPairingSpanGroups(poleId);
      if (!spanGroups.length) {
        await openAppDialog({
          title: "Match Cables",
          contentHtml: `<div class="comm-match-empty">This pole has no connected spans yet. Match Cables is available for existing and newly imported jobs once a span is connected.</div>`,
          confirmLabel: "Close",
          getResult: () => []
        });
        return null;
      }
      const defaultGroup = warning
        ? spanGroups.find(group => group.otherPoleId === warning.otherPoleId)
        : spanGroups[0];
      if (warning && defaultGroup) defaultGroup.spanId = warning.spanId || defaultGroup.spanId;
      let activeOtherPoleId = defaultGroup?.otherPoleId || spanGroups[0].otherPoleId;
      const pendingByOtherPole = new Map();
      const content = () => {
        const optionMarkup = spanGroups.map(group => `<option value="${escapeHtml(group.otherPoleId)}" ${group.otherPoleId === activeOtherPoleId ? "selected" : ""}>${escapeHtml(group.label)}</option>`).join("");
        const activeGroup = spanGroups.find(group => group.otherPoleId === activeOtherPoleId) || spanGroups[0];
        const table = renderCommPairingTable(poleId, activeGroup, commPairingPendingSelections(poleId, activeGroup, pendingByOtherPole));
        return `<div class="comm-match-controls"><label for="comm-match-span">Span</label><select class="input" id="comm-match-span" name="match-span" data-match-span>${optionMarkup}</select></div>
          <div data-match-table-body>${table}</div>
          <p class="comm-match-note">Only the same owner and cable class can be paired. Missing type is accepted only when the match is unambiguous. Each cable can be used once.</p>`;
      };
      const rememberCurrentSelection = form => {
        const activeGroup = spanGroups.find(group => group.otherPoleId === activeOtherPoleId) || spanGroups[0];
        const localRows = commPairingRowsForSpan(poleId, activeGroup.spanId, "local");
        const availableRemoteRows = commPairingRowsForSpan(poleId, activeGroup.spanId, "remote");
        if (!localRows.length || !availableRemoteRows.some(remote => localRows.some(local => calculations().canMatchCommRows(local, remote)))) return;
        const pairs = localRows.map((local, index) => {
          const selected = form.querySelector(`[data-match-local="${index}"]`)?.value || "";
          const remote = selected === "" ? null : availableRemoteRows[Number(selected)];
          return { local: commMatchIdentityForUI(local), remote: remote ? commMatchIdentityForUI(remote) : null };
        });
        pendingByOtherPole.set(activeOtherPoleId, { spanId: activeGroup.spanId, pairs });
      };
      const values = await openAppDialog({
        title: "Match Cables",
        contentHtml: content(),
        dialogClass: "comm-match-dialog-wide",
        confirmLabel: "Save Matches",
        onOpen: ({ form }) => {
          const updateExclusiveOptions = () => {
            const selects = Array.from(form.querySelectorAll("[data-match-local]"));
            const selected = new Set(selects.map(select => select.value).filter(Boolean));
            selects.forEach(select => Array.from(select.options).forEach(option => {
              option.disabled = Boolean(option.value && selected.has(option.value) && option.value !== select.value);
            }));
          };
          const updateRemoteDetails = localIndex => {
            const select = form.querySelector(`[data-match-local="${localIndex}"]`);
            const activeGroup = spanGroups.find(group => group.otherPoleId === activeOtherPoleId) || spanGroups[0];
            const remoteRows = commPairingRowsForSpan(poleId, activeGroup.spanId, "remote");
            const remote = select?.value === "" ? null : remoteRows[Number(select?.value)];
            const midspan = form.querySelector(`[data-match-remote-midspan="${localIndex}"]`);
            if (midspan) midspan.textContent = remote ? calculations().getCommMatchMidspanDisplay(remote) || "—" : "—";
          };
          form.addEventListener("change", event => {
            if (event.target.matches("[data-match-span]")) {
              rememberCurrentSelection(form);
              activeOtherPoleId = event.target.value;
              const activeGroup = spanGroups.find(group => group.otherPoleId === activeOtherPoleId) || spanGroups[0];
              form.querySelector("[data-match-table-body]").innerHTML = renderCommPairingTable(
                poleId, activeGroup, commPairingPendingSelections(poleId, activeGroup, pendingByOtherPole)
              );
              updateExclusiveOptions();
              commPairingRowsForSpan(poleId, activeGroup.spanId, "local").forEach((_, index) => updateRemoteDetails(index));
              return;
            }
            if (event.target.matches("[data-match-local]")) {
              updateExclusiveOptions();
              updateRemoteDetails(event.target.dataset.matchLocal);
            }
          });
          updateExclusiveOptions();
        },
        getResult: form => {
          rememberCurrentSelection(form);
          return Array.from(pendingByOtherPole, ([otherPoleId, pending]) => ({ otherPoleId, ...pending }));
        }
      });
      return values?.length ? values : null;
    }

    return { open };
  }

  global.CommPairingDialogModule = { create };
})(window);
