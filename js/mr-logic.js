(function (global) {
  "use strict";

  const H = () => global.HeightUtils;
  const S = () => global.AppStore;

  function detectRaiseLower(spanComm) {
    const existing = H().parseHeight(spanComm.existingHOA);
    const changed = H().parseHeight(spanComm.existingHOAChange);
    if (existing === null || changed === null || existing === changed) return null;
    return changed > existing ? "Raise" : "Lower";
  }

  function detectAttachFromSpanSide(spanSide) {
    return Boolean(spanSide && spanSide.proposedHOA);
  }

  function detectOverlash(spanComm) {
    return /overlash/i.test(`${spanComm.notes || ""} ${spanComm.mr || ""}`);
  }

  function detectSlack(spanSide) {
    return /slack/i.test(`${spanSide.notes || ""}`);
  }

  function detectAnchor(spanSide) {
    return /\banc\b|anchor/i.test(`${spanSide.notes || ""}`);
  }

  function detectRiser(spanSide) {
    return /riser/i.test(`${spanSide.notes || ""}`);
  }

  function ownerForMR(spanComm) {
    const raw = spanComm.rawOwner || spanComm.ownerBase || spanComm.owner || "COMM";
    if (/century\s*link/i.test(raw)) return "CTL";
    return String(raw).replace(/^COMMUNICATION\s*>\s*/i, "").replace(/,\s*.*$/, "").trim() || "COMM";
  }

  function proposedOwnerForMR() {
    const settings = S().getState().settings || {};
    return String(settings.proposedOwner || "Wecom").trim() || "Wecom";
  }

  function transferOwnersForPole(poleId) {
    const owners = S().getSpanCommsForPole(poleId)
      .map(ownerForMR)
      .filter(Boolean);
    return Array.from(new Set(owners)).join(" & ");
  }

  function mrHeight(value) {
    return String(value || "").replace(/'(?=\d)/, "' ");
  }

  function applyCase(text) {
    const settings = S().getState().settings || {};
    return (settings.mrCase || "LOWER") === "UPPER" ? text.toUpperCase() : text;
  }

  function generateMRForComm(spanComm) {
    if (!spanComm) return "";
    if (spanComm.mr && spanComm.mr.trim()) return spanComm.mr.trim();
    const action = detectRaiseLower(spanComm);
    if (!action) return "";
    const owner = ownerForMR(spanComm);
    // Service drops use different MR wording than regular comm movement.
    if (spanComm.serviceDrop) return `Relocate ${owner} drop at HOA ${mrHeight(spanComm.existingHOA)} to HOA ${mrHeight(spanComm.existingHOAChange)}.`;
    const verb = action === "Lower" ? "Lower" : "Raise";
    return `${verb} ${owner} from HOA ${mrHeight(spanComm.existingHOA)} to HOA ${mrHeight(spanComm.existingHOAChange)}.`;
  }

  function generateMRForSpanSide(spanSide) {
    if (!spanSide || !spanSide.proposedHOA) return "";
    const span = S().getSpan(spanSide.spanId);
    const dir = span && span.direction ? ` ${span.direction}` : "";
    // Proposed attach owner is configurable because it is project/customer specific.
    const items = [`Attach ${proposedOwnerForMR()} at HOA ${mrHeight(spanSide.proposedHOA)}${dir}.`];
    if (spanSide.clearanceMSReason === "LOW_POWER" && spanSide.clearanceMSIssue) {
      items.push(`Ensure min 30" to low power at midspan.`);
    }
    if (detectSlack(spanSide)) items.push(`Proposed slack span${dir}.`.replace("  ", " "));
    if (detectAnchor(spanSide)) items.push(`PL NEW ANC${dir}.`.replace("  ", " "));
    if (detectRiser(spanSide)) items.push(`PL NEW RISER${dir}.`.replace("  ", " "));
    return items.join("\n");
  }

  function generateMRForSpan(spanId) {
    const sideItems = S().getSpanSidesForSpan(spanId).map(generateMRForSpanSide).filter(Boolean);
    const commItems = S().getSpanCommsForSpan(spanId).map(generateMRForComm).filter(Boolean);
    return [...sideItems, ...commItems];
  }

  function generateMRForPole(poleId) {
    const state = S().getState();
    state.mr = state.mr.filter(item => item.poleId !== poleId);

    const ug = [];
    const power = [];
    const commMoves = [];
    const proposed = [];
    const ensure = [];
    const pole = S().getPole(poleId);
    if (pole?.ugActive && pole.ugReason) {
      ug.push(`Unable to attach due to ${pole.ugReason}.`);
    }
    if (pole?.pcoActive && pole.pcoScope && pole.pcoType && pole.pcoDetail) {
      const transfers = transferOwnersForPole(poleId);
      const transferText = transfers ? ` Transfer ${transfers} to new pole.` : "";
      if (pole.pcoType === "OVERLOAD") {
        power.push(`${pole.pcoScope} pole overloaded by ${pole.pcoDetail}, replace pole.${transferText}`);
      } else {
        power.push(`${pole.pcoScope} clearance violations (${pole.pcoDetail}), replace pole.${transferText}`);
      }
    }
    S().getSpanSidesForPole(poleId).forEach(side => {
      const text = generateMRForSpanSide(side);
      if (text) text.split(/\n+/).filter(Boolean).forEach(line => {
        if (/ensure min 30/i.test(line)) ensure.push(line);
        else proposed.push(line);
      });
    });
    S().getSpanCommsForPole(poleId).forEach(sc => {
      const text = generateMRForComm(sc);
      if (text) commMoves.push(text);
    });

    const lines = [...ug, ...power, ...commMoves, ...proposed, ...ensure].map(applyCase);
    const unique = Array.from(new Set(lines.map(line => line.trim()).filter(Boolean)));
    if (unique.length) state.mr.push({ poleId, spanId: "", owner: "MR", text: unique.join("\n"), imported: false });
    return state.mr.filter(item => item.poleId === poleId);
  }

  function generateAllMR() {
    const state = S().getState();
    state.mr = [];
    Object.keys(state.poles).forEach(generateMRForPole);
    return state.mr;
  }

  global.MRLogic = {
    generateMRForPole,
    generateMRForSpan,
    generateMRForComm,
    generateMRForSpanSide,
    generateAllMR,
    detectAttach: detectAttachFromSpanSide,
    detectRaiseLower,
    detectOverlash,
    detectSlack,
    detectAnchor,
    detectRiser
  };
})(window);
