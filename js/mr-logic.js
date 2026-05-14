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

  function generateMRForComm(spanComm) {
    if (!spanComm) return "";
    if (spanComm.mr && spanComm.mr.trim()) return spanComm.mr.trim();
    const action = detectRaiseLower(spanComm);
    if (!action) return "";
    const owner = spanComm.ownerBase || spanComm.owner || "COMM";
    const verb = action === "Lower" ? "Lower" : "Raise";
    return `${verb} ${owner} from HOA ${spanComm.existingHOA} to HOA ${spanComm.existingHOAChange}.`;
  }

  function generateMRForSpanSide(spanSide) {
    if (!spanSide || !spanSide.proposedHOA) return "";
    const span = S().getSpan(spanSide.spanId);
    const dir = span && span.direction ? ` ${span.direction}` : "";
    const items = [`Attach proposed at HOA ${spanSide.proposedHOA}${dir}.`];
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

    const lines = [];
    S().getSpanSidesForPole(poleId).forEach(side => {
      const text = generateMRForSpanSide(side);
      if (text) lines.push(...text.split(/\n+/).filter(Boolean));
    });
    S().getSpanCommsForPole(poleId).forEach(sc => {
      const text = generateMRForComm(sc);
      if (text) lines.push(text);
    });

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
