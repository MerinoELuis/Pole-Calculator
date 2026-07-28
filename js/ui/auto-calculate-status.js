(function (global) {
  "use strict";

  function text(value) {
    return String(value ?? "").trim();
  }

  function escapeHtml(value) {
    return String(value ?? "")
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&#039;");
  }

  function resultForPole(poleId) {
    return global.AppStore?.getPole?.(poleId)?.metadata?.autoCalculateResult || null;
  }

  function statusLabel(status) {
    if (status === "BEST_AVAILABLE") return "BEST AVAILABLE";
    if (status === "CRITICAL") return "CRITICAL";
    if (status === "MANUAL") return "MANUAL REVIEW";
    if (status === "SKIPPED") return "UNCHANGED";
    return "SAFE";
  }

  function statusClass(status) {
    if (status === "BEST_AVAILABLE") return "best";
    if (status === "CRITICAL") return "critical";
    if (status === "MANUAL") return "manual";
    if (status === "SKIPPED") return "skipped";
    return "safe";
  }

  function modeLabel(mode) {
    return mode === "LOW_COMM" ? "LOW COMM" : "TOP COMM";
  }

  function formatHeight(inches) {
    if (inches === null || inches === undefined || !Number.isFinite(Number(inches))) return "";
    return global.HeightUtils?.formatHeight?.(Number(inches)) || String(inches);
  }

  function metric(label, value) {
    return `<span><strong>${escapeHtml(label)}</strong>${escapeHtml(value)}</span>`;
  }

  function resultSignature(result) {
    return JSON.stringify([
      result?.status || "",
      result?.mode || "",
      result?.message || "",
      result?.poleViolationCount || 0,
      result?.midspanViolationCount || 0,
      result?.movedCommCount || 0,
      result?.totalMovementInches || 0,
      result?.idealProposed || "",
      result?.recommendation || ""
    ]);
  }

  function resultMarkup(result) {
    if (!result) return "";
    const status = text(result.status || "SAFE");
    const details = [];
    if (["SAFE", "BEST_AVAILABLE", "CRITICAL"].includes(status)) {
      details.push(metric("Pole issues", String(result.poleViolationCount || 0)));
      details.push(metric("Midspan issues", String(result.midspanViolationCount || 0)));
      details.push(metric("Comms moved", String(result.movedCommCount || 0)));
      details.push(metric("Total movement", formatHeight(result.totalMovementInches || 0)));
      if (result.idealProposed) details.push(metric("Ideal Proposed", result.idealProposed));
    }

    return `<section class="auto-calc-result auto-calc-result--${statusClass(status)}" data-component="auto-calculate-result">
      <div class="auto-calc-result__heading">
        <strong>${escapeHtml(statusLabel(status))}</strong>
        <span>${escapeHtml(modeLabel(result.mode))}</span>
      </div>
      <p>${escapeHtml(text(result.message))}</p>
      ${details.length ? `<div class="auto-calc-result__metrics">${details.join("")}</div>` : ""}
      ${result.recommendation ? `<p class="auto-calc-result__recommendation">${escapeHtml(text(result.recommendation))}</p>` : ""}
    </section>`;
  }

  function refreshButton() {
    const button = global.document?.getElementById?.("autoCalculateBtn");
    if (!button) return false;
    const state = global.AppStore?.getState?.() || {};
    const hasPoles = Object.keys(state.poles || {}).length > 0;
    const mode = String(state.settings?.position || "TOP_COMM").toUpperCase() === "LOW_COMM"
      ? "LOW_COMM"
      : "TOP_COMM";
    button.disabled = !hasPoles;
    button.title = hasPoles
      ? `Calculate the best ${modeLabel(mode)} aerial arrangement. UG/PCO is never selected automatically.`
      : "Import pole data before running Auto Calculate.";
    return hasPoles;
  }

  function refreshCard(card) {
    if (!card?.querySelector || !card?.dataset) return false;
    const poleId = text(card.dataset.poleId || card.dataset.poleCard);
    if (!poleId) return false;
    const existing = card.querySelector("[data-component='auto-calculate-result']");
    const result = resultForPole(poleId);
    if (!result) {
      existing?.remove?.();
      return false;
    }

    const signature = resultSignature(result);
    if (existing?.dataset?.autoCalcSignature === signature) return false;

    const holder = global.document?.createElement?.("div");
    if (!holder) return false;
    holder.innerHTML = resultMarkup(result);
    const replacement = holder.firstElementChild;
    if (!replacement) return false;
    replacement.dataset.autoCalcSignature = signature;

    if (existing) existing.replaceWith(replacement);
    else {
      const anchor = card.querySelector(".pole-card-header, .panel-header, header");
      if (anchor?.insertAdjacentElement) anchor.insertAdjacentElement("afterend", replacement);
      else card.prepend?.(replacement);
    }
    return true;
  }

  function refresh(root = global.document) {
    refreshButton();
    if (!root?.querySelectorAll) return 0;
    let changed = 0;
    root.querySelectorAll("[data-pole-card], [data-component='pole-card']").forEach(card => {
      if (refreshCard(card)) changed += 1;
    });
    return changed;
  }

  global.AutoCalculateStatusUI = {
    escapeHtml,
    statusLabel,
    statusClass,
    modeLabel,
    resultSignature,
    resultMarkup,
    refreshButton,
    refreshCard,
    refresh
  };
})(window);
