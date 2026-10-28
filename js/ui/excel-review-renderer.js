(function (global) {
  "use strict";

  global.ExcelReviewRenderer = {
    create(context) {
      const { els, escapeHtml, markDirty, recordUndoSnapshot, S, selectPole, toast } = context;

      function reviewStatusBadge(status) {
        const normalized = String(status || "PASS").toUpperCase();
        const badgeClass = normalized === "ERROR" ? "danger"
          : normalized === "WARNING" || normalized === "NOT_READY" ? "warning"
            : normalized === "IGNORED" || normalized === "NOT_APPLICABLE" ? "ignored"
              : "changed";
        const label = normalized === "NOT_READY" ? "NOT READY"
          : normalized === "NOT_APPLICABLE" ? "N/A"
            : normalized;
        return `<span class="badge ${badgeClass}">${label}</span>`;
      }

      function renderReviewCheck(item) {
        const canIgnore = item.poleId && item.status !== "PASS";
        return `<li class="excel-review-check ${String(item.status || "").toLowerCase()} ${item.level === "low" ? "low-level" : ""} ${item.ignored ? "ignored" : ""}">
          <div class="excel-review-check-heading">
            ${reviewStatusBadge(item.ignored ? "IGNORED" : item.status)}
            <strong>${escapeHtml(item.section || "Review")} · ${escapeHtml(item.title || item.code || "Check")}</strong>
            ${item.level === "low" ? `<span class="badge">Low</span>` : ""}
          </div>
          <p>${escapeHtml(item.message || "")}</p>
          ${item.expected || item.actual ? `<dl class="review-values">
            ${item.expected ? `<div><dt>Expected</dt><dd>${escapeHtml(item.expected)}</dd></div>` : ""}
            ${item.actual ? `<div><dt>Found</dt><dd>${escapeHtml(item.actual)}</dd></div>` : ""}
          </dl>` : ""}
          ${item.details?.length ? `<div class="review-source-details">${item.details.map(detail => (
            `<span>Span Index: ${escapeHtml(detail.spanIndex || "—")} · Span Id: ${escapeHtml(detail.spanId || "—")} · Type: ${escapeHtml(detail.type || "—")} · Linked: ${escapeHtml(detail.linkedCollectionTitle || "—")}</span>`
          )).join("")}</div>` : ""}
          ${canIgnore ? `<div class="review-check-actions"><button class="btn review-ignore-btn" type="button" data-review-ignore-key="${escapeHtml(item.ignoreKey)}" data-review-ignore-value="${item.ignored ? "false" : "true"}">${item.ignored ? "Restore" : "Ignore"}</button></div>` : ""}
        </li>`;
      }

      function renderExcelReviewTabWarning(review) {
        if (!els.excelReviewTabWarning || !els.excelReviewTabWarningCount) return;
        const summary = review?.summary || {};
        const activeCount = Number(summary.errors || 0) + Number(summary.warnings || 0);
        els.excelReviewTabWarningCount.textContent = String(activeCount);
        els.excelReviewTabWarning.classList.toggle("hidden", activeCount === 0);
        els.excelReviewTabWarning.title = activeCount === 1
          ? "1 pole has active Excel Review findings"
          : `${activeCount} poles have active Excel Review findings`;
      }

      function updateExcelReviewWarningActions(review) {
        const controls = [
          { button: els.ignoreHoaWarningsBtn, phase: "HOA", label: "HOA" },
          { button: els.ignoreFinalWarningsBtn, phase: "FINAL", label: "Final QC" }
        ];
        controls.forEach(({ button, phase, label }) => {
          if (!button) return;
          const state = global.ExcelReview?.warningControlState?.(phase) || { count: 0, allIgnored: false };
          const enabled = Boolean(review?.reviewedAt && state.count);
          button.disabled = !enabled;
          button.classList.toggle("btn-disabled", !enabled);
          button.textContent = state.allIgnored ? `Restore all ${label} warnings` : `Ignore all ${label} warnings`;
          button.setAttribute("aria-label", button.textContent);
          button.title = enabled
            ? `${state.count} ${label} warning${state.count === 1 ? "" : "s"} in this review`
            : `No ${label} warnings to ignore`;
        });
      }

      function renderExcelReviewResults() {
        if (!els.excelReviewResults || !global.ExcelReview) return;
        const openPoleIds = new Set(Array.from(els.excelReviewResults.querySelectorAll("details[open][data-review-result-pole]"))
          .map(details => details.dataset.reviewResultPole));
        const review = global.ExcelReview.getReviewState();
        const summary = review.summary || {};
        renderExcelReviewTabWarning(review);
        updateExcelReviewWarningActions(review);
        if (els.excelReviewTimestamp) {
          els.excelReviewTimestamp.textContent = review.reviewedAt
            ? `Last reviewed ${new Date(review.reviewedAt).toLocaleString()}`
            : "No review has been run.";
        }
        if (els.rerunExcelReviewBtn) {
          const hasCollectionSource = Boolean(S.getState().excelReviewSource?.collection?.rows?.length);
          els.rerunExcelReviewBtn.disabled = !hasCollectionSource;
          els.rerunExcelReviewBtn.classList.toggle("btn-disabled", !hasCollectionSource);
        }
        if (!review.reviewedAt) {
          els.excelReviewResults.innerHTML = `<div class="detail-placeholder">Import a raw Excel file to prepare HOA and Final PLA / MR review results.</div>`;
          return;
        }

        const globalChecks = review.globalChecks || [];
        els.excelReviewResults.innerHTML = `
          <div class="excel-review-summary" aria-label="Excel Review summary">
            <div class="review-summary-item error"><strong>${Number(summary.errors || 0)}</strong><span>Errors</span></div>
            <div class="review-summary-item warning"><strong>${Number(summary.warnings || 0)}</strong><span>Warnings</span></div>
            <div class="review-summary-item pass"><strong>${Number(summary.passed || 0)}</strong><span>Passed</span></div>
          </div>
          ${globalChecks.length ? `<section class="excel-review-global">
            <h3>Workbook Issues</h3>
            <ul class="excel-review-checks">${globalChecks.map(renderReviewCheck).join("")}</ul>
          </section>` : ""}
          <div class="excel-review-poles">${(review.results || []).map(result => {
            const visibleChecks = result.checks.filter(item => item.status !== "PASS" || item.applicable === false || item.code === "ALL_APPLICABLE_CHECKS_PASSED");
            return `<div class="excel-review-pole-entry">
              <button class="review-pole-link" type="button" data-review-pole="${escapeHtml(result.poleId)}" title="Open this pole in Calculator" aria-label="Open pole ${escapeHtml(result.poleId)} in Calculator">↗</button>
              <details class="excel-review-pole review-${String(result.overallStatus || "pass").toLowerCase()}" data-review-result-pole="${escapeHtml(result.poleId)}">
              <summary>
                <span class="review-pole-identity"><strong>${escapeHtml(result.poleId)}</strong>${result.sequence ? `<small>Sequence ${escapeHtml(result.sequence)}</small>` : ""}</span>
                <span class="review-phase-status"><small>HOA Review</small>${reviewStatusBadge(result.hoaStatus)}</span>
                <span class="review-phase-status"><small>Final Review</small>${reviewStatusBadge(result.finalStatus)}</span>
                <span class="review-phase-status overall"><small>Overall</small>${reviewStatusBadge(result.overallStatus)}</span>
              </summary>
              <ul class="excel-review-checks">${visibleChecks.map(renderReviewCheck).join("")}</ul>
              </details>
            </div>`;
          }).join("") || `<div class="detail-placeholder">No Collection rows were available for review.</div>`}</div>`;
        els.excelReviewResults.querySelectorAll("details[data-review-result-pole]").forEach(details => {
          if (openPoleIds.has(details.dataset.reviewResultPole)) details.open = true;
        });
        els.excelReviewResults.querySelectorAll("[data-review-pole]").forEach(button => {
          button.addEventListener("click", event => {
            event.preventDefault();
            event.stopPropagation();
            const requested = button.dataset.reviewPole || "";
            const poleId = Object.keys(S.getState().poles || {}).find(id =>
              S.canonicalPoleIdentity(id) === S.canonicalPoleIdentity(requested)
            ) || requested;
            selectPole(poleId);
          });
        });
        els.excelReviewResults.querySelectorAll("[data-review-ignore-key]").forEach(button => {
          button.addEventListener("click", event => {
            event.preventDefault();
            event.stopPropagation();
            recordUndoSnapshot();
            global.ExcelReview.setCheckIgnored(button.dataset.reviewIgnoreKey, button.dataset.reviewIgnoreValue === "true");
            renderExcelReviewResults();
          });
        });
      }

      function setAllExcelReviewWarningsIgnored(phase) {
        if (!global.ExcelReview?.warningControlState || !global.ExcelReview?.setAllWarningsIgnored) return;
        const control = global.ExcelReview.warningControlState(phase);
        if (!control.count) return;
        recordUndoSnapshot();
        const ignored = !control.allIgnored;
        global.ExcelReview.setAllWarningsIgnored(phase, ignored);
        renderExcelReviewResults();
        markDirty();
        const label = phase === "FINAL" ? "Final QC" : "HOA";
        toast(`${ignored ? "Ignored" : "Restored"} all ${label} warnings.`, "info");
      }

      return {
        reviewStatusBadge,
        renderReviewCheck,
        renderExcelReviewTabWarning,
        updateExcelReviewWarningActions,
        renderExcelReviewResults,
        setAllExcelReviewWarningsIgnored
      };
    }
  };
})(window);
