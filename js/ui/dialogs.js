(function (global) {
  "use strict";

  function create(escapeHtml) {
    function closeAppDialog(result = null) {
      const dialog = global.document.querySelector("[data-app-dialog]");
      if (!dialog) return;
      const resolve = dialog._resolveDialog;
      dialog.remove();
      if (resolve) resolve(result);
    }

    function openAppDialog({ title, description = "", fields = [], confirmLabel = "Accept", danger = false }) {
      closeAppDialog(null);
      return new Promise(resolve => {
        const dialog = global.document.createElement("div");
        dialog.className = "app-dialog-backdrop";
        dialog.dataset.appDialog = "true";
        dialog._resolveDialog = resolve;
        dialog.innerHTML = `<section class="app-dialog" role="dialog" aria-modal="true" aria-label="${escapeHtml(title)}">
          <div class="app-dialog-header">
            <h3>${escapeHtml(title)}</h3>
            ${description ? `<p>${escapeHtml(description)}</p>` : ""}
          </div>
          <form class="app-dialog-form">
            ${fields.map(field => {
              if (field.type === "checkbox-list") {
                return `<fieldset class="dialog-checkbox-list">
                  <legend>${escapeHtml(field.label)}</legend>
                  ${field.options.map(option => `<label>
                    <input type="checkbox" name="${escapeHtml(field.name)}" value="${escapeHtml(option.value)}" ${option.checked ? "checked" : ""}>
                    <span>${escapeHtml(option.label)}</span>
                  </label>`).join("")}
                </fieldset>`;
              }
              if (field.type === "select") {
                return `<label class="dialog-field">
                  <span>${escapeHtml(field.label)}</span>
                  <select class="input" name="${escapeHtml(field.name)}" ${field.exclusiveGroup ? `data-exclusive-group="${escapeHtml(field.exclusiveGroup)}"` : ""}>
                    ${field.options.map(option => `<option value="${escapeHtml(option.value)}" ${String(option.value) === String(field.value ?? "") ? "selected" : ""}>${escapeHtml(option.label)}</option>`).join("")}
                  </select>
                </label>`;
              }
              return `<label class="dialog-field">
                <span>${escapeHtml(field.label)}</span>
                <input class="input" name="${escapeHtml(field.name)}" value="${escapeHtml(field.value || "")}" placeholder="${escapeHtml(field.placeholder || "")}">
              </label>`;
            }).join("")}
            <div class="app-dialog-actions">
              <button class="btn" type="button" data-dialog-cancel>Cancel</button>
              <button class="btn ${danger ? "btn-danger-solid" : "btn-primary"}" type="submit">${escapeHtml(confirmLabel)}</button>
            </div>
          </form>
        </section>`;
        global.document.body.appendChild(dialog);
        const form = dialog.querySelector("form");
        form.querySelectorAll("select[data-exclusive-group]").forEach(select => {
          select.addEventListener("change", () => {
            const group = select.dataset.exclusiveGroup;
            const selects = Array.from(form.querySelectorAll("select[data-exclusive-group]")).filter(item => item.dataset.exclusiveGroup === group);
            const claimed = new Set(selects.filter(item => item !== select).map(item => item.value).filter(Boolean));
            selects.forEach(item => Array.from(item.options).forEach(option => {
              option.disabled = Boolean(option.value && claimed.has(option.value) && option.value !== item.value);
            }));
          });
        });
        dialog.querySelector("[data-dialog-cancel]").addEventListener("click", () => closeAppDialog(null));
        dialog.addEventListener("click", event => {
          if (event.target === dialog) closeAppDialog(null);
        });
        dialog.addEventListener("keydown", event => {
          if (event.key === "Escape") closeAppDialog(null);
        });
        form.addEventListener("submit", event => {
          event.preventDefault();
          const result = {};
          fields.forEach(field => {
            if (field.type === "checkbox-list") {
              result[field.name] = Array.from(form.querySelectorAll(`[name="${global.CSS.escape(field.name)}"]:checked`)).map(input => input.value);
              return;
            }
            result[field.name] = form.elements[field.name]?.value || "";
          });
          closeAppDialog(result);
        });
        const firstInput = dialog.querySelector("input, select");
        if (firstInput) firstInput.focus();
      });
    }

    async function confirmInApp(title, description, confirmLabel = "Delete") {
      return Boolean(await openAppDialog({ title, description, confirmLabel, danger: true }));
    }

    return { closeAppDialog, openAppDialog, confirmInApp };
  }

  global.AppDialogs = Object.freeze({ create });
})(window);
