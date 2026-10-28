(function (global) {
  "use strict";

  global.PoleNavigation = {
    create(context) {
      const { document, els, mobileKeyboard, recordUndoSnapshot, render, S } = context;

      function setOpen(open) {
        const isOpen = Boolean(open);
        if (isOpen) mobileKeyboard?.setOpen(false);
        document.body.classList.toggle("pole-index-open", isOpen);
        els.poleIndexDrawer?.classList.toggle("open", isOpen);
        els.poleIndexDrawer?.setAttribute("aria-hidden", String(!isOpen));
        if (els.poleIndexDrawer) els.poleIndexDrawer.inert = !isOpen;
        els.poleIndexToggle?.setAttribute("aria-expanded", String(isOpen));
        els.poleIndexBackdrop?.classList.toggle("hidden", !isOpen);
        // Do not focus the search field on open; mobile browsers would show the keyboard.
        if (isOpen && document.activeElement && document.activeElement !== document.body) {
          document.activeElement.blur?.();
        }
      }

      function updateToggleVisibility() {
        if (!els.poleIndexToggle) return;
        if ((S.getState().ui.activeView || "calculator") !== "calculator") {
          els.poleIndexToggle.classList.remove("visible");
          return;
        }
        els.poleIndexToggle.classList.add("visible");
      }

      function scrollToPole(poleId) {
        const card = document.querySelector(`[data-pole-card="${global.CSS.escape(poleId)}"]`);
        if (card) card.scrollIntoView({ behavior: "smooth", block: "start" });
      }

      function selectPole(poleId) {
        const state = S.getState();
        if ((state.ui.hiddenPoleIds || []).includes(poleId)) {
          recordUndoSnapshot();
          state.ui.hiddenPoleIds = state.ui.hiddenPoleIds.filter(id => id !== poleId);
          if (/(?:^unknown(?:-|\b)|-unknown(?:-|\b))/i.test(poleId)) {
            state.ui.revealedUnknownPoleIds = Array.from(new Set([...(state.ui.revealedUnknownPoleIds || []), poleId]));
          }
        }
        state.selectedPoleId = poleId;
        state.selectedSpanId = "";
        state.ui.activeView = "calculator";
        render();
        scrollToPole(poleId);
      }

      return { setOpen, updateToggleVisibility, scrollToPole, selectPole };
    }
  };
})(window);
