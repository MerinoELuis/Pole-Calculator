(function (global) {
  "use strict";

  let activeControl = null;

  global.MobileKeyboardControl = Object.freeze({
    isEnabled: () => Boolean(activeControl?.isEnabled()),
    subscribe(listener) {
      if (typeof listener !== "function") return () => {};
      if (activeControl) return activeControl.subscribe(listener);
      return () => {};
    }
  });

  function create({ els, commitInput }) {
    let mobileHeightInput = null;
    let mobileKeyboardEnabled = false;
    const subscribers = new Set();

    function isViewport() {
      return Boolean(global.matchMedia?.("(max-width: 700px), (orientation: landscape) and (max-height: 700px)").matches);
    }

    function updateToggle() {
      const button = els.mobileKeyboardToggleBtn;
      if (!button) return;
      button.textContent = mobileKeyboardEnabled ? "Keyboard On" : "Keyboard Off";
      button.setAttribute("aria-pressed", String(mobileKeyboardEnabled));
      button.title = mobileKeyboardEnabled ? "Disable the on-page feet and inches keypad" : "Enable the on-page feet and inches keypad";
      button.classList.toggle("btn-primary", mobileKeyboardEnabled);
    }

    function setEnabled(enabled) {
      mobileKeyboardEnabled = Boolean(enabled);
      updateToggle();
      if (!mobileKeyboardEnabled) setOpen(false);
      subscribers.forEach(listener => listener(mobileKeyboardEnabled));
      if (mobileKeyboardEnabled) {
        const active = global.document.activeElement;
        if (active?.classList?.contains("height-input") || active?.classList?.contains("decimal-height-input")) setOpen(true, active);
      }
    }

    function setOpen(open, input = mobileHeightInput) {
      const keyboard = els.mobileHeightKeyboard;
      const shouldOpen = Boolean(mobileKeyboardEnabled && open && input && !input.disabled && !input.readOnly && isViewport());
      mobileHeightInput = shouldOpen ? input : null;
      keyboard?.classList.toggle("open", shouldOpen);
      keyboard?.setAttribute("aria-hidden", String(!shouldOpen));
      global.document.body.classList.toggle("mobile-height-keyboard-open", shouldOpen);
    }

    function applyKey(key) {
      const input = mobileHeightInput;
      if (!input) return;
      input.focus({ preventScroll: true });
      const start = Number.isInteger(input.selectionStart) ? input.selectionStart : input.value.length;
      const end = Number.isInteger(input.selectionEnd) ? input.selectionEnd : start;
      if (key === "Backspace") {
        if (start !== end) input.setRangeText("", start, end, "end");
        else if (start > 0) input.setRangeText("", start - 1, start, "end");
      } else if (key === "Enter") {
        setOpen(false);
        commitInput(input);
        input.blur();
        return;
      } else {
        input.setRangeText(key, start, end, "end");
      }
      input.dispatchEvent(new global.Event("input", { bubbles: true }));
    }

    function bind() {
      const keyboard = els.mobileHeightKeyboard;
      if (!keyboard) return;
      keyboard.querySelectorAll("[data-mobile-key]").forEach(button => {
        button.addEventListener("pointerdown", event => event.preventDefault());
        button.addEventListener("click", () => applyKey(button.dataset.mobileKey || ""));
      });
      global.document.addEventListener("focusin", event => {
        const target = event.target;
        if (!target?.classList?.contains("height-input") && !target?.classList?.contains("decimal-height-input")) setOpen(false);
      });
    }

    const control = {
      isEnabled: () => mobileKeyboardEnabled,
      subscribe(listener) {
        if (typeof listener !== "function") return () => {};
        subscribers.add(listener);
        return () => subscribers.delete(listener);
      },
      updateToggle,
      isViewport,
      ownsInput: input => mobileHeightInput === input,
      setEnabled,
      setOpen,
      bind
    };
    activeControl = control;
    return control;
  }

  global.MobileHeightKeyboard = Object.freeze({ create });
})(window);
