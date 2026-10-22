(function (global) {
  "use strict";

  const STORAGE_KEY = "poleCalculatorTheme";

  function create({ els }) {
    function apply(theme, persist = true) {
      const dark = String(theme || "").toLowerCase() === "dark";
      global.document.documentElement.dataset.theme = dark ? "dark" : "light";
      if (els.themeToggleBtn) {
        els.themeToggleBtn.textContent = dark ? "☀" : "☾";
        els.themeToggleBtn.setAttribute("aria-pressed", String(dark));
        els.themeToggleBtn.setAttribute("aria-label", dark ? "Switch to light mode" : "Switch to dark mode");
        els.themeToggleBtn.title = dark ? "Switch to light mode" : "Switch to dark mode";
      }
      global.document.querySelector('meta[name="theme-color"]')?.setAttribute("content", dark ? "#0f172a" : "#eef4ff");
      if (persist) {
        try { global.localStorage?.setItem(STORAGE_KEY, dark ? "dark" : "light"); } catch (error) { /* storage can be unavailable */ }
      }
    }

    function toggle() {
      apply(global.document.documentElement.dataset.theme === "dark" ? "light" : "dark");
    }

    function savedTheme() {
      try { return global.localStorage?.getItem(STORAGE_KEY) || "light"; } catch (error) { return "light"; }
    }

    return { apply, toggle, savedTheme };
  }

  global.AppTheme = Object.freeze({ create });
})(window);
