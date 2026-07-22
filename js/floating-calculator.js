(function (global) {
  "use strict";

  // Small independent helper for quick feet/inches math while editing the app.
  function setupFloatingCalculator() {
    const panel = document.getElementById("floatingCalculator");
    const openBtn = document.getElementById("openCalculatorBtn");
    const scrollTopBtn = document.getElementById("scrollTopBtn");
    const closeBtn = document.getElementById("closeCalculatorBtn");
    const expression = document.getElementById("calcExpression");
    const result = document.getElementById("calcResult");
    const H = global.HeightUtils;

    if (!panel || !openBtn || !scrollTopBtn || !closeBtn || !expression || !result || !H) return;

    function convertSingleValue(value) {
      const source = String(value || "").trim();
      if (!source) return "";
      if (/(?:\"|\bin)$/i.test(source) && !source.includes("'")) return H.inchesToHeight(source);
      if (/^[+-]?\d+\.\d+$/.test(source)) return H.decimalFeetToHeight(source);
      if (source.includes("'")) {
        const decimal = H.heightToDecimalFeet(source);
        return decimal === "" ? "" : `${decimal} ft`;
      }
      if (/^[+-]?\d+$/.test(source)) return H.decimalFeetToHeight(source);
      return "";
    }

    function runCalculation() {
      try {
        const source = expression.value.trim();
        if (!source) {
          result.textContent = "";
          return;
        }
        const operation = source.match(/^(.+?)\s*([+-])\s*(.+)$/);
        const output = operation
          ? (operation[2] === "+" ? H.addHeights(operation[1], operation[3]) : H.subtractHeights(operation[1], operation[3]))
          : convertSingleValue(source);
        result.textContent = output || "Invalid input";
      } catch (error) {
        result.textContent = error.message;
      }
    }

    openBtn.addEventListener("click", () => panel.classList.toggle("hidden"));
    scrollTopBtn.addEventListener("click", () => window.scrollTo({ top: 0, behavior: "smooth" }));
    closeBtn.addEventListener("click", () => panel.classList.add("hidden"));
    expression.addEventListener("input", runCalculation);
    expression.addEventListener("keydown", event => {
      if (event.key === "Enter") {
        event.preventDefault();
        runCalculation();
      }
    });
  }

  global.FloatingCalculator = { setupFloatingCalculator };
})(window);
