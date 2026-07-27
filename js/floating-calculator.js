(function (global) {
  "use strict";

  // Small independent helper for quick feet/inches math while editing the app.
  function normalizeInchQuotes(value) {
    return String(value ?? "").replace(/''/g, '"');
  }

  function normalizeExpressionInput(input) {
    if (!input) return false;

    const original = String(input.value ?? "");
    const normalized = normalizeInchQuotes(original);
    if (normalized === original) return false;

    const selectionStart = Number.isInteger(input.selectionStart)
      ? input.selectionStart
      : null;
    const selectionEnd = Number.isInteger(input.selectionEnd)
      ? input.selectionEnd
      : selectionStart;

    input.value = normalized;

    if (selectionStart !== null && typeof input.setSelectionRange === "function") {
      const nextStart = normalizeInchQuotes(original.slice(0, selectionStart)).length;
      const nextEnd = normalizeInchQuotes(original.slice(0, selectionEnd)).length;
      input.setSelectionRange(nextStart, nextEnd);
    }

    return true;
  }

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

    function handleExpressionInput() {
      normalizeExpressionInput(expression);
      runCalculation();
    }

    openBtn.addEventListener("click", () => panel.classList.toggle("hidden"));
    scrollTopBtn.addEventListener("click", () => window.scrollTo({ top: 0, behavior: "smooth" }));
    closeBtn.addEventListener("click", () => panel.classList.add("hidden"));
    expression.addEventListener("input", handleExpressionInput);
    expression.addEventListener("keydown", event => {
      if (event.key === "Enter") {
        event.preventDefault();
        normalizeExpressionInput(expression);
        runCalculation();
      }
    });
  }

  global.FloatingCalculator = {
    setupFloatingCalculator,
    normalizeInchQuotes,
    normalizeExpressionInput
  };
})(window);
