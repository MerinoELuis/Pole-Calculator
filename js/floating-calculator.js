(function (global) {
  "use strict";

  function setupFloatingCalculator() {
    const panel = document.getElementById("floatingCalculator");
    const openBtn = document.getElementById("openCalculatorBtn");
    const closeBtn = document.getElementById("closeCalculatorBtn");
    const runBtn = document.getElementById("runCalculatorBtn");
    const a = document.getElementById("calcA");
    const b = document.getElementById("calcB");
    const op = document.getElementById("calcOperation");
    const result = document.getElementById("calcResult");
    const H = global.HeightUtils;

    if (!panel || !openBtn || !closeBtn || !runBtn || !a || !b || !op || !result || !H) return;

    function runCalculation() {
      try {
        let output = "";
        if (op.value === "add") output = H.addHeights(a.value, b.value);
        if (op.value === "subtract") output = H.subtractHeights(a.value, b.value);
        if (op.value === "toDecimal") output = `${H.heightToDecimalFeet(a.value)} ft`;
        if (op.value === "fromDecimal") output = H.decimalFeetToHeight(a.value);
        if (op.value === "inchesToHeight") output = H.inchesToHeight(a.value);
        result.textContent = output || "Entrada invalida";
      } catch (error) {
        result.textContent = error.message;
      }
    }

    openBtn?.addEventListener("click", () => panel.classList.remove("hidden"));
    closeBtn?.addEventListener("click", () => panel.classList.add("hidden"));
    runBtn?.addEventListener("click", runCalculation);
    [a, b].forEach(input => {
      input.addEventListener("keydown", event => {
        if (event.key === "Enter") {
          event.preventDefault();
          runCalculation();
        }
      });
    });
    op.addEventListener("keydown", event => {
      if (event.key === "Enter") {
        event.preventDefault();
        runCalculation();
      }
    });
  }

  global.FloatingCalculator = { setupFloatingCalculator };
})(window);
