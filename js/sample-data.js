(function (global) {
  "use strict";

  const S = () => global.AppStore;

  // Demo basado en EXCEL_VEX-AZPR022-7_2026-05-11.xlsx.
  // Este caso conserva un forespan con midspan real (P74 -> P75) y el
  // backspan de referencia (P75 -> P74) con los mismos Wire Ids. Sirve para
  // probar que mover el comm del segundo poste recalcula el midspan del primero.
  function loadVexDemoData() {
    const store = S();
    const state = store.resetState();

    state.importedFileName = "Datos demo · VEX-AZPR022-7";
    state.importedAt = new Date().toISOString();
    state.selectedPoleId = "P75-506851 STEEL UG";
    state.autoCreateSpanComms = false;

    [
      store.createPole("P74-494811 UG", "30'", "", { lowPower: "23'", sequence: "P74" }),
      store.createPole("P75-506851 STEEL UG", "35'", "", { lowPower: "26'10\"", sequence: "P75" }),
      store.createPole({ poleId: "Unknown-DYWnFiLgq2", isGenerated: true }),
      store.createPole({ poleId: "Unknown-5v7U5qxesL", isGenerated: true })
    ].forEach(store.upsertPole);

    [
      store.createSpan("JAdOOUWQij", "P74-494811 UG", "P75-506851 STEEL UG", "E", "Fore span P74 -> P75", {
        spanIndex: "4", lengthDisplay: "125'6\"", environment: "STREET", environmentClearance: "15'6\"", bearingDegrees: 95.32
      }),
      store.createSpan("rlxoA23G99", "P75-506851 STEEL UG", "P74-494811 UG", "W", "Back span P75 -> P74", {
        spanIndex: "14", lengthDisplay: "125'6\"", environment: "STREET", environmentClearance: "15'6\"", bearingDegrees: 275.32
      }),
      store.createSpan("DYWnFiLgq2", "P75-506851 STEEL UG", "Unknown-DYWnFiLgq2", "E", "Fore span P75 -> unknown", {
        spanIndex: "5", lengthDisplay: "90'", environment: "STREET", environmentClearance: "15'6\"", isGeneratedOtherPole: true
      }),
      store.createSpan("5v7U5qxesL", "P75-506851 STEEL UG", "Unknown-5v7U5qxesL", "E", "Other span P75 -> unknown", {
        spanIndex: "6", lengthDisplay: "80'", environment: "STREET", environmentClearance: "15'6\"", isGeneratedOtherPole: true
      })
    ].forEach(store.upsertSpan);

    [
      ["P74-494811 UG", "COMMUNICATION > Cable One, Prescott", "19'2\"", "CATV", "COMMUNICATION > Cable One, Prescott"],
      ["P74-494811 UG", "COMMUNICATION > Century Link Communications", "18'10\"", "CTL", "COMMUNICATION > Century Link Communications"],
      ["P75-506851 STEEL UG", "COMMUNICATION > Cable One, Prescott", "23'4\"", "CATV", "COMMUNICATION > Cable One, Prescott"],
      ["P75-506851 STEEL UG", "COMMUNICATION > Century Link Communications", "22'4\"", "CTL", "COMMUNICATION > Century Link Communications"]
    ].forEach(([poleId, owner, hoa, ownerBase, rawOwner]) => {
      store.upsertComm(poleId, owner, hoa, "", { ownerBase, rawOwner });
    });

    [
      {
        spanId: "JAdOOUWQij", poleId: "P74-494811 UG", owner: "COMMUNICATION > Cable One, Prescott", ownerBase: "CATV",
        existingHOA: "19'2\"", midspan: "15'5\"", wireId: "e58fMq8UGK", wireIndex: "3",
        rawOwner: "COMMUNICATION > Cable One, Prescott", size: "CATV Bundles > 1\" Communication Bundle", construction: "ON_POLE", insulator: "Single Bolt"
      },
      {
        spanId: "JAdOOUWQij", poleId: "P74-494811 UG", owner: "COMMUNICATION > Century Link Communications", ownerBase: "CTL",
        existingHOA: "18'10\"", midspan: "14'", wireId: "pu6TJVzuSi", wireIndex: "4",
        rawOwner: "COMMUNICATION > Century Link Communications", size: "Communication Drops > 0.25in Telco", construction: "ON_POLE", insulator: "J-Hook", serviceDrop: true
      },
      {
        spanId: "rlxoA23G99", poleId: "P75-506851 STEEL UG", owner: "COMMUNICATION > Cable One, Prescott", ownerBase: "CATV",
        existingHOA: "23'4\"", midspan: "", wireId: "e58fMq8UGK", wireIndex: "3",
        rawOwner: "COMMUNICATION > Cable One, Prescott", size: "CATV Bundles > 1\" Communication Bundle", construction: "ON_POLE", insulator: "Single Bolt"
      },
      {
        spanId: "rlxoA23G99", poleId: "P75-506851 STEEL UG", owner: "COMMUNICATION > Century Link Communications", ownerBase: "CTL",
        existingHOA: "22'4\"", midspan: "", wireId: "pu6TJVzuSi", wireIndex: "4",
        rawOwner: "COMMUNICATION > Century Link Communications", size: "Communication Drops > 0.25in Telco", construction: "ON_POLE", insulator: "J-Hook", serviceDrop: true
      },
      {
        spanId: "DYWnFiLgq2", poleId: "P75-506851 STEEL UG", owner: "COMMUNICATION > Cable One, Prescott", ownerBase: "CATV",
        existingHOA: "23'4\"", midspan: "22'", wireId: "cml45oGkQr", wireIndex: "2",
        rawOwner: "COMMUNICATION > Cable One, Prescott", size: "CATV Bundles > 1\" Communication Bundle", construction: "ON_POLE", insulator: "Single Bolt"
      },
      {
        spanId: "5v7U5qxesL", poleId: "P75-506851 STEEL UG", owner: "COMMUNICATION > Century Link Communications", ownerBase: "CTL",
        existingHOA: "22'4\"", midspan: "21'4\"", wireId: "lpuJi2oP4x", wireIndex: "3",
        rawOwner: "COMMUNICATION > Century Link Communications", size: "Communication Drops > 0.25in Telco", construction: "ON_POLE", insulator: "J-Hook", serviceDrop: true
      }
    ].forEach(store.upsertSpanComm);

    [
      { spanId: "JAdOOUWQij", poleId: "P74-494811 UG", label: "Secondary", attachmentHeight: "24'4\"", midspan: "16'11\"", size: "Secondary", owner: "UTILITY > APS", wireId: "76DLW7yzso" },
      { spanId: "rlxoA23G99", poleId: "P75-506851 STEEL UG", label: "Secondary", attachmentHeight: "27'4\"", midspan: "", size: "Secondary", owner: "UTILITY > APS", wireId: "76DLW7yzso" },
      { spanId: "DYWnFiLgq2", poleId: "P75-506851 STEEL UG", label: "Secondary", attachmentHeight: "27'4\"", midspan: "24'6\"", size: "Secondary", owner: "UTILITY > APS", wireId: "dEcYp8lVzw" }
    ].forEach(store.addSpanPower);

    return store.normalizeState(state);
  }

  if (S()) {
    S().loadSampleData = loadVexDemoData;
  }
})(window);
