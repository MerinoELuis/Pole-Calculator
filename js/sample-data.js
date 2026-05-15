(function (global) {
  "use strict";

  const S = () => global.AppStore;

  function loadVexDemoData() {
    const store = S();
    const state = store.resetState();

    state.importedFileName = "Datos demo · VEX-AZPR022-7";
    state.importedAt = new Date().toISOString();
    state.selectedPoleId = "P79-495806";
    state.autoCreateSpanComms = false;

    [
      store.createPole("P77-495804 STEEL", "40'", "", { lowPower: "29'6\"", sequence: "P77" }),
      store.createPole("P78-495805", "30'", "", { lowPower: "23'10\"", sequence: "P78" }),
      store.createPole("P79-495806", "40'", "", { lowPower: "32'4\"", sequence: "P79" }),
      store.createPole("P80-506854", "35'", "", { lowPower: "29'5\"", sequence: "P80" }),
      store.createPole("P81-495807", "35'", "", { lowPower: "26'1\"", sequence: "P81" }),
      store.createPole("P82-X495808", "35'", "", { lowPower: "", sequence: "P82", isGenerated: true })
    ].forEach(store.upsertPole);

    [
      store.createSpan("cBl3NJhIH3", "P77-495804 STEEL", "P79-495806", "E", "Other span P77 → P79", {
        spanIndex: "11", lengthDisplay: "102'11\"", environment: "RESIDENTIAL_DRIVEWAY", environmentClearance: "15'6\"", bearingDegrees: 103.49
      }),
      store.createSpan("sbTP2GvSMo", "P79-495806", "P77-495804 STEEL", "W", "Back span P79 → P77", {
        spanIndex: "19", lengthDisplay: "102'11\"", environment: "RESIDENTIAL_DRIVEWAY", environmentClearance: "15'6\"", bearingDegrees: 283.49
      }),
      store.createSpan("tzcrik4dzM", "P79-495806", "P80-506854", "E", "Fore span P79 → P80", {
        spanIndex: "18", lengthDisplay: "184'6\"", environment: "RESIDENTIAL_DRIVEWAY", environmentClearance: "15'6\"", bearingDegrees: 105.31
      }),
      store.createSpan("KGstPoeO4K", "P80-506854", "P79-495806", "W", "Back span P80 → P79", {
        spanIndex: "20", lengthDisplay: "184'6\"", environment: "RESIDENTIAL_DRIVEWAY", environmentClearance: "15'6\"", bearingDegrees: 285.31
      }),
      store.createSpan("vXXhjs3xKo", "P79-495806", "P81-495807", "NE", "Other span P79 → P81", {
        spanIndex: "15", lengthDisplay: "125'11\"", environment: "STREET", environmentClearance: "15'6\"", bearingDegrees: 27.45
      }),
      store.createSpan("UUFGpRqK45", "P81-495807", "P79-495806", "SW", "Back span P81 → P79", {
        spanIndex: "22", lengthDisplay: "125'11\"", environment: "STREET", environmentClearance: "15'6\"", bearingDegrees: 207.45
      })
    ].forEach(store.upsertSpan);

    [
      ["P77-495804 STEEL", "COMMUNICATION > Cable One, Prescott", "22'2\"", "CATV", "Cable One, Prescott"],
      ["P77-495804 STEEL", "COMMUNICATION > Century Link Communications", "21'2\"", "CTL", "Century Link Communications"],
      ["P79-495806", "COMMUNICATION > Cable One, Prescott", "26'", "CATV", "Cable One, Prescott"],
      ["P79-495806", "COMMUNICATION > Century Link Communications", "25'", "CTL", "Century Link Communications"],
      ["P80-506854", "COMMUNICATION > Cable One, Prescott", "21'4\"", "CATV", "Cable One, Prescott"],
      ["P80-506854", "COMMUNICATION > Century Link Communications", "20'4\"", "CTL", "Century Link Communications"],
      ["P81-495807", "COMMUNICATION > Cable One, Prescott", "20'9\"", "CATV", "Cable One, Prescott"],
      ["P81-495807", "COMMUNICATION > Century Link Communications", "20'9\"", "CTL", "Century Link Communications"]
    ].forEach(([poleId, owner, hoa, ownerBase, rawOwner]) => {
      store.upsertComm(poleId, owner, hoa, "", { ownerBase, rawOwner });
    });

    const comms = [
      {
        spanId: "tzcrik4dzM", poleId: "P79-495806", owner: "COMMUNICATION > Cable One, Prescott", ownerBase: "CATV",
        existingHOA: "26'", existingHOAChange: "", midspan: "15'7\"", wireId: "QxPpNpghUs", wireIndex: "54",
        rawOwner: "COMMUNICATION > Cable One, Prescott", size: "CATV Bundles > 1\" Communication Bundle  Msgr:0.242\" > Tension to Sag", construction: "ON_POLE", insulator: "Three Bolt"
      },
      {
        spanId: "tzcrik4dzM", poleId: "P79-495806", owner: "COMMUNICATION > Century Link Communications", ownerBase: "CTL",
        existingHOA: "25'", existingHOAChange: "24'6\"", midspan: "14'11\"", wireId: "PIRZvCtj7d", wireIndex: "55",
        rawOwner: "COMMUNICATION > Century Link Communications", size: "Telco Bundles > 1\" Communication Bundle  Msgr:0.242\" > Tension to Sag", construction: "ON_POLE", insulator: "Three Bolt"
      },
      {
        spanId: "KGstPoeO4K", poleId: "P80-506854", owner: "COMMUNICATION > Cable One, Prescott", ownerBase: "CATV",
        existingHOA: "21'4\"", existingHOAChange: "", midspan: "", wireId: "QxPpNpghUs", wireIndex: "62",
        rawOwner: "COMMUNICATION > Cable One, Prescott", size: "CATV Bundles > 1\" Communication Bundle  Msgr:0.242\" > Tension to Sag", construction: "ON_POLE", insulator: "Single Bolt"
      },
      {
        spanId: "KGstPoeO4K", poleId: "P80-506854", owner: "COMMUNICATION > Century Link Communications", ownerBase: "CTL",
        existingHOA: "20'4\"", existingHOAChange: "", midspan: "", wireId: "PIRZvCtj7d", wireIndex: "63",
        rawOwner: "COMMUNICATION > Century Link Communications", size: "Telco Bundles > 1\" Communication Bundle  Msgr:0.242\" > Tension to Sag", construction: "ON_POLE", insulator: "Single Bolt"
      },
      {
        spanId: "vXXhjs3xKo", poleId: "P79-495806", owner: "COMMUNICATION > Century Link Communications", ownerBase: "CTL",
        existingHOA: "25'", existingHOAChange: "", midspan: "17'3\"", wireId: "yINRlKJaqn", wireIndex: "44",
        rawOwner: "COMMUNICATION > Century Link Communications", size: "Communication Drops > 0.25in Telco > Tension to Sag", construction: "ON_POLE", insulator: "J-Hook", serviceDrop: true
      },
      {
        spanId: "UUFGpRqK45", poleId: "P81-495807", owner: "COMMUNICATION > Century Link Communications", ownerBase: "CTL",
        existingHOA: "20'9\"", existingHOAChange: "", midspan: "", wireId: "yINRlKJaqn", wireIndex: "65",
        rawOwner: "COMMUNICATION > Century Link Communications", size: "Communication Drops > 0.25in Telco > Tension to Sag", construction: "ON_POLE", insulator: "J-Hook", serviceDrop: true
      },
      {
        spanId: "cBl3NJhIH3", poleId: "P77-495804 STEEL", owner: "COMMUNICATION > Cable One, Prescott", ownerBase: "CATV",
        existingHOA: "22'2\"", existingHOAChange: "", midspan: "21'4\"", wireId: "wW8J7vOgXz", wireIndex: "32",
        rawOwner: "COMMUNICATION > Cable One, Prescott", size: "CATV Bundles > 1\" Communication Bundle  Msgr:0.242\" > Tension to Sag", construction: "ON_POLE", insulator: "Three Bolt"
      },
      {
        spanId: "cBl3NJhIH3", poleId: "P77-495804 STEEL", owner: "COMMUNICATION > Century Link Communications", ownerBase: "CTL",
        existingHOA: "21'2\"", existingHOAChange: "", midspan: "19'7\"", wireId: "Ro1xvbJPqo", wireIndex: "33",
        rawOwner: "COMMUNICATION > Century Link Communications", size: "Telco Bundles > 1\" Communication Bundle  Msgr:0.242\" > Tension to Sag", construction: "ON_POLE", insulator: "Single Bolt"
      }
    ];

    comms.forEach(data => store.upsertSpanComm(data));

    [
      { spanId: "tzcrik4dzM", poleId: "P79-495806", label: "Secondary", attachmentHeight: "32'4\"", midspan: "21'3\"", size: "Secondary > A1/0A", owner: "UTILITY > APS", wireId: "UVsaUkqjDY" },
      { spanId: "KGstPoeO4K", poleId: "P80-506854", label: "Secondary", attachmentHeight: "29'5\"", midspan: "21'3\"", size: "Secondary > A1/0A", owner: "UTILITY > APS", wireId: "UVsaUkqjDY" },
      { spanId: "vXXhjs3xKo", poleId: "P79-495806", label: "Power", attachmentHeight: "32'4\"", midspan: "24'4\"", size: "Secondary > A1/0A", owner: "UTILITY > APS", wireId: "SLtEgiwJ0j" },
      { spanId: "cBl3NJhIH3", poleId: "P77-495804 STEEL", label: "Secondary", attachmentHeight: "29'4\"", midspan: "26'6\"", size: "Secondary > A1/0A", owner: "UTILITY > APS", wireId: "DOcqFC9oYC" }
    ].forEach(store.addSpanPower);

    return store.normalizeState(state);
  }

  if (S()) {
    S().loadSampleData = loadVexDemoData;
  }
})(window);
