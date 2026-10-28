(function (global) {
  "use strict";

  const text = value => String(value ?? "").trim();

  function canonicalPoleIdentity(value) {
    const parts = text(value).replace(/\s+/g, " ").split(" ").filter(Boolean);
    while (parts.length > 1 && /^(STEEL|UG|PCO)$/i.test(parts[parts.length - 1])) parts.pop();
    return parts.join(" ").toUpperCase();
  }

  function isUnknownPoleId(value) {
    return /^unknown(?:-|\b)/i.test(text(value));
  }

  function isAddressableUnknownPoleId(value) {
    const id = text(value);
    return isUnknownPoleId(id) || /-unknown(?:-|\b)/i.test(id);
  }

  function keyForSpanSide(spanId, poleId) {
    return [spanId, poleId].map(text).join("__");
  }

  function keyForSpanComm(spanId, poleId, owner, wireId) {
    return [spanId, poleId, owner, wireId || ""].map(text).join("__");
  }

  global.AppStoreIdentity = Object.freeze({
    canonicalPoleIdentity,
    isUnknownPoleId,
    isAddressableUnknownPoleId,
    keyForSpanSide,
    keyForSpanComm
  });
})(window);
