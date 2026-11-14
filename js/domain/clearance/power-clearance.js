(function (global) {
  "use strict";

  function create(context = {}) {
    const { S, H, isMidAmProfile, isComedProfile, isCsuProfile, isPrimaryPowerRow } = context;

    function validEquipmentActionHeightInches(equipment) {
      if (!equipment?.actionActive) return null;
      const target = H().parseHeight(equipment.actionHeight || "");
      if (target === null) return null;
      const category = String(equipment.category || equipment.type || "").toUpperCase();
      if (category.includes("RISER")) {
        const source = H().parseHeight(equipment.attachmentHeight || "");
        if (source === null || target <= source) return null;
      }
      return target;
    }

    // INTEC and Metronet MidAm allow a Streetlight to be raised independently
    // of grounding, but never by more than one foot from the imported HOA.
    function validStreetlightRaiseInches(equipment) {
      const project = String(S().getState().settings?.projectProfile || "").toUpperCase();
      if ((!isMidAmProfile() && project !== "INTEC") || !equipment?.raiseActive) return null;
      const source = H().parseHeight(equipment.attachmentHeight || "");
      const target = H().parseHeight(equipment.raiseHeight || "");
      if (source === null || target === null || target <= source || target > source + 12) return null;
      return target;
    }

    // Returns the comm-height ceiling imposed by one imported Power Equipment
    // row. Streetlight bracket and drip-loop clearances are evaluated separately
    // because grounding changes only the INTEC bracket rule, never the drip loop.
    function powerEquipmentCeilingInches(equipment) {
      if (!equipment) return null;
      const settings = S().getState().settings || {};
      const genericClearance = H().parseHeight(settings.polePowerCommsClearance || settings.clearanceToPower || "40\"");
      const category = String(equipment.category || equipment.type || "").toUpperCase();
      const actionHeight = validEquipmentActionHeightInches(equipment);
      let attachment = category.includes("RISER") && actionHeight !== null
        ? actionHeight
        : H().parseHeight(equipment.attachmentHeight || "");
      let bottom = H().parseHeight(equipment.bottomHeight || "");
      let dripLoop = category.includes("TRANSFORMER") && actionHeight !== null
        ? actionHeight
        : H().parseHeight(equipment.dripLoopHeight || "");

      if (category.includes("STREETLIGHT")) {
        if (isMidAmProfile()) {
          // MidAm Streetlight rows use only the imported Attachment Height. The
          // workbook's Bottom/Drip Loop fields are invalid for this WI and are
          // reported by Excel Review rather than used in pole clearance math.
          const midAmAttachment = validStreetlightRaiseInches(equipment)
            ?? H().parseHeight(equipment.attachmentHeight || "");
          const midAmClearance = H().parseHeight(settings.streetlightBracketCommClearance || "20\"");
          return midAmAttachment !== null && midAmClearance !== null
            ? midAmAttachment - midAmClearance
            : null;
        }
        const raisedAttachment = validStreetlightRaiseInches(equipment);
        if (raisedAttachment !== null && attachment !== null) {
          // The complete Streetlight assembly moves by the same delta, so its
          // bracket and drip-loop clearance references remain physically aligned.
          const delta = raisedAttachment - attachment;
          attachment = raisedAttachment;
          if (bottom !== null) bottom += delta;
          if (dripLoop !== null) dripLoop += delta;
        }
        const specialCeilings = [];
        const bracket = bottom !== null ? bottom : attachment;
        const bracketClearance = isMidAmProfile()
          ? H().parseHeight(settings.streetlightBracketCommClearance || "")
          : isComedProfile()
            ? H().parseHeight(settings.streetlightBracketCommClearance || "4\"")
          : isCsuProfile()
            ? H().parseHeight(equipment.actionActive
              ? (settings.streetlightBracketCommClearance || "16\"")
              : (settings.streetlightUngroundedBracketCommClearance || "52\""))
            : H().parseHeight(equipment.actionActive
              ? (settings.streetlightGroundedBracketCommClearance || "12\"")
              : (settings.streetlightBracketCommClearance || settings.polePowerCommsClearance || settings.clearanceToPower || "40\""));
        const dripLoopClearance = isMidAmProfile()
          ? H().parseHeight(settings.streetlightDripLoopCommClearance || "")
          : isComedProfile()
            ? H().parseHeight(settings.streetlightDripLoopCommClearance || "12\"")
          : isCsuProfile()
            ? H().parseHeight(settings.streetlightDripLoopCommClearance || "24\"")
          : genericClearance;
        if (bracket !== null && bracketClearance !== null) specialCeilings.push(bracket - bracketClearance);
        if (dripLoop !== null && dripLoopClearance !== null) specialCeilings.push(dripLoop - dripLoopClearance);
        if (specialCeilings.length) return Math.min(...specialCeilings);
      }

      if (category.includes("TRANSFORMER")) {
        // Transformer Bottom Height is the primary physical reference. When the
        // workbook omits it, use the secured/current drip loop, then attachment.
        const transformerHeights = [bottom, dripLoop].filter(value => value !== null);
        const transformerBottom = transformerHeights.length ? Math.min(...transformerHeights) : attachment;
        const transformerClearance = isCsuProfile()
          ? H().parseHeight(settings.transformerCommClearance || "42\"")
          : genericClearance;
        return transformerBottom !== null && transformerClearance !== null
          ? transformerBottom - transformerClearance
          : null;
      }

      const physicalHeights = [attachment, bottom, dripLoop].filter(value => value !== null);
      return physicalHeights.length && genericClearance !== null
        ? Math.min(...physicalHeights) - genericClearance
        : null;
    }

    // The source height that can make an equipment row the pole's imported
    // Low Power. Transformer work moves its drip loop; riser work moves the
    // riser attachment. Streetlight grounding has no vertical movement.
    function equipmentLowPowerSourceInches(equipment) {
      const category = String(equipment?.category || equipment?.type || "").toUpperCase();
      if (category.includes("TRANSFORMER")) {
        return H().parseHeight(equipment.dripLoopHeight || equipment.bottomHeight || equipment.attachmentHeight || "");
      }
      if (category.includes("RISER")) return H().parseHeight(equipment.attachmentHeight || "");
      return null;
    }

    // A raised riser becomes the new Low Power reference eight inches below its
    // attachment HOA. The offset represents the riser/drip-loop clearance.
    function equipmentLowPowerTargetInches(item) {
      const target = item?.target;
      if (target === null || target === undefined) return null;
      const category = String(item.row?.category || item.row?.type || "").toUpperCase();
      return category.includes("RISER") ? target - 8 : target;
    }

    /**
     * Applies active Transformer/Riser work to the imported Low Power baseline.
     * A raised item replaces the baseline only when every equipment item at that
     * same limiting height was moved. A lowered target can always become the new
     * lowest Power point.
     */
    function effectivePoleLowPowerInches(pole, equipment) {
      const metadata = pole?.metadata || {};
      const baselineText = Object.prototype.hasOwnProperty.call(metadata, "lowPowerBaseline")
        ? metadata.lowPowerBaseline
        : pole?.lowPower;
      const baseline = H().parseHeight(baselineText || "");
      const movable = (equipment || []).map(row => ({
        row,
        source: equipmentLowPowerSourceInches(row),
        target: validEquipmentActionHeightInches(row)
      }));
      const powerMovable = S().getSpanPowerForPole(pole?.poleId || "")
        .filter(row => !isPrimaryPowerRow(row))
        .map(row => ({
          row,
          source: H().parseHeight(row.attachmentHeight || ""),
          target: H().parseHeight(row.attachmentHeightChange || "")
        }));
      let effective = baseline;

      if (baseline !== null) {
        const limitingEquipment = movable.filter(item => item.source === baseline);
        const limitingPower = powerMovable.filter(item => item.source === baseline);
        const limitingTargets = [
          ...limitingEquipment.map(equipmentLowPowerTargetInches),
          ...limitingPower.map(item => item.target)
        ];
        if (limitingTargets.length && limitingTargets.every(value => value !== null)) {
          effective = Math.min(...limitingTargets);
        }
      }

      movable.forEach(item => {
        const lowPowerTarget = equipmentLowPowerTargetInches(item);
        if (lowPowerTarget !== null && (effective === null || lowPowerTarget < effective)) effective = lowPowerTarget;
      });
      powerMovable.forEach(item => {
        if (item.target !== null && (effective === null || item.target < effective)) effective = item.target;
      });
      return effective;
    }

    return {
      validEquipmentActionHeightInches,
      validStreetlightRaiseInches,
      powerEquipmentCeilingInches,
      equipmentLowPowerSourceInches,
      equipmentLowPowerTargetInches,
      effectivePoleLowPowerInches
    };
  }

  global.PowerClearanceModule = { create };
})(window);
