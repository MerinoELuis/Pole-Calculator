# Business Rules

This document records the rules currently implemented in `calculations.js` and `mr-logic.js`. Editable values always come from `AppState.settings`.

## Height Arithmetic

All heights are parsed to integer inches before calculations. Decimal feet are rounded to the nearest inch. Final half-movement results are also rounded with `Math.round()` before display.

## Derived Pole Limits

```text
Max Height on Pole = Low Power on Pole - Pole Power-comms clearance
```

For Metronet/MidAm, the final maximum is the lowest available ceiling:

```text
Low Power ceiling        = Low Power - 40"
Streetlight bracket      = Bottom Height - 20"
Uncovered drip-loop      = Drip Loop Height - 12"
Max Height on Pole       = minimum available ceiling
```

MidAm comms and Proposed must also remain at least `3\"` from imported MidAm utility guy attachment heights.

The highest and lowest effective comm heights use `HOA Change` when present; otherwise they use `Existing HOA`. INTEC Self-Supporting Fiber is POF: it stays visible but does not define Top Comm or Low Comm for Proposed placement.

## Derived Midspan Power Limit

The calculator selects the lowest valid power midspan on a span.

```text
Low Power at MS = minimum imported/edited power midspan on the span
Max Height at MS = Low Power at MS - Midspan Power-comm clearance
```

If no power midspan exists, `Max Height at MS` remains blank and the status reports missing power data rather than inventing a limit.

## Pole Type Check Severity

Pole height issues are critical when the imported row has a missing Tip, a calculated height without a reference-table row, or a calculated height that differs from the imported pole type. Class-only differences remain review warnings. Critical height issues appear on the `Pole Type Check` tab and beside the affected pole in the calculator index; they are also included by the `Flagged only` filter.

## Existing Comm Midspan

The baseline belongs to the current SpanComm row:

1. Use its imported `midspan`.
2. Use its `ocalcMS` only as a compatibility fallback.
3. Never borrow a baseline from an unrelated span or another pole.

The movement formula is:

```text
local movement  = effective local HOA - local Existing HOA
remote movement = effective remote HOA - remote Existing HOA

calculated midspan = imported midspan
                   + local movement / 2
                   + remote movement / 2
```

Example:

- Imported midspan: `18'`
- Local movement: `20'` to `19'` = `-12"`, contributes `-6"`
- Remote movement: `20'` to `20'8"` = `+8"`, contributes `+4"`
- Result: `17'10"`

## Remote Comm Selection

The opposite endpoint is found from the current span. Candidates at that pole are ranked using:

1. Exact `wireId` match.
2. Exact `spanId` match.
3. Same physical unordered pole pair.
4. Normalized owner match.
5. Presence of an imported midspan as a small tie breaker.

Owner normalization recognizes common equivalents such as CenturyLink/CTL/Telco and CATV/Cable One/Cox. Exact Wire ID has the highest priority because one owner can have several cables.

## Midspan Display and Validation Source

For a non-REF comm, the UI and Comm-comm MS validation select the first valid value in this order:

1. Fresh `calculatedMidspan` from the current recalculation.
2. Stored `finalMidspan`.
3. Stored `msProposed`.
4. Stored `calculatedMidspan`.
5. Imported `midspan`.
6. Legacy `ocalcMS`.

Using the same priority prevents the table from displaying one value while flagging another stale value.

## REF Rules

| Span type | Own midspan | Behavior |
| --- | --- | --- |
| Back Span, INTEC | missing or present | REF for current calculations. An imported value can remain visible for inspection. |
| Back Span, MidAm | missing | REF with no calculated midspan. |
| Back Span, MidAm | present | Remains labeled REF, but its own imported baseline is recalculated from both endpoint movements. |
| Other | missing | REF. No calculated midspan or Other Pole HOA. |
| Other | present | Real midspan row; calculations and flagging apply. |
| Fore Span | missing | Editable so the user can create the midspan. |
| Fore Span | present | Real midspan row; calculations and flagging apply. |

For `Proposed by Span`, a valid imported Power midspan also counts as real
midspan data. This allows a Fore Span to be proposed when it contains Power
clearance data but no communication rows in `Span.Wire`.

## Comm Flagging

One compact comm flagging field combines these checks:

1. Midspan is not below Environment clearance.
2. Midspan does not exceed Max Height at MS.
3. Different comms on the same `spanId` keep Midspan Comm-comm clearance.
4. Comm order at the pole and midspan does not reverse, preventing cable crossings.
5. Top Comm mode leaves room above the highest comm midspan for Proposed.
6. Effective HOA does not exceed Max Height on Pole.
7. Different owners keep Pole Comm-comm clearance.
8. Same owners keep Pole Bolt-bolt clearance.
9. New movement bolts keep Bolt-bolt clearance from previous Existing HOA points.

MidAm environment defaults are profile-specific: railroad `23'6\"`; truck traffic, parking lots, alleys, farms and along-road spans `15'6\"`; pedestrian-only areas `9'6\"`; and water without sailboats `14'`.

## Power Equipment and Pole Maximum

The `Equipment` sheet contributes only rows owned by Utility/Power and categorized as Streetlight, Transformer, or Riser. Communication risers do not participate.

1. INTEC keeps `40"` below Low Power, Transformer Bottom Height, Streetlight drip loops, and ungrounded Streetlight brackets.
2. An INTEC grounded Streetlight bracket keeps `12"` to the top comm. Grounding changes only the bracket rule; its drip loop still keeps `40"`.
3. MidAm Streetlights use `Bottom Height - Streetlight bracket-comm` and `Drip Loop Height - Streetlight drip loop-comm`; the lower result controls.
4. MidAm Transformers and Power Risers subtract Pole Power-comms clearance from their lowest imported physical height.
5. `Max Height on Pole` is the lowest result from Low Power and every applicable Power Equipment row.

### Power Equipment Actions

1. Every MidAm Streetlight has mandatory `Ground`: import and recalculation activate it automatically, the UI cannot disable it, and MR adds `MNT GROUND STREETLIGHT`.
2. Grounding does not invent a vertical height or remove physical clearances. MidAm still applies `Bottom Height - 20\"` and `Drip Loop Height - 12\"`; the lower ceiling controls.
3. Transformer `Redress` requires New HOA. INTEC adds `Secure transformer drip loop to HOA <height>.`; Metronet adds `POWER REDRESS TRANSFORMER DRIP LOOP TO HOA <height>.`
4. Streetlight `Ground` adds `Install flex conduit to STLT circuit. bond STLT housing to pole GRND/NEUT.` for INTEC and `MNT GROUND STREETLIGHT` for Metronet.
5. Power Riser `Raise` requires a New HOA above the imported attachment. INTEC adds `Raise APS riser from HOA <old> to HOA <new>.`; Metronet adds `AT HOA <old> RAISE POWER RISER TO HOA <new> DUE TO CLEARANCES.`
6. INTEC Streetlight `Raise` is independent of `Ground`. New HOA must be above Attachment Height and no more than `12"` higher. The bracket, bottom, and drip-loop references move by the same delta, and MR adds `Raise streetlight from HOA <old> to <new>.`
5. A valid Transformer/Riser target replaces that equipment height in its pole-clearance calculation.
6. When the moved equipment supplied the imported Low Power, the target becomes effective Low Power. If several equipment rows share that limiting height, all must be moved before a higher Low Power can replace it.
7. Disabling an optional action restores calculation from `metadata.lowPowerBaseline`. Updating Excel refreshes that baseline but preserves matching user actions.

## MidAm Collection Identity

1. Sequence contains three digits and may end with one letter.
2. Numeric Excel values are left-padded before validation, so `58` is treated as `058`.
3. Letter suffixes are case-insensitive and normalized to uppercase, so `51b` is treated as `051B`.
4. The first block of Collection `Id` must independently use the same `000` or `000A` format.
5. The normalized Sequence and first Id block must be exactly equal; `058` does not match `058A`.
6. Collection Owner must be `UTILITY > MidAm`; empty is an error and another populated owner is a warning.

A same-owner Service Drop may reuse exactly the same bolt height. This exception applies only when the separation is zero; nearby nonzero placements still use Bolt-bolt clearance.

When `Transfer to New Pole` is active, the old `Existing HOA` bolt belongs to the prior pole and is excluded from Bolt-bolt checks on the destination pole. The transferred comm's effective HOA remains subject to Comm-comm and every other applicable clearance.

### Transfer and Re-sag Controls

`Transfer to New Pole` is available for every project profile and is controlled once per physical comm at a pole. The value is synchronized across that comm's span relationships for persistence, but it generates only one Make Ready instruction. When the comm has an HOA movement, it replaces the normal raise/lower or relocate instruction with:

```text
Transfer <owner> to new pole at HOA <HOA Change> [with DG].
```

`Re-sag Service Drop` is available only for INTEC and only applies to a row marked `Service Drop`. If its original calculated midspan is below `15'6"`, the effective midspan used by display and validation becomes `15'6"` without overwriting the imported baseline. Other Power, Comm-comm, and crossing checks still apply at the adjusted height. The generated Make Ready line is:

```text
Re-sag <owner> comm drop <direction>, ensure 15'6" at midspan.
```

Comm-comm MS comparisons are restricted to the same `spanId`. The message includes the physical span label, both owners, both midspans, calculated separation, and required minimum.

## Proposed on the Pole

Proposed Flagging checks:

- Proposed must not exceed Max Height on Pole.
- Top Comm mode cannot place Proposed below Top Comm.
- Low Comm mode cannot place Proposed above Low Comm.
- Proposed keeps Pole Comm-comm clearance from effective comm heights.
- Proposed keeps Pole Bolt-bolt clearance from every Existing HOA point.
- Proposed attachments on the same pole keep Bolt-bolt clearance from one another.
- An exact occupied HOA is allowed only after the existing comm at that height has moved away.

## Proposed Midspan

The Proposed MS base is `O-CALC MS`, with imported `proposedMidspan` as fallback. The value is evaluated only against comm midspans from the same physical pole-to-pole connection. Unrelated spans on the same pole are excluded.

For INTEC/Wecom, both fields remaining blank means MS Proposed and Adjusted Final MS remain blank until the user enters `O-CALC MS`.

For Metronet/MidAm, when both fields are blank the base is calculated automatically:

1. If the physical span contains one or more comm midspans, use the highest comm midspan plus `12\"`. MidAm checks both directed rows of the same pole pair because the measured value may live on the reciprocal Back/Other span.
2. Otherwise, subtract estimated sag from Proposed HOA.
3. Round span length to the nearest `50 ft` before estimating sag at `1 ft` per `100 ft`: `100 ft -> 12\"`, `150 ft -> 18\"`, `200 ft -> 24\"`, and `250 ft -> 30\"`.
4. Existing MS adjustment and flagging rules run after this base is obtained.

In Top Comm mode:

```text
minimum Proposed MS = highest applicable comm MS + Midspan Comm-comm clearance
```

In Low Comm mode:

```text
maximum Proposed MS = lowest applicable comm MS - Midspan Comm-comm clearance
```

The Environment minimum is then applied. Finally, Max Height at MS is checked.

- INTEC can cap Adjusted Final MS to Max Height at MS and generates the low-power Make Ready reminder when applicable.
- Metronet does not silently cap this value. The violation stays visible and must be solved by lowering comms.

`MS Flagging` reports the Proposed midspan adjustment/result. `Proposed Flagging` reports pole-side Proposed placement rules.

## End Drop and Next Pole Proposed

```text
End Drop = Next Pole Proposed - local Proposed
```

Next Pole Proposed selection:

1. Keep a manually edited value when `nextPoleProposedAuto` is false.
2. Otherwise find the first non-additional forward Proposed owned by the connected pole.
3. If none exists, use that pole's `standaloneProposedHOA`.
4. If no value exists, leave Next Pole Proposed and End Drop blank.

Editing Next Pole Proposed locally does not change the Proposed value on the connected pole.

## Multiple Proposed Attachments

The first proposal on an imported connection reuses the physical span. Additional proposals receive separate state while retaining `sourceSpanId` so geometry, Environment, and power limits come from the same physical span.

Only the first non-additional Proposed at the next pole is used automatically for the preceding pole's End Drop.

## Auto Calculate

Auto Calculate is enabled only in Top Comm mode.

For every pass it:

1. Recalculates the full state.
2. Visits poles and builds comm groups ordered by Existing HOA.
3. Tests candidate Proposed heights against a clean state snapshot.
4. Builds a downward comm stack when Proposed or midspan limits require room.
5. Uses Comm-comm spacing between different owners and Bolt-bolt spacing between the same owner.
6. Recalculates affected poles and rejects candidates that introduce a new violation.
7. Repeats until the movement signature no longer changes.

The solver stops at `max(8, pole count * 2 + 4)` passes or when a repeated signature is detected. Existing user-entered HOA changes are not overwritten by an automatic stack movement.

## Make Ready

Generated Make Ready is one block per pole. Imported Make Ready is reference data only.

Normal ordering:

1. Connected UG instruction.
2. Power section when applicable.
3. Regular comm movements, ordered from highest effective HOA to lowest.
4. Service Drop relocations.
5. Attach/proposed/anchor/riser instructions.
6. Low-power midspan reminder.

UG or PCO mode replaces the normal Make Ready block with the profile-specific replacement template. UG and PCO are mutually exclusive.

For INTEC, an active UG pole replaces its normal block with the editable six-line UG decision template headed by `Unable to attach due to <reason>.` The multiline editor is preloaded automatically; older saves containing only `ugReason` are migrated into its first line. When an adjacent pole is UG, the current pole extracts only that first-line reason for its relation-specific instruction, such as `Backspan to go UG SE due to on adj pole proposed pole overloaded.` The relation direction comes from the physical span. A Backspan UG enables one riser automatically. The `Riser` action beside UG/PCO is always enabled and can explicitly add or remove the riser MR without an adjacent UG pole; manual work uses the primary Proposed span. The riser is 12 inches below the current pole's primary Proposed and is also appended to an active UG/PCO replacement block. Direction precedence is saved user selection, imported Make Ready/IO, then the selected span direction; Fore Span and Back Span both keep that direction. Other Span remains blank unless a direction is imported or selected. The separate riser instruction is always the final line in the pole Make Ready.

`Transfer to New Pole` uses HOA Change when populated and otherwise uses Existing HOA. Transfer rows for the same normalized owner are combined into one instruction, their unique heights are ordered from lowest to highest, and `with DG` is appended when any related transfer has DG. Two heights use `A and B`; three or more use commas and a final `and`, for example `Transfer CTL to new pole at HOA 19'10", 20'2" and 20'6".` Owner groups themselves follow the physical pole stack from highest to lowest comm. CenturyLink, CTL and TELCO resolve to the same MR owner.

Excel Review treats the UG replacement reasons as alternatives, not cumulative instructions. One explicit underground instruction or one `Unable to attach due to <specific reason>` statement is valid. A missing or unresolved UG reason produces one consolidated Make Ready error.

Slack spans are selected by the PLA model. The Calculator does not infer or generate `Proposed slack span` from notes; Excel Review accepts that instruction as supplemental model work.

INTEC and Metronet wording is selected by `mrTemplate`/`projectProfile`; case is applied after text generation.
