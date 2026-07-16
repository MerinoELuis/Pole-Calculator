# Data Model

`AppStore` is the single source of truth. The saved JSON stores this complete state under `state` so a job can resume without rebuilding user work from display tables.

## AppState

| Field | Type | Purpose |
| --- | --- | --- |
| `version` | string | State format version. |
| `importedFileName` | string | Source job filename used for display and save/export naming. |
| `importedAt` | ISO string | Last raw import time. |
| `settings` | object | Editable clearance values and project-profile behavior. |
| `poles` | object map | Poles keyed by `poleId`. |
| `spans` | object map | Span records keyed by `spanId`. |
| `spanSides` | object map | Proposed data keyed by `spanId__poleId`. |
| `spanComms` | object map | Physical comm rows keyed by span, pole, owner, and Wire ID. |
| `spanPower` | object map | Imported power wires used to derive Low Power at MS. |
| `makeReadyReferences` | array | Imported Make Ready attachment references. |
| `excelReviewSource` | object | Original headers and rows from Collection, Span, Span.Wire, Make Ready, and Make Ready.Comm Transfers. |
| `poleClassChecks` | array | Pole height/class comparison data. |
| `mr` | array | Generated Make Ready blocks. |
| `warnings` | array | Broad validation and data-integrity warnings. |
| `updateDiagnostics` | object | Latest Update Data reconciliation summary when available. |

`excelReviewSource` is separate from normalized graph entities because empty Collection IDs, duplicate Span rows, and unused Span.Wire rows still matter during auditing. Review results and timestamps are not saved in AppState; they are regenerated from this source and current calculator work.

## Settings

Important settings and defaults:

| Field | Default | Meaning |
| --- | --- | --- |
| `projectProfile` | `INTEC` | Active project behavior. |
| `position` | `TOP_COMM` | Proposed is placed above or below the comm stack. |
| `polePowerCommsClearance` | `40"` | Low Power to comm clearance on the pole. |
| `commClearance` | `12"` | Different-owner comm spacing on the pole. |
| `boltClearance` | `4"` | Same-owner and bolt-point spacing on the pole. |
| `midspanPowerCommClearance` | `30"` | Power-to-comm spacing at midspan. |
| `midspanCommCommClearance` | `4"` | Comm-to-comm spacing at midspan. |
| `proposedOwner` | `Wecom` | Owner used for INTEC attachment Make Ready. |
| `allowLowPowerMidspanAdjustment` | `true` | Whether Proposed MS may be capped at Max Height at MS. |
| `showServiceDrop` | `true` | Whether the Service Drop control is displayed and used. |
| `showResagServiceDrop` | `true` | Whether INTEC exposes the Re-sag Service Drop control. |
| `attachmentMessengerSize` | blank | Messenger diameter exported to AutoProposed. |
| `fiberSizes` | `{}` | Fiber diameter by detected count, such as `24CT Fiber`. |

## Pole

Imported, editable, and derived values live together but have different ownership.

| Field | Ownership | Meaning |
| --- | --- | --- |
| `poleId` | imported/key | Visible pole name and graph node ID. |
| `collectionId` | imported | Collection identity used while linking Span rows. |
| `poleHeight` | imported/editable | Pole height/type reference. |
| `tipHeight` | imported | Above-ground tip measurement used by Pole Type Check. |
| `lowPower` | imported/editable | Lowest power attachment on the pole. |
| `maxCommHeight` | derived | `lowPower - Pole Power-comms clearance`. |
| `topComm`, `lowComm` | derived | Highest and lowest effective comm HOA, excluding INTEC POF rows. |

`canonicalPoleIdentity(poleId)` is used only for matching. It removes trailing `STEEL`, `UG`, and `PCO` descriptors and compares case-insensitively. The `poleId` stored from Collection remains unchanged for display and export.
| `standaloneProposedHOA` | editable | Proposed height on a terminal pole with no outgoing Proposed span. |
| `ugActive`, `pcoActive` | editable | Mutually exclusive Make Ready resolution modes. |
| `notes` | editable | User-owned pole notes. |
| `metadata` | imported reference | Source notes, location, status, and other non-authoritative fields. |

Generated poles use stable `Unknown-<spanId>`-style IDs and remain editable.

## Span

| Field | Ownership | Meaning |
| --- | --- | --- |
| `spanId` | imported/key | Internal edge identity. |
| `fromPole`, `toPole` | imported/editable relationship | Directed endpoints. |
| `type`, `rawType` | imported | Fore Span, Back Span, Other, or source spelling. |
| `direction` | derived from bearing/imported | N, NE, E, SE, S, SW, W, or NW. |
| `bearingDegrees` | imported | Numeric bearing used by downstream geometry. |
| `length`, `lengthDisplay` | imported | Raw length and feet/inches display length. |
| `environment` | imported/editable | Span environment category. |
| `environmentClearance` | defaulted/editable | Minimum comm height for the environment. |
| `midspanLowPower` | derived | Lowest valid power midspan on this span. |
| `midspanMaxCommHeight` | derived | `midspanLowPower - Midspan Power-comm clearance`. |
| `sourceSpanId` | internal relationship | Physical span supplying geometry and power data to an additional Proposed row. |
| `isManualProposed` | internal/editable workflow | Marks a user-created Proposed connection. |

Two spans describe the same physical connection when their sorted endpoint pairs are equal. They can still have different `spanId` values because Excel may contain Fore and Back records.

## SpanSide

`SpanSide` contains Proposed data for one pole side of a span.

| Field | Ownership | Meaning |
| --- | --- | --- |
| `spanId`, `poleId` | key | Links the Proposed row to a span and pole. |
| `proposedHOA` | editable | Proposed attachment height on the current pole. |
| `proposedHOAChange` | auto/editable | Next Pole Proposed used to calculate End Drop. |
| `nextPoleProposedAuto` | derived | Indicates that Next Pole Proposed came from the connected pole. |
| `ocalcMS` | editable | Decimal-feet O-Calc midspan input. |
| `proposedMidspan` | imported/legacy | Alternate Proposed midspan source. |
| `msProposed` | derived | Parsed and formatted Proposed MS base. |
| `finalMidspan` | derived | Adjusted Final MS after comm, Environment, and power checks. |
| `endDrop` | derived | `Next Pole Proposed - Proposed`. |
| `clearanceMSStatus`, `clearanceMSMessage` | derived | MS Flagging status and explanation. |
| `proposedFlaggingStatus`, `proposedFlaggingMessage` | derived | Pole-side Proposed Flagging status and explanation. |
| `notes` | editable | User notes for this proposal. |

An additional Proposed uses its own SpanSide and may point to a physical span through `sourceSpanId`.

## SpanComm

`SpanComm` is the central entity for existing communication movements.

| Field | Ownership | Meaning |
| --- | --- | --- |
| `spanId`, `poleId`, `owner`, `wireId` | key/imported | Row identity and cable relationship. |
| `ownerBase`, `rawOwner` | imported | Matching and original display owner values. |
| `existingHOA` | imported/editable | Baseline attachment height. |
| `existingHOAChange` | editable/auto | New attachment height. Blank means no movement. |
| `midspan` | imported/editable before movement | Baseline midspan from Span.Wire. |
| `ocalcMS` | imported/legacy | Fallback baseline for older states. |
| `calculatedMidspan` | derived | Baseline plus half-movements at both endpoints. |
| `msProposed`, `finalMidspan` | derived | Current display and clearance result for the comm. |
| `remotePoleId`, `remoteHOA` | derived | Selected opposite endpoint and effective height. |
| `flaggingStatus`, `flaggingMessage` | derived | Combined pole and midspan issues. |
| `serviceDrop`, `downGuy` | imported/editable | Make Ready modifiers. |
| `transferToNewPole` | editable | Comm-level transfer flag synchronized across the comm's `SpanComm` rows; generates one transfer instruction. |
| `resagServiceDrop` | editable | INTEC-only resolution that treats a low service-drop midspan as 15'6". |
| `size`, `construction`, `insulator` | imported reference | Source wire details. |

## REF Semantics

REF is a display/calculation state, not a different entity type.

- Back Span rows are always treated as REF for calculations on the current row.
- Other rows are REF when they do not own a valid imported midspan.
- Other rows with their own midspan calculate normally.
- Fore Span rows are calculation rows; a missing midspan can be entered manually.
- A REF row never borrows a baseline midspan from a previous or next span.

A Back Span row that contains an imported midspan may display the source value for inspection, but it remains reference-only in the current calculation path.

## SpanPower

Power rows come from Span.Wire records classified as Primary, Secondary, Neutral, Power, or similar utility wires.

| Field | Meaning |
| --- | --- |
| `spanId`, `poleId`, `wireId` | Relationship identity. |
| `label` | Simplified Primary/Secondary/Neutral/Power type. |
| `attachmentHeight` | Power attachment height on the pole. |
| `midspan` | Editable power midspan used in the lowest-power calculation. |
| `owner`, `size` | Imported reference values. |

## Height Representation

All arithmetic uses integer inches. Display strings are produced only at input/output boundaries.

Accepted examples:

| Input | Internal value | Display |
| --- | ---: | --- |
| `20` | 240 | `20'` |
| `20'` | 240 | `20'` |
| `20'6"` | 246 | `20'6"` |
| `20.5` | 246 | `20'6"` |
| `6"` | 6 | `6"` |

## Minimal Saved-State Example

```json
{
  "version": "1.3.0",
  "importedFileName": "EXCEL_JOB_2026-07-13.xlsx",
  "settings": {
    "projectProfile": "INTEC",
    "position": "TOP_COMM"
  },
  "poles": {
    "P01": {
      "poleId": "P01",
      "lowPower": "28'8\"",
      "maxCommHeight": "25'4\""
    }
  },
  "spans": {},
  "spanSides": {},
  "spanComms": {},
  "spanPower": {}
}
```

Constructors in `state.js` fill missing fields during `normalizeState()`. Consumers should not assume this abbreviated example is the complete serialized shape.
