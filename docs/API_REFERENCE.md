# Public API Reference

The application does not use ES modules. Public APIs are attached to `window` and must be available before dependent scripts run. Internal helpers not listed here can change without becoming cross-module contracts.

## HeightUtils

| Method | Returns | Purpose |
| --- | --- | --- |
| `parseHeight(value)` | `number|null` | Convert supported input to integer inches. |
| `formatHeight(inches, options?)` | string | Format inches as compact feet/inches. |
| `addHeights(a, b)` | string | Add two parseable heights. |
| `subtractHeights(a, b)` | string | Subtract the second height from the first. |
| `decimalFeetToHeight(value)` | string | Convert decimal feet to feet/inches. |
| `heightToDecimalFeet(value)` | number or blank | Convert feet/inches to decimal feet. |
| `inchesToHeight(value)` | string | Convert an inches-only value. |
| `compareHeights(a, b)` | `-1`, `0`, `1`, or `null` | Compare two parseable heights. |
| `isValidHeight(value, allowBlank?)` | boolean | Validate a user/imported height. |
| `diffLabel(from, to)` | string | Return signed `to - from`. |

## ProjectProfiles

| Method/property | Purpose |
| --- | --- |
| `PROFILES` | INTEC and Metronet profile definitions. |
| `normalizeProfileId(value)` | Resolve unknown profile names to INTEC. |
| `getProfile(value)` | Return one profile definition. |
| `applyProfileSettings(settings, profileId)` | Apply profile defaults over settings. |
| `detectProfile({fileName, owners})` | Detect Metronet markers, otherwise INTEC. |

## AppStore

### State lifecycle

| Method | Purpose |
| --- | --- |
| `getState()` | Return the mutable current AppState. |
| `setState(nextState)` | Normalize and replace current state. |
| `resetState()` | Create a clean state with defaults. |
| `saveToLocal()` / `loadFromLocal()` | Legacy localStorage helpers; primary UI persistence uses Save/Load files. |
| `normalizeState(state)` | Upgrade and reconcile imported/saved state. |

### Constructors and keys

| Method | Purpose |
| --- | --- |
| `createPole(...)` | Normalize a Pole. |
| `createComm(...)` | Normalize a pole-level comm summary. |
| `createSpan(...)` | Normalize a Span and generate an Unknown endpoint when required. |
| `createSpanSide(data)` | Normalize Proposed state. |
| `createSpanComm(data)` | Normalize an existing comm/span row. |
| `createSpanPower(data)` | Normalize a power/span row. |
| `createMakeReadyReference(data)` | Normalize an imported attachment reference. |
| `keyForSpanSide(spanId, poleId)` | Build a SpanSide map key. |
| `keyForSpanComm(spanId, poleId, owner, wireId)` | Build a SpanComm map key. |

### Queries

| Method | Purpose |
| --- | --- |
| `getPole(poleId)` / `getSpan(spanId)` | Fetch graph entities. |
| `getSpanSide(spanId, poleId)` | Fetch one Proposed row. |
| `getSpanComm(spanId, poleId, owner, wireId?)` | Fetch one exact comm row. |
| `getConnectedSpans(poleId)` | Fetch all graph edges touching a pole. |
| `getOtherPoleId(span, poleId)` | Resolve the opposite endpoint. |
| `getSpanSidesForPole/ForSpan(...)` | Query Proposed rows. |
| `getSpanCommsForPole/ForSpan(...)` | Query comm rows. |
| `getSpanPowerForPole/ForSpan(...)` | Query power rows. |
| `poleHasChanges(poleId)` | Report whether user work exists on a pole. |

### Mutations

`upsertPole`, `upsertComm`, `upsertSpan`, `upsertSpanSide`, `upsertSpanComm`, and `addSpanPower` normalize and store entities. Field update helpers restrict editable fields. Removal helpers delete a SpanSide, manual span, or SpanComm. Mutations do not universally guarantee recalculation; UI code normally calls a `Calculations` update/recalculation entry point.

`ensureSpanSides`, `ensureSpanComms`, and `ensureEndpointComms` repair required relationships during normalization. They should not be used to invent business values.

## Calculations

### User-edit entry points

| Method | Purpose |
| --- | --- |
| `updateExistingHOAChange(...)` | Store a new HOA and recalculate both endpoints. |
| `updateSpanCommField(...)` | Update an allowed SpanComm field, including Service Drop, DG, transfer and re-sag controls, then recalculate. |
| `clearSpanCommMidspan(...)` | Clear only the selected midspan and its derived fields. |
| `updateSpanSideField(...)` | Update Proposed, Next Pole Proposed, O-CALC MS, End Drop, or notes. |
| `updateSpanField(...)` | Update an allowed physical span field and recalculate. |

### Calculation and validation

| Method | Purpose |
| --- | --- |
| `calculateMidspanForComm(spanComm)` | Persist endpoint-adjusted comm midspan and flagging. |
| `findRemoteComm(...)` | Select the opposite comm using Wire ID/span/owner ranking. |
| `displayMidspanForComm(spanComm)` | Return the same value source used by the comm UI. |
| `calculateSpanPowerDerived(spanId)` | Derive Low Power and Max Height at MS. |
| `calculateSpanSideMidspan(spanId, poleId)` | Derive MS Proposed, Adjusted Final MS, and MS Flagging. |
| `calculateEndDropForSpanSide(spanId, poleId)` | Resolve Next Pole Proposed and End Drop. |
| `evaluateCommFlagging(...)` | Evaluate combined comm pole/midspan rules. |
| `evaluateSpanSideFlagging(...)` | Evaluate compact Proposed Flagging. |
| `evaluateProposedPoleClearance(...)` | Check comm and bolt spacing for Proposed. |
| `getReferenceMidspansForSpanSide(...)` | Get Proposed references from the same physical connection. |

### Recalculation

| Method | Purpose |
| --- | --- |
| `recalculateSpan(spanId)` | Recalculate one edge and endpoints. |
| `recalculateSpansForPole(poleId)` | Recalculate a pole network neighborhood and reciprocal Wire IDs. |
| `recalculateAll()` | Rebuild every derived value, MR block, and warning. |
| `autoCalculateMovements()` | Run the iterative Top Comm movement solver. |

## ExcelImport

| Method/property | Purpose |
| --- | --- |
| `importExcelFile(file)` | Detect and import raw/exported workbook data. |
| `importJsonFile(file)` | Restore a saved JSON state. |
| `importDataFile(file)` | Route JSON or Excel by extension/type. |
| `rowsToObjects`, `pick`, `findSheet` | Reusable import helpers. |
| `directionFromBearingDisplay(value)` | Normalize bearing and cardinal direction. |
| `recalculatePoleClassCheck(row)` | Re-evaluate one Pole Type Check row. |
| ANSI constants | Reference table data used by Pole Type Check. |

## MRLogic

| Method | Purpose |
| --- | --- |
| `generateMRForComm(spanComm)` | Generate one movement line. |
| `generateResagServiceDropMR(spanComm)` | Generate the INTEC re-sag instruction when its conditions apply. |
| `generateMRForSpanSide(spanSide)` | Generate Proposed/anchor/riser/slack lines. |
| `generateMRForSpan(spanId)` | Generate all lines related to one span. |
| `generateMRForPole(poleId)` | Replace one pole's ordered MR block. |
| `generateAllMR()` | Replace generated MR for the complete state. |
| `detectAttach/RaiseLower/Overlash/Slack/Anchor/Riser` | Expose MR condition detectors. |

## ExcelReview

| Method | Purpose |
| --- | --- |
| `runReview()` | Clear and rebuild the complete HOA and Final review from current source/state. |
| `reviewPole(poleId)` | Return the latest result for one Collection pole. |
| `getResults()` | Return latest naturally sorted pole results. |
| `getSummary()` | Return Errors, Warnings, Passed, Final Not Ready, and total counts. |
| `getReviewState()` | Return results, global checks, summary, and review timestamp. |
| `clearResults()` | Remove in-memory review results without changing AppState. |

## Validations

| Method | Purpose |
| --- | --- |
| `validatePole(poleId)` | Replace broad warnings for one pole. |
| `validateSpan(spanId)` | Replace broad warnings for one span. |
| `validateAll()` | Rebuild all broad warnings. |

## ProjectExport

| Method | Purpose |
| --- | --- |
| `exportProposedJson()` | Download compact AutoProposed data for O-Calc. |
| `exportDebugJson()` | Download full state and calculation traces. |
| `exportJson()` | Legacy full-state download helper. The primary UI uses Save. |
| `downloadJson(filename, data)` | Browser JSON download utility. |

## UI Boundary

`app.js` intentionally does not publish a broad global API. It owns DOM rendering and events. Business logic needed by another module should be exposed through the relevant domain module rather than by calling an `app.js` function.
