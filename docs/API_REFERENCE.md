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
| `setState(nextState)` | Normalize and replace current state. `AutoCalculateSourceCompat` preserves automatic Proposed source fields around this call. |
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
| `canonicalPoleIdentity(poleId)` | Build a stable matching key that ignores trailing STEEL/UG/PCO descriptors. |
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

`upsertPole`, `upsertComm`, `upsertSpan`, `upsertSpanSide`, `upsertSpanComm`, and `addSpanPower` normalize and store entities. `updatePowerEquipmentField(poleId, index, field, value)` stores `actionActive` or `actionHeight` without changing imported equipment heights. Other field update helpers restrict editable fields. Removal helpers delete a SpanSide, manual span, or SpanComm. Mutations do not universally guarantee recalculation; UI code normally calls a `Calculations` update/recalculation entry point.

`ensureSpanSides`, `ensureSpanComms`, and `ensureEndpointComms` repair required relationships during normalization. They should not be used to invent business values.

## Calculations

### User-edit entry points

| Method | Purpose |
| --- | --- |
| `updateExistingHOAChange(...)` | Store a new HOA and recalculate both endpoints. The installed solver wrapper clears the AUTO source when the user edits it. |
| `updateSpanCommField(...)` | Update an allowed SpanComm field, including Service Drop, DG, transfer and re-sag controls, then recalculate. |
| `clearSpanCommMidspan(...)` | Clear only the selected midspan and its derived fields. |
| `updateSpanSideField(...)` | Update Proposed, Next Pole Proposed, O-CALC MS, End Drop, or notes. A user Proposed edit clears the AUTO source. |
| `updateSpanField(...)` | Update an allowed physical span field and recalculate. |

### Calculation and validation

| Method | Purpose |
| --- | --- |
| `calculateMidspanForComm(spanComm)` | Persist endpoint-adjusted comm midspan and flagging. |
| `findRemoteComm(...)` | Select the opposite comm using Wire ID/span/owner ranking. |
| `displayMidspanForComm(spanComm)` | Return the same value source used by the comm UI. |
| `calculateSpanPowerDerived(spanId)` | Derive Low Power and Max Height at MS. |
| `calculateSpanSideMidspan(spanId, poleId)` | Derive MS Proposed, Adjusted Final MS, and MS Flagging, including the Wecom/MidAm same-span or estimated-sag fallback. |
| `calculateEndDropForSpanSide(spanId, poleId)` | Resolve Next Pole Proposed and End Drop. |
| `evaluateCommFlagging(...)` | Evaluate combined comm pole/midspan rules. |
| `evaluateSpanSideFlagging(...)` | Evaluate compact Proposed Flagging. |
| `evaluateProposedPoleClearance(...)` | Check comm and bolt spacing for Proposed. |
| `getReferenceMidspansForSpanSide(...)` | Get Proposed references from the same physical connection. |
| `getPowerEquipmentCeiling(equipment)` | Return the effective comm ceiling for an imported equipment row and its active action. |

### Recalculation

| Method | Purpose |
| --- | --- |
| `recalculateSpan(spanId)` | Recalculate one edge and endpoints. |
| `recalculateSpansForPole(poleId)` | Recalculate a pole network neighborhood and reciprocal Wire IDs. |
| `recalculateAll()` | Rebuild every derived value, MR block, and warning. |
| `autoCalculateMovements(options?)` | Asynchronously run the installed TOP/LOW COMM best-arrangement solver. Returns a Promise with SAFE/BEST AVAILABLE/CRITICAL summary counts and accepts `onProgress(detail)`. |

## AutoCalculateSolver

| Method/property | Purpose |
| --- | --- |
| `RESULT_STATUS` | SAFE, BEST_AVAILABLE, CRITICAL, MANUAL, and SKIPPED constants. |
| `modeFromState(state?)` | Resolve TOP_COMM or LOW_COMM from current settings. |
| `normalizeOwner(value)` | Normalize common owner aliases before grouping. |
| `groupsForPole(poleId)` | Group duplicate span relationships that represent one physical comm. |
| `proposedSpansForPole(poleId)` | Return eligible forward/manual Proposed relationships. |
| `idealProposedHeight(groups, mode, state?)` | Calculate Top Comm + clearance or Low Comm - clearance. |
| `candidateHeights(options)` | Generate progressive direct, pole-limit, Midspan-boundary and fallback integer-inch candidates. |
| `buildStackPlan(groups, proposed, mode, maxPole, state?)` | Build a TOP downward or LOW upward comm stack. |
| `collectIssues(poleId)` | Separate current-pole violations from physical-span Midspan violations. |
| `issueSeverity(message)` | Estimate remaining clearance shortfall from a validation message. |
| `categoryForAnalysis(value)` | Map an evaluated result to safe, pole-safe partial, or pole-failing category. |
| `rankAnalysis(value)` / `compareAnalyses(a, b)` | Apply strict pole-first candidate ordering. |
| `statusForAnalysis(value)` | Return SAFE, BEST_AVAILABLE, or CRITICAL. |
| `analyzeCurrentState(poleId, mode)` | Return violations, movement cost, and ideal-distance metrics. |
| `solvePole(poleId, mode, options?)` | Asynchronously evaluate and retain the best aerial arrangement for one pole. `onCandidateProgress(detail)` reports candidate batches. |
| `autoCalculateMovements(options?)` | Process every pole once, then selectively retry only affected poles that can improve, with a three-attempt cap per pole. `onProgress(detail)` reports phase, percentage, pass, pole and candidate. |
| `install()` | Replace `Calculations.autoCalculateMovements()` with the solver entry point. |

## AutoCalculateSourceCompat

| Method/property | Purpose |
| --- | --- |
| `PROPOSED_SOURCE_FIELDS` | `autoCalcProposedStatus` and `autoCalcProposedMode`. |
| `keyForSide(side)` | Build the normalized SpanSide identity. |
| `copySourceFields(target, source)` | Restore automatic Proposed source fields after constructor normalization. |

The module wraps `AppStore.upsertSpanSide()` and `AppStore.setState()` before `app.js` loads. It preserves source metadata through candidate copies, Undo, and Save/Load without changing clearance formulas.

## ExcelImport

| Method/property | Purpose |
| --- | --- |
| `importExcelFile(file)` | Detect and import raw/exported workbook data. |
| `importJsonFile(file)` | Restore a saved JSON state. |
| `importOriginalWorkbook(workbook, fileName)` | Import an already parsed raw workbook; primarily used by dependency-free regression tests. |
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
| `generatePowerEquipmentMRForPole(poleId)` | Generate Ground, Transformer Secure, and independent Power Riser Raise/Secure instructions. |
| `generateMRForSpanSide(spanSide)` | Generate Proposed/anchor/riser/slack lines. |
| `generateMRForSpan(spanId)` | Generate all lines related to one span. |
| `generateMRForPole(poleId)` | Replace one pole's ordered MR block. CompactAutoProposed augments and filters this method after load. |
| `generateAllMR()` | Replace generated MR for the complete state. CompactAutoProposed preserves the same public entry point. |
| `getResolvedRiserDirection(poleId)` | Resolve saved, imported, or relation-derived INTEC riser direction. |
| `getDefaultRiserDirection(relation, direction)` | Keep the UG span direction for Fore/Back and leave Other Span blank. |
| `isRiserAvailable/isRiserEnabled(poleId)` | Disable Riser on UG/PCO poles and expose automatic Fore/Back adjacent-UG or manual state elsewhere. |
| `generateRiserInstruction(poleId)` | Build the final riser line from primary Proposed, saved/imported direction, and span fallback. |
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
| `exportProposedJson()` | Download compact AutoProposed data for O-Calc. Replaced at runtime by `CompactAutoProposed.exportCompactProposedJson`. |
| `exportDebugJson()` | Download full state and calculation traces. |
| `exportJson()` | Legacy full-state download helper. The primary UI uses Save. |
| `downloadJson(filename, data)` | Browser JSON download utility. |

## CompactAutoProposed

| Method | Purpose |
| --- | --- |
| `applyAttachmentDefaults(state?)` | Apply supported messenger/fiber defaults when references justify them. |
| `detectedReferenceFiberCounts(state)` | Return fiber counts found in Make Ready references. |
| `fiberEntries(state)` | Return configured/detected fiber rows for the settings UI. |
| `buildCompactPayload(state?)` | Recalculate-ready compact serialization with integrated UG/PCO sanitization. |
| `validationErrors(payload)` | Return missing size requirements for aerial fiber attachments. |
| `compactMoves(state, poleId)` | Serialize deduplicated local comm movements. |
| `compactSpansForPole(state, poleId)` | Serialize physical spans, geometry, aerial proposals and UG markers. |
| `movementCandidates(state, poleId)` | Return normalized movement candidates before compact field reduction. |
| `augmentPoleMakeReady(poleId)` | Add compact movement instructions to the current pole MR. |
| `exportCompactProposedJson()` | Recalculate, validate and download the compact payload. |
| `isPoleFullyUg(state, poleId)` | Resolve active or name-token UG state. |
| `isPolePco(state, poleId)` | Resolve active or name-token PCO state. |
| `blocksLocalActions(state, poleId)` | Report whether local export/MR movement actions are blocked. |
| `sanitizeBlockedLocalPayload(payload, state?)` | Remove local proposal/movement fields from UG/PCO pole entries. |
| `sanitizeBlockedLocalMr(state, poleId)` | Remove generated regular movement lines from UG/PCO replacement MR. |
| `isUgSpan(state, span, poleId?)` | Resolve whether one physical relation exports as UG. |
| `spanKind(span)` | Map imported relation type to `F`, `B`, or `O`. |
| `spanLengthInches(span)` | Convert display/raw span length to integer inches. |
| `directionFromPole(span, poleId)` | Resolve direction from one endpoint. |
| `directionTokensForReference(ref)` | Normalize reference direction tokens. |

## FloatingCalculator

| Method | Purpose |
| --- | --- |
| `setupFloatingCalculator()` | Bind calculator panel controls and live evaluation. |
| `normalizeInchQuotes(value)` | Convert every `''` pair to `"`. |
| `normalizeExpressionInput(input)` | Normalize the field value while preserving cursor/selection positions. |

## UiDomContract

| Method | Purpose |
| --- | --- |
| `normalizeLabel(value)` | Normalize visible labels for stable keys. |
| `columnKey(value)` | Convert a table heading into a `data-column` value. |
| `spanIdFrom(element)` | Read stable or legacy span identity from an element. |
| `pairCommRows(tableRow)` | Pair communication span and Midspan rows by their group position. |
| `annotateTableColumns(table)` | Add column identity to headers and cells. |
| `annotatePoleCard(card)` | Add component, pole and span identities to one card. |
| `apply(root?)` | Annotate every rendered pole card below a root. |

## CommTableUI

| Method/property | Purpose |
| --- | --- |
| `HIDDEN_COLUMN_KEY` | Stable key for the hidden `Other Pole HOA` column. |
| `findColumnIndex(headers)` | Find the hidden column by stable key, with legacy label fallback. |
| `removeColumnFromTable(table, index)` | Remove one presentation column from every table row. |
| `refresh(root?)` | Apply the visual column policy to rendered communication tables. |

## SpanColorUI

| Method/property | Purpose |
| --- | --- |
| `COLOR_CLASS_COUNT` | Number of reusable span color classes. |
| `physicalSpanId(store, spanId)` | Resolve a synthetic Proposed span to its source span. |
| `sortVisibleSpanIds(store, poleId, spanIds)` | Order only visible spans for one pole. |
| `classForVisibleIndex(index)` | Map a local visible index to `span-color-N`. |
| `firstSpanId(element)` | Prefer stable `data-span-id`, then legacy markup. |
| `collectCommRowPairs(tableRow)` | Pair communication span and Midspan rows. |
| `resetCard(card, store?)` | Restart and apply the sequence for one pole card. |
| `refresh(root?, store?)` | Reapply colors to all rendered pole cards. |

## PoleCalculatorUI

| Method | Purpose |
| --- | --- |
| `runRefresh(root?)` | Run DOM annotation, table cleanup, and span colors in order. |
| `queueRefresh()` | Coalesce mutations into one microtask. |
| `start()` | Run initial refresh and attach the workspace observer. |
| `stop()` | Disconnect the observer, primarily for controlled tests. |

## UI Boundary

`app.js` intentionally does not publish a broad global API. It owns primary DOM rendering and events. Business logic needed by another module should be exposed through the relevant domain module. Presentation-only post-render behavior should be implemented in `js/ui/` and coordinated through `PoleCalculatorUI` rather than by adding another independent workspace observer.
