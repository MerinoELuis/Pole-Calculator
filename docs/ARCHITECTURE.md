# Architecture

This document describes the current runtime architecture of Pole Span MR Calculator.

## Runtime Model

The application is a static browser application. It has no backend, framework, runtime package dependency, or production build step. GitHub Pages serves the HTML, CSS, and JavaScript files directly.

Node and `package.json` provide the dependency-free `npm test` tooling; the package's numeric version is also used as the release identifier during deployment. They do not add a runtime dependency to the browser application.

The numeric application version follows `package.json` semantic versioning. GitHub Pages injects that version plus the exact commit, branch, deployment time, and Actions run into `js/core/deployment-version.js`; the header displays the numeric version and build identifier.

Every browser module is an IIFE that publishes a small API on `window`.

`index.html` loads modules in dependency order:

1. `libs/xlsx.full.min.js`
2. `js/core/deployment-version.js`
3. `js/domain/height-math.js`
4. `js/core/project-config.js`
5. `js/domain/pole-resolution.js`
6. `js/core/state-compat.js`
7. `js/core/store-identity.js`
8. `js/core/store.js`
9. `js/domain/comm-movements.js`
10. `js/domain/pole-inset-calculation.js`
11. `js/domain/clearance-rules.js`
12. `js/domain/midspan-engine.js`
13. `js/rules/mr-logic.js`
14. `js/rules/validations.js`
15. `js/import/import-normalizers.js`
16. `js/import/excel-import.js`
17. `js/import/update-data-reconciler.js`
18. `js/import/excel-review-wire-checks.js`
19. `js/import/excel-review.js`
20. `js/export/excel-export.js`
21. `js/export/json-export.js`
22. `js/export/compact-autoproposed.js`
23. `js/rules/auto-calculate-solver.js`
24. `js/ui/dialogs.js`
25. `js/ui/theme.js`
26. `js/ui/mobile-keyboard.js`
27. `js/ui/floating-calculator.js`
28. `js/ui/excel-review-renderer.js`
29. `js/ui/pole-renderer.js`
30. `js/ui/controls.js`
31. `js/core/persistence.js`
32. `js/core/history.js`
33. `js/ui/pole-navigation.js`
34. `js/ui/app.js`
35. `js/ui/dom-contract.js`
36. `js/ui/comm-table.js`
37. `js/ui/span-colors.js`
38. `js/ui/runtime.js`

Changing this order can break global dependencies. Pole resolution, state compatibility, and identity helpers load before the store; movement and Pole Inset calculation modules register before clearance calculations inject their helpers; import normalizers load before ExcelImport; Excel Review wire rules load before the audit coordinator, and its result renderer loads before the app initializes it. Update Data reconciliation loads after both the store and importer. The post-render runtime loads after the renderer because it annotates and normalizes generated DOM.

## Module Responsibilities

| Module | Global API | Responsibility |
| --- | --- | --- |
| `core/store.js` | `AppStore` | Own normalized mutable application state and identity/relation accessors. |
| `core/state-compat.js` | `StateCompatibility` | Normalize older saved/imported state shapes at explicit compatibility boundaries. |
| `core/store-identity.js` | `AppStoreIdentity` | Canonical pole identities and stable span-side/communication keys. |
| `core/history.js` | `AppHistory` | Own undo/redo snapshots and the movement trace written to Debug data. |
| `core/persistence.js` | `AppPersistence` | Save/load JSON, file handles, dirty-state tracking, and safe job filenames. |
| `core/project-config.js` | `ProjectProfiles` | Keep INTEC and Metronet defaults outside calculation code. |
| `core/deployment-version.js` | `DeploymentVersion` | Expose app version and deployment metadata. |
| `domain/height-math.js` | `HeightUtils` | Parse and format height values using integer inches. |
| `domain/pole-resolution.js` | `PoleResolution` | Normalize AERIAL, UG, and PCO state with UG priority. |
| `domain/comm-movements.js` | `CalculationAutoMovementModule` | Encapsulate Auto Calculate movement planning and apply routines using helpers injected by clearance calculations. |
| `domain/pole-inset-calculation.js` | `PoleInsetCalculationModule` | Calculate Pole Inset span options, selected-span power ceilings, and End Drops. |
| `domain/clearance-rules.js` | `Calculations` | Derive clearances, pole limits, comm midspans, Proposed results, End Drop, recalculation, and flagging. This remains the integrated calculation orchestrator. |
| `domain/midspan-engine.js` | `MidspanEngine` (`MidspanLogic` compatibility alias) | Calculate comm Midspan from the imported baseline plus half of each endpoint movement; also provide span-oriented Midspan read/recalculation APIs. |
| `domain/graph.js` | `PoleGraph` | Graph utilities for pole/span topology; not loaded by the browser runtime unless explicitly needed. |
| `rules/mr-logic.js` | `MRLogic` | Generate ordered Make Ready text for each pole. |
| `rules/auto-calculate-solver.js` | `AutoCalculateSolver` | Search TOP/LOW COMM arrangements, rank SAFE/BEST AVAILABLE/CRITICAL results, and install the public Auto Calculate entry point. |
| `rules/validations.js` | `Validations` | Generate broad data-integrity warnings. Table flagging remains in `Calculations`. |
| `import/import-normalizers.js` | `ImportNormalizers` | Normalize worksheet rows, headers, heights, identifiers, and common cell values. |
| `import/excel-import.js` | `ExcelImport` | Convert Excel worksheets or saved data into normalized AppStore entities. |
| `import/excel-review.js` | `ExcelReview` | Audit preserved workbook rows against current calculated and MR state without rendering HTML. |
| `import/excel-review-wire-checks.js` | `ExcelReviewWireChecks` | Isolated INTEC wire construction/insulator and Telco/CATV size checks. |
| `import/update-data-reconciler.js` | `UpdateDataReconciler` | Merge refreshed workbook data into saved operator work and report differences. |
| `export/excel-export.js` | `ExcelExport` | Produce workbook downloads. |
| `export/json-export.js` | `ProjectExport` | Produce general project and diagnostic JSON downloads. |
| `export/compact-autoproposed.js` | `CompactAutoProposed` | Produce compact Web-to-plugin data with UG/PCO safety and movement MR. |
| `ui/app.js` | browser coordinator | Initialize UI modules and coordinate events, history, persistence, dialogs, and targeted refreshes. |
| `ui/pole-renderer.js` | `PoleWorkspaceRenderer` | Render pole cards, communication and Proposed tables, review panels, and targeted pole refreshes. |
| `ui/controls.js` | `PoleActions` | Handle pole action controls, comm edits, and Proposed span actions. |
| `ui/dialogs.js` | `AppDialogs` | Shared confirmation, prompt, and message dialogs. |
| `ui/theme.js` | `ThemeControl` | Resolve and persist the active light/dark theme. |
| `ui/mobile-keyboard.js` | `MobileKeyboardControl` | Manage mobile viewport and keyboard-specific input behavior. |
| `ui/floating-calculator.js` | `FloatingCalculator` | Evaluate feet/inches expressions and normalize quotes. |
| `ui/excel-review-renderer.js` | `ExcelReviewRenderer` | Render Excel Review results, badges, warning controls, and actions. |
| `ui/pole-navigation.js` | `PoleNavigation` | Own pole-index drawer visibility, pole selection, and scroll navigation. |
| `ui/dom-contract.js` | `UiDomContract` | Annotate rendered cards and tables with stable component, pole, span, and column identities. |
| `ui/comm-table.js` | `CommTableUI` | Hide visual Other Pole HOA without deleting calculation data. |
| `ui/span-colors.js` | `SpanColorUI` | Restart visible span colors per pole and synchronize each Midspan row. |
| `ui/runtime.js` | `PoleCalculatorUI` | Coordinate post-render modules through one queued refresh and one workspace observer. |

The quick ownership index is in `FILE_MAP.md`.

## Data Flow

```mermaid
flowchart LR
    A[Raw Excel] --> B[ExcelImport]
    J[Saved JSON] --> B
    B --> C[AppStore normalized state]
    B --> R[Preserved review source]
    C --> D[Calculations]
    D --> C
    C --> S[AutoCalculateSolver]
    S --> C
    C --> E[MRLogic]
    E --> C
    C --> F[Validations]
    F --> C
    C --> G[app.js render]
    G --> U[UI DOM contract and status modules]
    R --> X[ExcelReview]
    C --> X
    X --> G
    G --> H[User edit]
    H --> C
    C --> I[Save / AutoProposed / Debug JSON]
```

A normal edit follows this lifecycle:

1. `app.js` records an undo snapshot.
2. The edit is normalized and written through `AppStore` or `Calculations`.
3. The affected span, pole, remote endpoint, and reciprocal wire rows are recalculated.
4. Make Ready and validations are regenerated.
5. Affected pole cards are rendered again.
6. `PoleCalculatorUI` annotates and normalizes the resulting DOM.
7. The job is marked dirty until Save succeeds.

`render()` also calls `recalculateAll()` before drawing the full workspace. This protects imports, undo restores, loaded JSON, and solver candidates from displaying stale derived values.

## Auto Calculate Search Boundary

`calculations.js` remains authoritative for formulas and validation. `auto-calculate-solver.js` does not duplicate clearance rules. It:

1. Clones the current normalized state.
2. Generates Proposed candidates in integer inches.
3. Builds a TOP COMM downward stack or LOW COMM upward stack.
4. Applies one candidate to a clean state copy.
5. Calls the existing recalculation and flagging APIs.
6. Measures the resulting pole and Midspan issues.
7. Retains the lexicographically best arrangement.

The solver can retain a partial result. A pole-compliant arrangement with a remaining Midspan problem always outranks an arrangement where Midspan passes but the pole fails. UG and PCO remain operator-controlled construction decisions.

User-entered HOA Change and Proposed values are distinguished from values carrying `AUTO` source metadata. Manual values are fixed constraints; automatic values may be replaced during a later run.

## UI Post-render Contract

`ui/pole-renderer.js` owns pole-card rendering. `ui/controls.js` owns the handlers for actions generated by those cards. These UI modules receive callbacks and business APIs from `ui/app.js`; they do not own application state or calculation rules. Update Data reconciliation is kept in `import/update-data-reconciler.js`, while the app coordinates file selection and refresh lifecycle.

The post-render sequence is:

1. Add stable `data-component`, `data-pole-id`, `data-span-id`, and `data-column` values.
2. Remove presentation-only columns such as Other Pole HOA.
3. Re-annotate only if column removal changed the table structure.
4. Apply the per-pole span color sequence to span and Midspan elements.
5. Add or update the per-pole Auto Calculate result card.

Only `ui/runtime.js` observes the pole workspace. Individual UI modules expose synchronous `refresh(root)` functions and do not create competing workspace observers.

## Graph Model

Poles are graph nodes and spans are directed imported edges. Direction is useful for Fore/Back/Other semantics, but calculations compare unordered pole pairs when two records describe the same physical connection.

Collection owns the visible Pole ID. Import and Update Data use `AppStore.canonicalPoleIdentity()` to resolve labels that differ only by trailing `STEEL`, `UG`, or `PCO` descriptors.

Important identities:

- Pole: `poleId`
- Span: `spanId`
- Span side: `spanId + poleId`
- Span comm: `spanId + poleId + owner + wireId`
- Physical connection: sorted pair of `fromPole` and `toPole`

Additional Proposed rows can use a synthetic span with `sourceSpanId`. The synthetic row owns independent Proposed values while inheriting geometry, Environment, and power-derived limits from the physical span.

## Recalculation Scope

| Entry point | Scope |
| --- | --- |
| `recalculateSpan(spanId)` | One span, its endpoints, comm rows, span sides, and End Drop. |
| `recalculateSpansForPole(poleId)` | Connected spans, both endpoints, reciprocal rows matched by Wire ID, MR, and validation. |
| `recalculateAll()` | Every span, pole, comm, Proposed side, MR block, and warning. |
| `AutoCalculateSolver.solvePole(poleId, mode, options?)` | Asynchronous candidate evaluation for one pole plus affected endpoint recalculation. It yields between candidate batches. |
| `Calculations.autoCalculateMovements(options?)` | Installed asynchronous solver entry point that processes every pole once, selectively retries affected improvable poles, and reports UI progress. |

Edits to an endpoint HOA normally use `recalculateSpansForPole` because one movement can affect a midspan displayed from the opposite pole.

## Save, Load, and Update Data

Save serializes the complete normalized state after a full recalculation. Load restores the full state and recalculates derived data. Update Data imports a fresh workbook and then reconciles previous user work by exact identity and physical fallbacks.

Auto-generated HOA/Proposed source markers and `metadata.autoCalculateResult` are part of normalized saved state. A later Auto Calculate run may replace values marked `AUTO`; a user edit clears that marker and becomes fixed input.

Auto Calculate remains on the browser main thread, but yields control between
small candidate batches and poles. This lets the browser paint the progress
overlay and process events during large jobs without changing candidate order
or calculation rules.

Raw Excel Review sheets and Pole Type Check rows are merged by source identity so a partial update does not remove the rest of the job. Derived values and review results are recalculated after merging.

## Compact AutoProposed Export

The formal structural contract is `schemas/autoproposed.schema.json`, with business details in `AUTOPROPOSED_CONTRACT.md`.

- Normal poles may export spans, movements, and terminal attachments.
- PCO poles keep geometry-only local spans and cannot export local movements or terminal attachments.
- Fully UG poles keep geometry-only local spans marked `ug: true` and cannot export local movements or terminal attachments.
- UG takes priority when both UG and PCO are active.
- A blocked pole with no remaining geometry is omitted.

The same module filters generated regular movement lines from UG/PCO replacement Make Ready.

## Project Profiles

`core/project-config.js` is the extension point for customer defaults. When the selected department changes, profile-owned clearances reset to the shared baseline first and then the selected profile's overrides are applied. This prevents a clearance omitted by one profile from inheriting the prior department's value.

- INTEC: Top Comm by default, proposed owner visible, Service Drop/Re-sag visible, low-power Proposed MS adjustment enabled.
- Metronet: Low Comm by default, WI selector with MidAm, Service Drop/Re-sag hidden, low-power Proposed MS adjustment disabled.
- Olsson OPPD: Top Comm by default, proposed owner visible, Service Drop relocation visible, INTEC-only Re-sag hidden.

Auto Calculate supports either position setting for either profile. Profile-specific validations and active equipment actions remain authoritative during candidate evaluation.

## Architectural Invariants

1. Heights are converted to integer inches before arithmetic.
2. Imported values and calculated values remain separate fields.
3. A comm midspan is never borrowed from an unrelated span.
4. Comm-to-comm MS comparisons use rows from the same physical relation defined by current rules.
5. `sourceSpanId` carries physical data for additional Proposed rows.
6. Save and export recalculate before serializing.
7. Reciprocal Fore/Back rows may have different Span/Wire IDs; endpoint-pair matching keeps physical calculations synchronized.
8. Make Ready is generated from current state, not imported Make Ready.
9. Other Pole HOA may be hidden from the DOM but remains available to calculations.
10. Visible span colors restart at `span-color-0` for every pole and Midspan uses the paired span color.
11. UG has priority over PCO for blocked local export behavior.
12. Auto Calculate supports TOP COMM and LOW COMM.
13. Pole compliance outranks Midspan compliance during partial-solution ranking.
14. Auto Calculate never marks UG or PCO automatically.
15. User-entered HOA Change and Proposed values are not overwritten by the solver.

## Safe Extension Points

- Add project defaults in `core/project-config.js`.
- Add worksheet aliases and mappings in `import/excel-import.js`; put reusable cell/header conversions in `import/import-normalizers.js`.
- Add normalized fields in constructors and `normalizeState()` together.
- Add formulas or flagging rules in `domain/clearance-rules.js`.
- Add candidate generation/ranking behavior in `rules/auto-calculate-solver.js` without duplicating validations.
- Add Make Ready wording in `rules/mr-logic.js`.
- Add compact Web-to-plugin fields in `export/compact-autoproposed.js` and update schema, fixtures, and golden output together.
- Add post-render behavior as a small `js/ui/` module with `refresh(root)` and register it in `ui/runtime.js`.
- Do not add another workspace MutationObserver when the unified UI runtime can run the behavior.
