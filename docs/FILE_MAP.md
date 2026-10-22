# File map

This is the quick ownership map for code changes. Detailed behavior remains in the architecture, business-rule, data-model, and import/export documents.

## Runtime and core

| File | Global API | Responsibility |
| --- | --- | --- |
| `index.html` | — | Loads the static application in dependency order. |
| `css/styles.css` | — | Application layout, tables, status and span colors. |
| `libs/xlsx.full.min.js` | `MiniXLSX` | Third-party browser Excel parser; do not edit for application behavior. |
| `js/core/deployment-version.js` | `DeploymentVersion` | Build and deployment metadata shown in the page header. |
| `js/core/project-config.js` | `ProjectProfiles` | INTEC and Metronet/MidAm defaults and profile differences. |
| `js/core/store.js` | `AppStore` | Normalized application state, constructors, identities, and graph accessors. |
| `js/core/state-compat.js` | `StateCompatibility` | Compatibility adapters for older JSON/imported state. |
| `js/core/history.js` | `AppHistory` | Undo/redo and movement trace generation for Debug exports. |
| `js/core/persistence.js` | `AppPersistence` | Save/load, file handles, dirty-state tracking, and safe filenames. |

## Domain and rules

| File | Global API | Responsibility |
| --- | --- | --- |
| `js/domain/height-math.js` | `HeightUtils` | Parse and format heights using integer inches. |
| `js/domain/pole-resolution.js` | `PoleResolution` | Normalize AERIAL/UG/PCO mode and its priority rules. |
| `js/domain/comm-movements.js` | `CalculationAutoMovementModule` | Plan and apply Auto Calculate communication movements through injected calculation helpers. |
| `js/domain/clearance-rules.js` | `Calculations` | Integrated calculation orchestrator: pole/span limits, Midspan, Proposed, End Drop, recalculation, and flagging. |
| `js/domain/midspan-engine.js` | `MidspanEngine` (`MidspanLogic` alias) | Midspan baseline-plus-endpoint-movements calculation and span-oriented APIs. |
| `js/domain/graph.js` | `PoleGraph` | Graph utilities; not currently loaded in the browser script sequence. |
| `js/rules/mr-logic.js` | `MRLogic` | Ordered Make Ready generation. |
| `js/rules/auto-calculate-solver.js` | `AutoCalculateSolver` | TOP/LOW COMM candidate search and result ranking. |
| `js/rules/validations.js` | `Validations` | Broad data-integrity warnings outside table-level calculation flagging. |

## Import and export

| File | Global API | Responsibility |
| --- | --- | --- |
| `js/import/import-normalizers.js` | `ImportNormalizers` | Reusable worksheet/header/identifier/cell/height normalization. |
| `js/import/excel-import.js` | `ExcelImport` | Convert Excel worksheets and saved JSON into normalized state. |
| `js/import/excel-review.js` | `ExcelReview` | Read-only workbook audit against calculated state and Make Ready. |
| `js/import/update-data-reconciler.js` | `UpdateDataReconciler` | Non-destructive merge of refreshed workbook data with operator work. |
| `js/export/excel-export.js` | `ExcelExport` | Produce workbook downloads. |
| `js/export/json-export.js` | `ProjectExport` | Produce regular project and diagnostic JSON. |
| `js/export/compact-autoproposed.js` | `CompactAutoProposed` | Produce compact plugin data and apply UG/PCO export safety. |

## UI

| File | Global API | Responsibility |
| --- | --- | --- |
| `js/ui/app.js` | browser coordinator | Initialize modules and coordinate events, history, persistence, dialogs, and refreshes. |
| `js/ui/pole-renderer.js` | `PoleWorkspaceRenderer` | Render pole cards, communications, Proposed tables, review panels, and targeted refreshes. |
| `js/ui/controls.js` | `PoleActions` | Handle action buttons, comm edits, and Proposed span controls. |
| `js/ui/dialogs.js` | `AppDialogs` | Shared confirmation, prompt, and message dialogs. |
| `js/ui/theme.js` | `ThemeControl` | Resolve and persist light/dark theme. |
| `js/ui/mobile-keyboard.js` | `MobileKeyboardControl` | Mobile viewport and keyboard-related input behavior. |
| `js/ui/floating-calculator.js` | `FloatingCalculator` | Feet/inches expression calculator and quote normalization. |
| `js/ui/dom-contract.js` | `UiDomContract` | Add stable component, pole, span, and column identities after render. |
| `js/ui/comm-table.js` | `CommTableUI` | Hide visual Other Pole HOA while preserving calculation data. |
| `js/ui/span-colors.js` | `SpanColorUI` | Restart visible span colors per pole and pair each Midspan with its span. |
| `js/ui/runtime.js` | `PoleCalculatorUI` | Coordinate post-render behavior with one queued refresh and one workspace observer. |

## Contracts and tests

| Path | Responsibility |
| --- | --- |
| `schemas/autoproposed.schema.json` | Formal structural contract shared by Web and the O-Calc plugin. |
| `tests/run-all-tests.js` | Dependency-free syntax, JSON, and regression-test runner. |
| `tests/fixtures/` | Reusable normalized states and exported payload examples. |
| `tests/expected/` | Golden expected outputs. |
| `.github/workflows/tests.yml` | Runs `npm test` for pushes and pull requests. |
