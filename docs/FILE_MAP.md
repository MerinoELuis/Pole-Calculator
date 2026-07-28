# File map

This is the quick ownership map for code changes. Detailed behavior remains in the existing architecture, business-rule and data-model documents.

## Runtime entry point

| File | Responsibility |
| --- | --- |
| `index.html` | Loads the static application in dependency order. |
| `css/styles.css` | Application layout, table presentation, status and span colors. |
| `css/auto-calculate-status.css` | SAFE/BEST AVAILABLE/CRITICAL result presentation. |
| `libs/xlsx.full.min.js` | Browser Excel parser. Third-party file; do not edit for application behavior. |

## Domain modules

| File | Global API | Responsibility |
| --- | --- | --- |
| `js/height-utils.js` | `HeightUtils` | Parse and format heights using integer inches. |
| `js/project-config.js` | `ProjectProfiles` | INTEC and Metronet/MidAm defaults and profile differences. |
| `js/state.js` | `AppStore` | Normalized mutable application state and identities. |
| `js/calculations.js` | `Calculations` | Pole limits, movements, midspans, Proposed values, End Drop and flagging. |
| `js/auto-calculate-solver.js` | `AutoCalculateSolver` | TOP/LOW COMM candidate search, best-partial ranking and Auto Calculate result metadata. |
| `js/midspan.js` | `MidspanLogic` | Midspan-oriented facade. |
| `js/mr-logic.js` | `MRLogic` | Ordered Make Ready generation. |
| `js/validations.js` | `Validations` | Broad data-integrity warnings. |
| `js/excel-import.js` | `ExcelImport` | Excel and saved JSON normalization. |
| `js/excel-review.js` | `ExcelReview` | Read-only workbook audit. |
| `js/json-export.js` | `ProjectExport` | General JSON downloads. |
| `js/compact-autoproposed.js` | `CompactAutoProposed` | Compact Web-to-plugin export, UG/PCO safety and fiber configuration UI. |
| `js/floating-calculator.js` | `FloatingCalculator` | Feet/inches expression calculator and quote normalization. |
| `js/app.js` | Browser handlers | Main rendering, edit dispatch, undo, persistence, dialogs and targeted refreshes. |

## UI modules

| File | Global API | Responsibility |
| --- | --- | --- |
| `js/ui/dom-contract.js` | `UiDomContract` | Adds stable `data-component`, `data-pole-id`, `data-span-id` and `data-column` attributes after render. |
| `js/ui/comm-table.js` | `CommTableUI` | Hides the visual `Other Pole HOA` column while preserving calculation data. |
| `js/ui/span-colors.js` | `SpanColorUI` | Restarts the visible span color sequence per pole and synchronizes each Midspan row. |
| `js/ui/auto-calculate-status.js` | `AutoCalculateStatusUI` | Shows SAFE, BEST AVAILABLE, CRITICAL, MANUAL and SKIPPED results and enables Auto Calculate for TOP/LOW COMM. |
| `js/ui/runtime.js` | `PoleCalculatorUI` | Runs all UI modules through one queued refresh and one MutationObserver. |

## Contracts and tests

| Path | Responsibility |
| --- | --- |
| `schemas/autoproposed.schema.json` | Formal structural contract shared by Web and the O-Calc plugin. |
| `tests/run-all-tests.js` | Dependency-free syntax, JSON and regression test runner. |
| `tests/fixtures/` | Reusable normalized states and exported payload examples. |
| `tests/expected/` | Golden expected outputs. |
| `.github/workflows/tests.yml` | Runs `npm test` for pushes and pull requests. |
