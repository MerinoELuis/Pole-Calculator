# UI runtime modules

`app.js` remains the application coordinator and edit controller. Rendering responsibilities are isolated in `js/ui/` modules; post-render DOM normalization remains coordinated by `runtime.js` so UI modules do not create independent observers.

Excel Review result markup and warning controls are owned by `ExcelReviewRenderer`. It is created by `PoleWorkspaceRenderer` with the state, selection, history, persistence, and escaping dependencies it needs. The renderer methods remain available through `PoleWorkspaceRenderer` so existing app wiring does not change.

Fiber planning is owned by `FiberTabRenderer`. It reads normalized Make Ready fiber references and connected physical spans, allows manual fiber rows when the workbook has none, and writes operator assignments to `state.fiberAssignments` plus global fiber-size/MST-port defaults to `state.settings`. Pole rows and the add-fiber pole selector use natural numeric ID ordering (P1, P2, …, P10), independent of import order. `CompactAutoProposed` consumes these resolved rows when assigning fiber types to exported spans. Diameter is global per fiber type; MST exports its global diameter and selected port count. New/Overlash is initialized from each imported Excel row and remains individually editable; it is not a global default. The view renders during the app's normal render pass; assignments and defaults are included in regular project Save/Load and preserved by Update Data.

## Refresh order

`PoleCalculatorUI.runRefresh()` executes:

1. `UiDomContract.apply()`
2. `CommTableUI.refresh()`
3. `UiDomContract.apply()` again only when a column was removed
4. `SpanColorUI.refresh()`

The order matters. Stable identities are added before behavior reads them, and span colors are assigned after the communication table reaches its final visible structure.

## Stable attributes

The DOM contract adds:

- `data-component="pole-card"`
- `data-pole-id="..."`
- `data-span-id="..."`
- `data-column="..."`
- component identifiers for proposed, communication, power and equipment tables

These attributes are presentation metadata only. They do not modify AppState.

## Observer policy

Only `js/ui/runtime.js` owns the post-render MutationObserver. It watches `#polesOverview` for child-list changes, queues one microtask and runs all modules together. UI modules must not create additional observers.

## Adding a UI module

A new module should:

1. Publish a small global API through an IIFE.
2. Expose `refresh(root)`.
3. Read stable data attributes where possible.
4. Avoid changing AppState unless the module explicitly owns an input workflow.
5. Add a dependency-free Node test.
6. Be registered in `PoleCalculatorUI.runRefresh()` and documented in `FILE_MAP.md`.
