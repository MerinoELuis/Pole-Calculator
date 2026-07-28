# UI runtime modules

`app.js` remains the main renderer and edit controller. Small post-render responsibilities are isolated in `js/ui/` so table behavior no longer depends on independent patch scripts with separate observers.

## Refresh order

`PoleCalculatorUI.runRefresh()` executes:

1. `UiDomContract.apply()`
2. `CommTableUI.refresh()`
3. `UiDomContract.apply()` again only when a column was removed
4. `SpanColorUI.refresh()`
5. `AutoCalculateStatusUI.refresh()`

The order matters. Stable identities are added before behavior reads them, span colors are assigned after the communication table reaches its final visible structure, and Auto Calculate status cards are rendered last so they do not affect table identity or color pairing.

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
