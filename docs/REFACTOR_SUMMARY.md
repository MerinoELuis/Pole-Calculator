# AI efficiency refactor summary

## Safety

- The reorganization work is kept on the active feature branch; `main` is not modified.
- `main` was not modified.
- Pre-refactor backup: `backup/compact-autoproposed-json-before-ai-refactor` at commit `837db33bb3dca88f0b8e1660846b8dc4610b904c`.

## Added

- One operational AI index: `AGENTS.md`.
- File ownership, decisions, UI runtime, test automation and AutoProposed contract documentation.
- Stable UI DOM identities through `js/ui/dom-contract.js`.
- Dedicated communication table and per-pole span color modules.
- One post-render UI coordinator in `js/ui/runtime.js`.
- Formal compact AutoProposed JSON Schema.
- Reusable fixtures and golden expected output.
- One dependency-free `npm test` command.
- GitHub Actions workflow that executes the same command.

## Integrated

UG/PCO export and Make Ready protection now lives in `js/export/compact-autoproposed.js` rather than a separate wrapper script.

The runtime is now grouped by responsibility:

```text
js/
├─ core/      store, state compatibility, history, persistence, project config, deployment version
├─ domain/    height math, pole mode, comm movement, clearance calculations, Midspan engine, graph helpers
├─ rules/     Auto Calculate solver, Make Ready, validations
├─ import/    import normalizers, Excel import/review, Update Data reconciliation
├─ export/    Excel, JSON, and compact AutoProposed export
└─ ui/        app coordinator, pole renderer, controls, dialogs, theme, mobile keyboard, post-render modules
```

`index.html` loads the modules in dependency order. The Midspan engine now owns the baseline-plus-half-movement calculation; `clearance-rules.js` remains the central orchestrator for clearance evaluations, persistence, and flagging so existing business-rule behavior stays consistent.

## Removed

- `js/compact-autoproposed-ug-guard.js`
- `js/comm-table-columns.js`
- `js/span-color-reset.js`
- Their superseded tests

## Intentionally retained

The current browser coordinator is `js/ui/app.js`. Shared history, persistence, dialogs, theme, mobile keyboard handling, controls, and post-render behaviors have dedicated modules. The pole renderer and the integrated clearance calculation module remain the largest cohesive modules; future extraction from those should preserve their documented contracts and be covered by regression tests.

## Verification

The complete dependency-free local suite passes with:

```bash
npm test
```
