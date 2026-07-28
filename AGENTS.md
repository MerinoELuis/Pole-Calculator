# AI working guide

This file is an operational index. The project documentation remains authoritative.

## Repository workflow

- Static web repository: `MerinoELuis/Pole-Calculator`.
- Do not modify `main` unless the user explicitly requests it.
- Use the active feature branch supplied by the user; current refactor branch: `compact-autoproposed-json`.
- Preserve the dependency order in `index.html`.
- When changing a file, provide the complete updated file rather than a partial patch.
- Do not open a pull request unless the user asks for one.

## Required verification

Run all dependency-free checks with:

```bash
npm test
```

The command performs JavaScript syntax checks, JSON parsing, fixture checks and every `tests/**/*.test.js` regression test.

## Documentation index

- Architecture and data flow: `docs/ARCHITECTURE.md`
- Current file ownership: `docs/FILE_MAP.md`
- Business rules: `docs/BUSINESS_RULES.md`
- Data model and identities: `docs/DATA_MODEL.md`
- Import/export behavior: `docs/IMPORT_EXPORT.md`
- AutoProposed structural contract: `schemas/autoproposed.schema.json`
- Approved implementation decisions: `docs/DECISIONS.md`
- Regression expectations: `docs/TEST_PLAN.md`
- UI post-render modules: `docs/UI_RUNTIME.md`
- Auto Calculate search and ranking: `docs/AUTO_CALCULATE_SOLVER.md`

## Architectural invariants

- Convert heights to integer inches before arithmetic.
- Keep imported values separate from calculated values.
- Never borrow a midspan from an unrelated physical span.
- UG has priority over PCO.
- A completely UG or PCO pole cannot export local moves or `terminalHoa`.
- PCO local spans remain geometry-only; they are not converted to UG.
- Span colors restart from `span-color-0` for each pole and the Midspan color must match its span.
- `Other Pole HOA` remains available to calculations but is not displayed in the communication table.
- Auto Calculate must prefer pole compliance over Midspan compliance and retain the best partial aerial result when no SAFE result exists.
- Auto Calculate never marks UG or PCO automatically.

## Preferred extension points

- Project differences: `js/project-config.js`
- Calculations and flagging: `js/calculations.js`
- Auto Calculate candidate search and ranking: `js/auto-calculate-solver.js`
- Make Ready wording: `js/mr-logic.js`
- Compact AutoProposed export: `js/compact-autoproposed.js`
- Post-render DOM identity and table behavior: `js/ui/`
