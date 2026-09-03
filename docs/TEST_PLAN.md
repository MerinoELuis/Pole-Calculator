# Test Plan

This document defines regression coverage for the current behavior. Dependency-free Node tests cover domain rules, the best-available Auto Calculate solver, compact export, fixtures, schema contracts, UI helper modules, runtime ordering, the floating calculator, Excel Review, Update Data merging, canonical pole identities, and equipment actions. Browser smoke testing remains required for full import/edit/save workflows.

## Run the complete suite

```powershell
npm test
```

The command:

1. Runs `node --check` for every file under `js/`.
2. Parses `package.json`, schemas, fixtures, and golden JSON files.
3. Runs every `tests/**/*.test.js` regression test.
4. Returns a non-zero exit code when any check fails.

For syntax and JSON checks without regression tests:

```powershell
npm run test:syntax
```

Node is optional development tooling only. It is not a runtime requirement and does not add a GitHub Pages build step.

## Verification Levels

1. Syntax: every file in `js/` passes `node --check`.
2. Contract: schemas, fixtures, and expected JSON obey documented field restrictions.
3. State tests: domain modules load with a controlled `window` and call calculations/exporters directly.
4. Solver tests: TOP/LOW candidate generation, ranking, partial retention, and source locking are checked independently.
5. UI module tests: DOM identities, table-column removal, per-pole colors, and refresh order are verified without a browser framework.
6. Browser smoke tests: import, edit, Auto Calculate, Save/Load, Update Data, and export through the real UI.
7. Excel fixture tests: import representative INTEC and Metronet workbooks when available.

## Automated Files

The runner discovers every file ending in `.test.js`. Important groups include:

- `auto-calculate-solver.test.js`
- `auto-calculate-solver-integration.test.js`
- `excel-review.test.js`
- `update-merge.test.js`
- `pole-identity.test.js`
- `equipment-actions.test.js`
- `compact-autoproposed.test.js`
- `compact-autoproposed-blocked-poles.test.js`
- `compact-autoproposed-golden.test.js`
- `schema-contract.test.js`
- `floating-calculator-input.test.js`
- `floating-calculator-ui.test.js`
- `ui-dom-contract.test.js`
- `ui-comm-table.test.js`
- `ui-span-colors.test.js`
- `ui-runtime.test.js`
- `index-runtime-order.test.js`

Exact input and expected-output data belong in `tests/fixtures/` and `tests/expected/`.

## Required Calculation Cases

### Heights

| Input | Expected |
| --- | --- |
| `20` | `20'` |
| `20'` | `20'` |
| `20'6"` | `20'6"` |
| `20.5` | `20'6"` |
| invalid text | rejected/blank derived value |

### Existing Comm Midspan

1. No endpoint movement: calculated midspan equals imported midspan.
2. Local pole lowers 12 inches: midspan lowers 6 inches.
3. Remote pole raises 12 inches: local displayed midspan rises 6 inches.
4. Both endpoint movements are applied once.
5. Same owner on two spans: only the physical matching span affects the row.
6. Same owner with two Wire IDs: exact Wire ID wins remote matching.
7. REF row: calculated midspan and Other Pole HOA remain blank.
8. Other with its own midspan calculates normally.

### Comm Flagging

1. Comm above Max Height on Pole.
2. Comm above Max Height at MS.
3. Comm below Environment minimum.
4. Different owners closer than Pole Comm-comm.
5. Same owner closer than Pole Bolt-bolt.
6. A same-owner Service Drop at zero or nonzero separation does not flag Bolt-bolt; a DG-only row still flags.
7. New bolt too close to a previous Existing HOA point.
8. Same-owner Service Drop at the exact same bolt is allowed.
9. Two comms on the same span closer than Midspan Comm-comm.
9. Similar comms on different spans are not compared for Comm-comm MS.
10. Reversed pole/midspan order produces crossing flagging.

### Proposed

1. Proposed above Max Height on Pole flags.
2. TOP COMM Proposed below Top Comm flags.
3. LOW COMM Proposed above Low Comm flags.
4. Proposed too close to an effective comm flags Comm-comm.
5. Proposed too close to an old attachment point flags Bolt-bolt.
6. Proposed can reuse an exact HOA only after that comm moves away.
7. Two Proposed attachments keep Bolt-bolt spacing.
8. Proposed MS references only the same physical pole pair.
9. INTEC low-power adjustment/reminder follows profile settings.
10. Metronet leaves low-power violation visible without the INTEC reminder.

### End Drop

1. Connected pole forward Proposed populates Next Pole Proposed.
2. Several Proposed rows: first non-additional row is used.
3. Connected terminal pole standalone Proposed is used.
4. No connected Proposed leaves Next Pole Proposed and End Drop blank.
5. Manual Next Pole Proposed remains local.

### Best-available Auto Calculate

1. The button is available in TOP COMM and LOW COMM when poles are loaded.
2. TOP COMM includes `Top Comm + Pole Comm-comm` as the ideal candidate.
3. Healthy TOP COMM tries the ideal direct candidate first; a TOP COMM arrangement with an existing Midspan violation tries Max Height first.
4. A direct SAFE TOP COMM result stops after one candidate without moving comms.
5. LOW COMM includes `Low Comm - Pole Comm-comm` as the ideal candidate.
6. Healthy TOP COMM with space above proposes at the ideal without moving comms.
7. When TOP COMM does not fit, the smallest permitted downward stack is evaluated.
8. When LOW COMM needs room, the smallest permitted upward stack is evaluated.
9. Different owners use Comm-comm spacing; same owners use Bolt-bolt.
10. User-entered HOA Change remains fixed.
11. User-entered Proposed remains fixed; AUTO Proposed may be replaced later.
12. Candidate heights include one-inch positions around the ideal and important boundaries.
13. `SAFE` outranks all partial arrangements.
14. `BEST_AVAILABLE` means pole rules pass while Midspan remains in violation.
15. Any pole-compliant result outranks a result where Midspan passes but the pole fails.
16. Within the same category, smaller pole/Midspan violation count and shortfall outrank movement cost.
17. After violations tie, fewer moved comm groups and fewer total inches win.
18. `CRITICAL` retains the best evaluated arrangement when no pole-compliant candidate exists.
19. A reciprocal span's Midspan issue affects the arrangement, but a remote endpoint's independent pole issue is scored on that endpoint.
20. Existing UG/PCO poles are skipped.
21. Non-SAFE INTEC results recommend UG/PCO review but never activate either option.
22. Every pole receives one initial evaluation; only affected poles with remaining violations or removable automatic movements are retried, with a three-attempt cap.
23. Streetlight Ground/Raise, Transformer Secure, Power Riser Raise/Secure, Re-sag, and Transfer remain user-controlled but active values participate in validation.
24. Midspan recovery stacks comms downward from the highest legal positions while reusing or clearing every existing bolt point.
25. P15 recovery retains Proposed at `20'10"` when the same residual Midspan issues remain, and P16 recovery uses `23'2"` while reducing the real debug case from three Midspan issues to one without adding a pole violation.
26. A pole-compliant result reached within the direct progressive candidates stops before the broad fallback height scan.
27. A stale automatic P17-style plan at `2'4"` is discarded before retry; recovery evaluates `20'4"` and the nearest legal `20'2"` rather than treating the old AUTO value as progressive.
28. With Proposed `20'2"`, P17 stacks at `19'2"`, `18'`, and `17'8"`; a zero-only propagated floor is never generated when no comm has a real Midspan minimum.
29. The latest Auto Calculate run exposes a Debug-only trace containing each evaluated candidate, its comm plan, analysis, ranking decision, stop reason, and selected plan.

### Make Ready Controls

1. Transfer to New Pole appears once per grouped comm and includes `with DG` when appropriate.
2. INTEC Re-sag Service Drop raises the effective low drop midspan to `15'6"` without overwriting the imported baseline.
3. Service Drop relocation text and compact movement payloads never include `with DG`/`DG`.
4. Transfer without HOA Change uses Existing HOA; multiple heights are ordered and joined correctly.
5. Adjacent UG relation reason/direction and riser precedence remain correct.
6. Partial Update Data keeps omitted poles in Excel Review and Pole Type Check.
7. Active INTEC UG replaces normal MR with the six-line editable template.
8. MidAm Streetlight Ground remains mandatory and locked.
9. Transformer Secure and Power Riser Raise update effective Low Power when limiting; Power Riser Secure generates a drip-loop instruction independently.
10. Disabling optional equipment work restores imported baseline.
10. Update Data preserves equipment actions and refreshes baseline values.
11. INTEC Streetlight Raise accepts at most 12 inches above Attachment Height.

## Compact AutoProposed Cases

1. Normal poles keep local movements and aerial proposals.
2. Normal-to-PCO remains aerial from the normal pole.
3. Normal-to-UG exports the connection as UG without aerial attachment fields.
4. PCO removes local movements/terminal/proposal fields while retaining geometry.
5. PCO local span is not converted to UG.
6. Fully UG removes local movements/terminal/proposal fields and retains geometry with `ug: true`.
7. UG takes priority over PCO.
8. Blocked pole with no remaining geometry is omitted.
9. UG/PCO replacement MR does not regain regular movement lines.
10. Normal fixture matches golden compact JSON exactly.
11. Fixtures obey integer-inch schema constraints.

## UI Runtime Cases

1. Every pole card receives stable `data-pole-id` and `data-component`.
2. Proposed/communication rows receive `data-span-id`.
3. Table headers/cells receive normalized `data-column`.
4. Other Pole HOA is found by stable key before label fallback.
5. Removing Other Pole HOA does not delete state/calculation fields.
6. Visible span colors restart at turquesa for each pole.
7. Hidden spans do not consume visible color positions.
8. Synthetic Proposed resolves to physical `sourceSpanId`.
9. Blank/REF Midspan uses the paired span color.
10. Refresh order is DOM contract, comm cleanup, optional re-annotation, then colors.
11. Only the unified UI runtime observes the workspace.
12. `index.html` loads the solver before `app.js`, the color module before runtime, and no superseded patch scripts.

## Import Cases

1. Flexible Low Power header matching.
2. Span links through title and Linked Collection ID.
3. Missing linked pole creates stable Unknown pole.
4. Bearing maps to correct direction.
5. Environment defaults to configured clearance.
6. Power wires do not become comm rows.
7. Owner comes from Span.Wire Owner.
8. Communication Drop marks Service Drop.
9. Anchor.Guys matches pole, owner, and exact attachment height for DG.
10. INTEC Self-Supporting Fiber remains visible as POF.
11. Metronet markers select Metronet profile.
12. Metronet exposes WI and applies MidAm.
13. `UTILITY > MidAm` applies MidAm defaults.
14. MidAm measured Back Span responds to both endpoints even with different IDs.
15. INTEC Proposed MS stays blank without O-CALC; MidAm uses same-span comm MS or sag fallback.
16. Fore/Back records keep their own imported midspans.
17. Equipment imports supported Utility/Power rows and excludes communication risers.
18. Supported Equipment can lower Max Height on Pole.

## Persistence Cases

1. Save/AutoProposed names remove `EXCEL_` and trailing date.
2. Save/Load round trip preserves user-owned fields.
3. Reload/new Excel asks for a new Save destination.
4. Load always opens picker and uses prior handle only as a location hint.
5. Update Data preserves matching movements and manual Proposed rows.
6. Update Data removes stale duplicates and recalculates.
7. Empty replacement cells retain prior values only on matched entities.
8. Ctrl+Z restores the previous state-changing action.
9. Unsaved changes activate the browser leave warning.
10. AUTO source metadata survives normal state serialization and is replaced by later solver runs; user edits clear source markers.

## Excel Review Cases

Excel Review coverage continues to require Collection identity, Fore/Back relationships, Environment, Anchor completeness, INTEC wire rules, HOA-only workbooks, Proposed/final Midspan comparison, UG resolution, structured transfers, ignored findings, natural order, supplemental model instructions, and tab badge counts.

## Browser Layout Smoke Tests

Verify at wide desktop, narrow desktop, and mobile widths:

- Configuration controls wrap without overlap.
- Proposed and comm tables remain horizontally scrollable.
- Horizontal table position survives checkbox changes.
- Span labels remain on one line.
- Other Pole HOA is not visible.
- Each pole begins visible colors at turquesa.
- Midspan colors match paired spans, including blank/REF.
- Auto Calculate follows the configured availability rules after import.
- UG/PCO is never selected automatically.
- Pole index, floating calculator, dialogs, and Excel Review remain usable.

## Definition of Done

A business-rule or architecture change is complete when:

1. Rules are updated in `BUSINESS_RULES.md` or `DECISIONS.md`.
2. Data fields are updated in `DATA_MODEL.md` when necessary.
3. Import/export contracts and schema are updated when payload changes.
4. Relevant fixtures/golden outputs are updated only for approved behavior.
5. `npm test` passes.
6. Save/Load round trip is not broken.
7. INTEC and Metronet behavior is checked separately when profile rules are involved.
8. A browser smoke test is completed for DOM, import, solver, or persistence changes.

## Automation Still Needed

The remaining highest-priority automation is browser-level testing for critical import/edit/Auto Calculate/save workflows using representative INTEC and Metronet workbooks. Future browser automation must not add a runtime build requirement to the GitHub Pages application.
