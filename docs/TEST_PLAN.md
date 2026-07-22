# Test Plan

This document defines regression coverage for the current behavior. Dependency-free Node tests cover Excel Review and Update Data merging; the remaining cases are the source for broader future automation and current manual verification.

## Verification Levels

1. Syntax: every file in `js/` passes `node --check`.
2. State tests: load domain modules with `window = global`, build a small AppState, and call calculations directly.
3. Browser smoke tests: import, edit, Save/Load, Update Data, and export through the real UI.
4. Excel fixture tests: import known INTEC and Metronet workbooks and compare normalized state summaries.

Node is optional development tooling only. It is not a runtime requirement for GitHub Pages.

Run the dependency-free Excel Review regression test with:

```powershell
node tests/excel-review.test.js
node tests/update-merge.test.js
node tests/pole-identity.test.js
node tests/equipment-actions.test.js
```

These tests cover decimal-height equivalence, exact Low Power source selection, Fore/Back counting with Other rows, generated-MR readiness, the INTEC wire rules, non-destructive Update Data reconciliation, and canonical pole identity across STEEL/UG/PCO suffix variants.

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
4. Both endpoints move: both half-movements are applied once.
5. Same owner on two spans: only the physical matching span affects the row.
6. Same owner with two Wire IDs: exact Wire ID wins remote matching.
7. REF row: calculated midspan and Other Pole HOA remain blank.
8. Other with its own midspan: row is not REF and calculates normally.

### Comm Flagging

1. Comm above Max Height on Pole.
2. Comm above Max Height at MS.
3. Comm below Environment minimum.
4. Different owners closer than Pole Comm-comm.
5. Same owner closer than Pole Bolt-bolt.
6. New bolt too close to a previous Existing HOA point.
7. Same-owner Service Drop at the exact same bolt is allowed.
8. Two comms on the same span closer than Midspan Comm-comm.
9. Similar comms on different spans are not compared for Comm-comm MS.
10. Pole order and midspan order reversed produces crossing flagging.

### Proposed

1. Proposed above Max Height on Pole flags.
2. Top Comm Proposed below Top Comm flags.
3. Low Comm Proposed above Low Comm flags.
4. Proposed too close to an effective comm flags Comm-comm.
5. Proposed too close to an old attachment point flags Bolt-bolt.
6. Proposed can reuse an exact HOA only after that comm moves away.
7. Two Proposed attachments keep Bolt-bolt spacing.
8. Proposed MS references only the same physical pole pair.
9. INTEC low-power adjustment and reminder follow profile settings.
10. Metronet leaves the low-power violation visible and does not add the INTEC reminder.

### End Drop

1. Connected pole has a forward Proposed: it populates Next Pole Proposed.
2. Connected pole has several Proposed rows: the first non-additional row is used.
3. Connected terminal pole has standalone Proposed: it is used.
4. No connected Proposed: Next Pole Proposed and End Drop remain blank.
5. Manual Next Pole Proposed remains local and does not modify the connected pole.

### Auto Calculate

1. Low Comm mode disables the button.
2. Healthy Top Comm stack with space above proposes without moving comms.
3. No space above causes a safe downward stack.
4. Midspan power violation lowers the necessary comm group.
5. Different owners use Comm-comm spacing; same owners use Bolt-bolt.
6. A remote movement can reduce a local move on a later pass.
7. Iteration stops when the movement signature converges.
8. Repeated state and max-pass limit stop safely.
9. Existing manual HOA changes are not overwritten.
10. A candidate introducing a new violation is rejected.

### Make Ready Controls

1. Transfer to New Pole appears once per grouped comm, synchronizes all related spans, generates one movement line, and includes `with DG` when any related span has DG checked.
2. INTEC Re-sag Service Drop raises the effective low drop midspan to `15'6"`, preserves other clearance checks, and is not applied for Metronet.
3. A transfer without HOA Change uses Existing HOA; multiple heights for the same owner produce one ascending-height instruction.
4. An adjacent Backspan UG creates one relation/direction line and one riser 12 inches below the primary Proposed.
5. An active INTEC UG pole replaces normal MR with the single final overload instruction.
6. MidAm automatically activates and locks Streetlight Ground, generates its MR, preserves both Streetlight clearances, and does not invent a new Low Power height.
7. Transformer Redress and Power Riser Raise update effective Low Power when they move the limiting equipment.
8. Disabling an equipment action restores imported Low Power.
9. Update Data preserves equipment actions and refreshes the Low Power baseline from non-empty Excel data.

## Import Cases

1. Flexible Low Power header containing `Low Power Attachment`.
2. Span links through title and through Linked Collection ID.
3. Missing linked pole creates a stable editable Unknown pole.
4. Bearing maps to the correct cardinal direction.
5. Environment defaults to the configured clearance, including None at `15'6"`.
6. Power wires do not become primary comm rows.
7. Owner comes from Span.Wire Owner.
8. Communication Drop in Size sets Service Drop.
9. Anchor.Guys matches pole, owner, and exact attachment height for DG.
10. INTEC Self-Supporting Fiber remains visible as POF.
11. Metronet markers select the Metronet profile.
12. Selecting Metronet exposes the WI selector and applies `MidAm` as Proposed owner.
13. `UTILITY > MidAm` automatically selects Metronet/MidAm and applies its crossing and streetlight defaults.
14. A MidAm Back Span with its own imported midspan responds to half the movement at each endpoint even when reciprocal Span/Wire IDs differ.
15. Wecom and MidAm Proposed MS uses the highest same-span comm midspan plus `12\"`; without a comm midspan it subtracts sag after rounding span length to the nearest `50 ft`, while manual O-CALC remains authoritative.
16. Fore/Back rows with different IDs keep each imported midspan on its own row.
17. Equipment imports Utility/Power Streetlights, Transformers, and Risers for both profiles while excluding communication risers.
18. Supported Equipment can lower Max Height on Pole using the active project rules.

## Persistence Cases

1. Save name removes `EXCEL_` and trailing date.
2. AutoProposed name removes `EXCEL_` and trailing date.
3. Save/Load round trip preserves user-owned fields.
4. Reload/new Excel causes the next Save to ask for a destination.
5. Load always opens the picker and starts near the previous handle when supported.
6. Update Data preserves matching movements and manual Proposed rows.
7. Update Data removes stale duplicates and recalculates before rendering.
8. Empty replacement cells retain prior values only on matched entities; omitted rows without user work do not recreate empty spans or comms.
9. Ctrl+Z restores the previous state-changing action.
10. Unsaved changes activate the browser leave-page warning.

## Excel Review Cases

1. Successful raw import and Update Data each replace the prior review without changing the active tab.
2. Re-run Review uses current Calculator movements, Proposed values, midspans, and generated MR.
3. A Collection row with no Id errors; a missing Id header creates a global error.
4. Low Power checks accept `Lowest Power.display` and the legacy `Low Power Attachment.display` source.
5. Missing Year Installed warns for INTEC but is not required for Metronet/MidAm.
6. MidAm derives the expected Sequence from the first Id block (`000` or `000A`); Sequence normalizes numeric `58` to `058` and suffix value `51b` to `051B`, then must equal the ID-derived value exactly.
7. MidAm Collection Owner `UTILITY > MidAm` passes; empty errors; another populated owner warns.
8. Exactly one Fore is required; one Back passes, zero Back warns, more than one Back errors, and Other rows do not affect those counts.
9. Incorrect reciprocal Fore/Back relationships warn and mismatched reciprocal Environments error.
10. Empty Fore/Back Linked Collection.Title produces a low-level warning. Populated external-job links and empty Other links do not warn.
11. An incomplete Anchor row produces one error listing all missing required fields; a complete row passes. Anchor and Anchor.Guys remain independent snapshots.
12. INTEC DAVIT, owner, utility owner, and insulator rules use raw Span.Wire rows; duplicate wires are not reviewed.
13. A workbook without Make Ready or Comm Transfers remains HOA-only: Final Review is N/A and produces no Final findings even when Calculator-derived work exists.
14. Calculator-only work errors; Excel-only final data warns.
15. Proposed and final Proposed midspan compare through integer inches, including decimal-feet equivalence.
16. UG requires Underground construction and one valid UG resolution note without requiring aerial attachment heights; `Unable to attach due to <specific reason>` passes, while an unresolved placeholder produces one error.
17. Expected structured transfers match normalized owner aliases and heights, including CenturyLink/CTL/TELCO equivalence.
18. Attachment Size is Not applicable when Calculator lacks reliable per-Proposed fiber/messenger identity.
19. Excel Review never emits clearance, Pole Type, loading, AutoQC, or O-Calc checks.
20. Review cards stay in natural sequence order rather than moving errors ahead of lower-numbered poles.
21. Duplicate identical MR rows and model-only slack/anchor/split instructions do not create unmatched-instruction warnings.
22. A height or required-direction difference, including UG direction, pairs the Calculator and Excel instructions into one mismatch instead of separate missing and additional results.
23. Ignore keeps a finding visible in gray but removes it from phase and summary status; Restore activates it again.
24. Save/Load and Update Data preserve ignored finding keys.
25. The Excel Review tab badge counts poles with active Error or Warning status and excludes ignored findings.

## Browser Layout Smoke Tests

Verify at wide desktop, narrow desktop, and mobile widths:

- Configuration controls wrap without overlap.
- Proposed and comm tables remain horizontally scrollable.
- Checking or unchecking a control in a horizontally scrolled comm table preserves its horizontal position.
- Span labels remain on one line.
- Midspan color boxes align with their span rows.
- Left pole index opens, closes, and does not cover the top index unnecessarily.
- Floating calculator button opens and closes the calculator.
- Floating up-arrow button scrolls to the top without opening the calculator.
- The one-line calculator converts a height by itself and evaluates `+`/`-` expressions without layout overlap.
- In-app dialogs replace native prompt/confirm flows for supported actions.
- Excel Review summaries wrap cleanly and each pole's problems expand without overlap.

## Definition of Done for a Calculation Change

A business-rule change is complete when:

1. The rule is updated in `BUSINESS_RULES.md`.
2. The data fields are updated in `DATA_MODEL.md` when necessary.
3. Import/export contracts are updated when their payload changes.
4. Relevant tests above pass or are updated with the approved behavior.
5. All JavaScript passes syntax checks.
6. Save/Load round trip is not broken.
7. INTEC and Metronet behavior is checked separately.

## Automation Still Needed

The highest-priority future work is a small, dependency-light regression harness for HeightUtils and Calculations, followed by browser automation for the critical import/edit/save workflow. Test implementation must not add a runtime build requirement to the GitHub Pages application.
