# Test Plan

This document defines regression coverage for the current behavior. The application does not yet include an automated test suite, so these cases are the source for future executable tests and current manual verification.

## Verification Levels

1. Syntax: every file in `js/` passes `node --check`.
2. State tests: load domain modules with `window = global`, build a small AppState, and call calculations directly.
3. Browser smoke tests: import, edit, Save/Load, Update Data, and export through the real UI.
4. Excel fixture tests: import known INTEC and Metronet workbooks and compare normalized state summaries.

Node is optional development tooling only. It is not a runtime requirement for GitHub Pages.

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
12. Fore/Back rows with different IDs keep each imported midspan on its own row.

## Persistence Cases

1. Save name removes `EXCEL_` and trailing date.
2. AutoProposed name removes `EXCEL_` and trailing date.
3. Save/Load round trip preserves user-owned fields.
4. Reload/new Excel causes the next Save to ask for a destination.
5. Load always opens the picker and starts near the previous handle when supported.
6. Update Data preserves matching movements and manual Proposed rows.
7. Update Data removes stale duplicates and recalculates before rendering.
8. Ctrl+Z restores the previous state-changing action.
9. Unsaved changes activate the browser leave-page warning.

## Browser Layout Smoke Tests

Verify at wide desktop, narrow desktop, and mobile widths:

- Configuration controls wrap without overlap.
- Proposed and comm tables remain horizontally scrollable.
- Checking or unchecking a control in a horizontally scrolled comm table preserves its horizontal position.
- Span labels remain on one line.
- Midspan color boxes align with their span rows.
- Left pole index opens, closes, and does not cover the top index unnecessarily.
- Floating calculator button opens and closes the calculator.
- In-app dialogs replace native prompt/confirm flows for supported actions.

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
