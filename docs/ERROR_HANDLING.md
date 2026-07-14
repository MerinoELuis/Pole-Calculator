# Error Handling and Diagnostics

The application distinguishes operational errors, missing data, clearance flagging, and broad warnings. These categories should not be merged because they require different user actions.

## Error Categories

| Category | Example | Presentation | Expected action |
| --- | --- | --- | --- |
| Operational error | Excel cannot be read | Red toast and console error | Correct the file or report the exception. |
| User cancellation | File picker closed | No error | Continue without changing state. |
| Missing data | Missing O-CALC MS or power MS | Warning badge | Enter/import the missing value. |
| Clearance issue | Comm exceeds Max Height at MS | Flagging badge and explicit message | Move comm/Proposed or change approved settings. |
| Data-integrity warning | Unknown linked pole | Warning record | Complete generated data manually or update Excel. |
| Auto Calculate review | No safe candidate | Summary count | Resolve manually; existing data remains unchanged. |

## Import and Update Failures

Excel import and Update Data are wrapped in `try/catch/finally`:

- Exceptions are written with `console.error`.
- The toast includes `error.message`.
- The file input is cleared in `finally`, allowing the same file to be selected again.
- Update Data keeps a pre-import snapshot for reconciliation and undo.

An import fails early when no readable Collection, Span, or Span.Wire data is found.

## Save and Load Failures

- `AbortError` from a picker is treated as cancellation.
- Save writes through `createWritable()`, closes the stream, and marks the job clean only after success.
- Browsers without overwrite support download a JSON fallback and show a warning toast.
- Load always asks for a file; a stored handle is only a starting-location hint.
- If a browser rejects a file handle in `startIn`, Load retries without that option.

## Unsaved Work

Every state-changing UI action records an undo snapshot and marks the job dirty. The Save button displays `Save *` while dirty. `Ctrl+S` uses the application Save flow.

The browser's `beforeunload` warning protects closing or reloading with unsaved changes. Browser security does not permit a fully custom Save/Cancel dialog during tab close.

## Table Flagging

Flagging is recalculated from current state and should always include enough context to identify the problem.

Comm examples:

```text
Power MS: 24'2" > max 23'10".
Pole bolt-bolt: 2" with CTL; minimum 4".
Comm-comm MS on P01 -> P02: CATV 17' vs CTL 16'10"; separation 2"; minimum 4".
```

Proposed rows have two independent columns:

- `MS Flagging`: adjustment/missing/clearance status for Proposed midspan.
- `Proposed Flagging`: pole-side placement, maximum, comm spacing, and bolt spacing.

Rows with a problem receive warning styling in addition to the badge.

## Broad Warning Codes

`validations.js` generates stable warning codes for data integrity and invalid values. Important examples:

| Code | Meaning |
| --- | --- |
| `EDITABLE_OTHER_POLE` | Span created an Unknown endpoint. |
| `EMPTY_SPAN` | Span.Wire supplied no comm rows. |
| `MISSING_MIDSPAN_POWER` | Max Height at MS cannot be derived. |
| `INVALID_ENV_CLEARANCE` | Environment clearance cannot be parsed. |
| `INVALID_PROPOSED` | Proposed height is invalid. |
| `PROPOSED_ABOVE_MAX` | Proposed exceeds its pole limit. |
| `PROPOSED_BOLT_CLEARANCE` | Proposed violates a bolt point. |
| `COMM_MIDSPAN_CLEARANCE` | Existing comm midspan violates a limit. |
| `COMM_ABOVE_MAX` | Existing or changed HOA exceeds Max Height on Pole. |
| `UNKNOWN_OWNER` | Owner is missing or cannot be normalized. |
| `DISCONNECTED_POLE` | Pole has no connected spans. |

The current UI emphasizes compact table flagging. Warning records remain in state for summaries, debugging, and future reporting.

## Diagnostic Workflow

When a value looks wrong:

1. Identify the pole, span label, owner, Existing HOA, HOA Change, and visible midspan.
2. Confirm whether the row is REF.
3. Check the matching row at the other pole and its HOA Change.
4. Check Low Power at MS and Max Height at MS for that exact span.
5. Export Debug.
6. Inspect `midspanCalculations[]` for the selected remote match and both half-movements.
7. Inspect `multipleWiresPerOwner` for duplicate identities.
8. Compare `expectedFromFormula`, `storedCalculated`, and `displayed`.

## Rules for Future Error Messages

New errors should:

1. State what failed.
2. Include pole/span/owner context when available.
3. Show actual and required values.
4. Distinguish missing data from a true violation.
5. Avoid changing state merely to display an error.
6. Keep user-entered values intact when an automatic operation fails.

