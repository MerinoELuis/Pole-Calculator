# Excel Review

Excel Review audits the currently imported workbook against the calculator state. It is a reporting layer: it does not edit poles, recalculate PLA, or replace Calculator flagging.

## Execution

The review runs after a successful raw Excel import and after Update Data has reconciled user work. It does not change the active tab. `Re-run Review` first recalculates the complete calculator and then replaces all prior review results.

Results remain in memory. The saved job retains the original review source rows, so a loaded job can be reviewed again without storing stale review outcomes.

## Status Model

Each Collection row has:

- `HOA Review`: always runs.
- `Final Review`: runs when Calculator or Excel has final work; otherwise `NOT READY`.
- `Overall`: `ERROR`, then `WARNING`, then `PASS`. `NOT READY` does not lower Overall status.

Pole results are ordered by Overall severity and then naturally by Sequence or Id. A missing Collection `Id` uses its source row as a visible temporary label so the problem is not lost.

## HOA Review

Collection checks require Id, matching Sequence, and a non-empty `Low Power Attachment.display`. Low Power format is not interpreted. Missing Year Installed is a warning with a reminder to review project loading manually.

Span checks operate on original rows:

- Exactly one Fore Span per pole.
- One Back Span passes, zero produces a warning, and more than one is an error.
- Zero or more Other rows; Other never satisfies Fore or Back.
- Reciprocal Fore/Back relationships are warnings when inconsistent.
- Missing or unknown Linked Collection.Title is a low-level warning.
- Environment differs between reciprocal representations of one physical connection is an error.

For INTEC, Span.Wire also checks DAVIT construction, permitted communication owners, APS ownership for Primary/Secondary/Neutral, and the configured insulator lists. Duplicate wire detection is intentionally excluded.

## Final Readiness

Final Calculator work includes HOA changes, Proposed/Next Pole values, O-CALC or final midspan values, End Drop, terminal Proposed, UG, or generated Make Ready instructions. Final Excel work includes populated Make Ready fields or rows in Make Ready.Comm Transfers.

| Calculator | Excel | Result |
| --- | --- | --- |
| Empty | Empty | `NOT READY`; HOA only |
| Has work | Empty | `ERROR` |
| Empty | Has work | `WARNING` |
| Has work | Has work | Full Final comparison |

When Final Review applies, PLA STATUS must be non-empty. MRE Construction Type must be Aerial or Underground and must agree with the Calculator solution.

## Final Comparison

Height comparisons use `HeightUtils.parseHeight`, so feet/inches and decimal feet are equivalent after conversion to integer inches.

For aerial work, Proposed HOA is compared with Attachment Height and the latest applicable Proposed midspan is compared with Proposed Mid Span using this priority:

1. `finalMidspan`
2. `msProposed`
3. `proposedMidspan`
4. `ocalcMS`

Rows are ranked by span identity or physical connection when those fields exist, then attachment direction, and only then closest height.

Make Ready Notes are compared as instructions rather than as exact strings. Action, owner aliases, heights, direction, DG, slack, anchor, and riser markers are normalized. Punctuation, case, whitespace, and line breaks are ignored.

Duplicate source instructions are collapsed before comparison, and an exact normalized instruction is consumed before semantic fallback matching. PLA/model-owned instructions such as Proposed slack spans, new anchors/down guys, transfers to a new comm anchor, and split power anchors are accepted as supplemental model work instead of producing an unmatched Calculator warning.

When Calculator and Excel contain the same kind of instruction but differ in direction, height, owner, or a required modifier, the review reports one paired mismatch. It does not also report the Excel line as an additional warning.

UG, riser, anchor and slack directions remain significant during Make Ready comparison. A populated Linked Collection may point to a pole in another job and is not required to match this workbook's Collection. Empty Fore/Back links remain reviewable; empty Other links are ignored.

For UG work, the audit requires Underground construction and a UG instruction. Aerial attachment height and midspan are not required.

Structured comm transfers compare Make Ready.Comm Transfers Owner and Height against Calculator transfer groups. Century Link, CTL and TELCO are treated as the same owner.

Review cards remain in natural Collection sequence/pole order regardless of severity. Only the pole-name text navigates to the matching Calculator card; clicking elsewhere on the summary row expands or collapses its checks. An `Other` span with no linked pole is allowed and does not create a Linked Collection warning.

Each pole-level non-PASS finding has an `Ignore` control. Ignored findings remain visible in gray inside the expanded pole, but they do not affect HOA, Final, Overall, or summary counts. `Restore` activates the finding again. These job-specific decisions are saved in Calculator JSON and preserved by Update Data.

The `Excel Review` view tab shows the number of poles with active Error or Warning status. Ignored findings do not contribute to this badge.

## Not Applicable

Attachment Size exact matching is marked Not applicable when Calculator does not contain a reliable per-Proposed fiber/messenger identity. Comm Transfers is Not applicable when no transfer is expected or the structured sheet is unavailable.

Excel Review intentionally excludes clearances, Pole Type Check, loading cases, PCO rules, AutoQC, O-Calc access, duplicate wire checks, and Low Power formatting.
