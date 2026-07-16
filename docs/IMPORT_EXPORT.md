# Import and Export

## Supported Inputs

The primary input is `.xlsx`. CSV is accepted only as a basic single-sheet fallback. Full saved jobs use `.json`.

Header matching is case-insensitive and punctuation-tolerant. Some fields also support contains matching so minor export header changes do not break imports.

## Collection

| Purpose | Accepted headers or matching rule |
| --- | --- |
| Pole ID | `Id`, `Pole ID`, `PoleId`, `PoleName`, `Structure Number`, `Pole` |
| Collection identity | `collectionId`, `Collection ID` |
| Sequence | `Sequence`, `Seq` |
| Pole type | `Type`, `Pole Type` |
| Pole height | Parsed from Type, then `Pole Height.display`, `Height.display`, `Length.display` |
| Tip | `Tip.display`, `Tip Display`, or a header containing `Tip` |
| Low Power | `Low Power Attachment.display`, related display names, or any header containing `Low Power Attachment`, `Lowest Power`, or `Low Power` |
| Owner | `Owner` |
| Location reference | headers containing `Location.latitude` and `Location.longitude` |
| Excel Review fields | exact `Year Installed`, `MRE Construction Type`, and `PLA STATUS` values are preserved in the raw review source |

Source notes are stored under pole metadata and do not replace user-owned notes.

## Span

| Purpose | Accepted headers or matching rule |
| --- | --- |
| Span ID | `Span Id`, `Span ID`, `spanId`, `Wire Span ID` |
| Current pole | `Id`, `Pole ID`, `Pole`, `From Pole`, or Collection lookup |
| Linked pole | `Linked Collection.Title`, `Linked Collection Title`, `Other Pole`, `To Pole`, `Remote Pole` |
| Linked collection ID | header containing `Linked Collection.ID` or `Linked Collection ID` |
| Type | `Type` |
| Length | `Span Length`, `Span Length.display` |
| Bearing | header containing `Span Length.bearing.display` or `bearing.display` |
| Environment | `Environment` |

Bearing is normalized to one of eight cardinal directions. A missing linked pole creates an editable `Unknown-<spanId>` pole.

## Span.Wire

| Purpose | Accepted headers or matching rule |
| --- | --- |
| Span | `Span Id`, `Span ID`, `spanId`, `Wire Span ID` |
| Pole | `Id`, `Pole ID`, `Pole`, `CollectionId`, `Structure Number` |
| Owner | `Owner`, `owner` |
| Existing HOA | `Attachment Height.display`, `Attachment Height Display`, `Existing HOA`, `HOA`, or matching height fallbacks |
| Midspan | `Mid Span Height.display`, `Midspan.display`, `Midspan` |
| Wire ID | `Wire Id`, `Wire ID`, `wireId` |
| Wire index | `Wire Index` |
| Size | `Size`, `Size.display`, `Wire Size` |
| Construction | `Construction` |
| Insulator | `Insulator` |

Power classification is evaluated before comm creation. Communication owners come from the `Owner` column; missing owners receive stable UNKNOWN labels instead of being discarded.

`Size` values containing Communication Drop/Service Drop mark that span relationship as a Service Drop. INTEC `Self-Supporting Fiber` is marked as POF by calculation logic.

## Anchor.Guys

The importer checks `Id`, `Owner`, and `Attachment Height.display`. A comm receives `downGuy: true` only when pole, normalized owner, and attachment height match.

## Make Ready

The importer reads attachment references, not final Make Ready:

- Pole ID
- `Attachment Size`
- Attachment type
- Attachment height
- Proposed midspan
- Reference notes
- `Comm Transfers`

`Make Ready.Comm Transfers` is also preserved for Excel Review using `Id`, `Owner`, and `Height.display`.

For Excel Review, original headers and row values from Collection, Span, Span.Wire, Make Ready, and Make Ready.Comm Transfers are retained alongside normalized entities. This is required to distinguish a missing column from an empty value and to review duplicate or otherwise non-graphable source rows.

An attachment such as `6.6M 24CT Fiber (E/W)` is preserved raw and parsed into messenger, fiber, and direction tokens for AutoProposed export.

## Project Detection

The raw filename and Span.Wire owners are inspected:

- Text containing `metronet`, `Proposed MNT`, or `MNT` selects Metronet.
- All other files default to INTEC.

The selected profile remains editable in the UI.

## Save JSON

Save writes the complete state:

```json
{
  "app": "pole-calculator",
  "exportedAt": "2026-07-13T18:00:00.000Z",
  "version": "1.4.0",
  "sourceFile": "EXCEL_JOB_2026-07-13.xlsx",
  "state": {}
}
```

The suggested filename removes a leading `EXCEL_`, a trailing `YYYY-MM-DD`, and old calculator/export suffixes:

```text
EXCEL_Wecom-SUPRAZ1.1_2026-07-13.xlsx
-> Wecom-SUPRAZ1.1_Pole_Calculator.json
```

Load restores `state`, normalizes missing fields from older saves, recalculates derived data, and makes the selected file the active Save destination for the current session.

## Update Data

Update Data is intended for a newer Excel from the same job. It does not simply replace AppState.

Preserved user work includes matching:

- HOA changes
- Proposed values and Next Pole Proposed overrides
- Notes
- Service Drop and DG edits
- Transfer to New Pole and INTEC Re-sag Service Drop edits
- Manual/additional Proposed rows
- UG/PCO pole state
- attachment diameter settings

Fresh non-empty imported geometry, owners, wires, power, Environment, and source references replace their previous imported equivalents. Empty replacement cells retain the prior known calculator value only after the old and new entities match. Rows omitted by the update are not recreated unless they contain manual user work. This prevents removed spans from generating empty endpoint comm rows. Logical duplicate comm rows are still reconciled to one physical row. Derived values are cleared and recalculated.

`excelReviewSource` is the exception: it always stores the newly selected workbook exactly as imported. This lets Excel Review report a blank or missing value even when the calculator retains an older value to avoid destructive data loss.

## AutoProposed JSON

Export Proposed recalculates first and creates:

```json
{
  "app": "pole-calculator",
  "exportType": "proposed-for-ocalc",
  "attachmentSizes": {
    "messengerSize": "0.25",
    "fibers": [
      { "fiber": "24CT Fiber", "size": "0.22" }
    ]
  },
  "settings": {},
  "poles": [
    {
      "poleId": "P01",
      "proposedOwner": "Wecom",
      "proposed": [
        {
          "spanLabel": "P01 -> P02",
          "fromPole": "P01",
          "toPole": "P02",
          "proposed": "22'",
          "endDrop": "-1'",
          "nextPoleProposed": "21'"
        }
      ],
      "attachments": [],
      "spans": [
        {
          "label": "P01 -> P02",
          "toPole": "P02",
          "type": "Fore Span",
          "direction": "E",
          "bearingDegrees": 90,
          "lengthDisplay": "125'"
        }
      ],
      "commMakeReady": []
    }
  ]
}
```

Only outgoing spans owned by the exported pole are included. Internal span IDs, Environment, power limits, O-Calc intermediate fields, and redundant `otherPole` fields are intentionally omitted.

Filename example:

```text
EXCEL_Wecom-SUPRAZ1.1_2026-07-13.xlsx
-> Wecom-SUPRAZ1.1_AutoProposed.json
```

## Debug JSON

Export Debug is intentionally verbose. It contains:

- Full current state
- Entity counts
- Multiple-wire groups for the same owner identity
- Latest Update Data diagnostics
- One trace per SpanComm
- Selected remote comm
- Local and remote half-movements
- Imported, expected, stored, and displayed midspans
- Current flagging status and message

Use Debug JSON when a visible midspan does not match the expected endpoint movements or when duplicate/stale imported data is suspected.
