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

Collection supplies the preferred visible Pole ID. During relationship matching, trailing `STEEL`, `UG`, and `PCO` tokens are ignored. This normalization applies to Collection aliases, Span endpoints, Span.Wire rows, Anchor.Guys, Make Ready references, and Update Data reconciliation.

For Metronet/MidAm, the first block of `Id` is the source of truth for Sequence. It must be `000` or `000A`. The imported `Sequence` cell is normalized to the same shape and audited against that ID-derived value.

Endpoint owner placeholders are marked separately from measured `Span.Wire` rows. During Update Data, a blank inverse-span placeholder is discarded when the previous state already contains the same owner on the same physical connection. Existing HOA and midspan baselines are retained when a later workbook omits them, so a partial update cannot silently change calculations.

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

`Size` values containing Communication Drop/Service Drop mark that span relationship as a Service Drop. In INTEC, Self-Supporting Fiber is eligible for manual POF activation from the comm table; the `POF` state is preserved in JSON and Excel exports.

## Anchor

The `Anchor` worksheet is captured as an independent raw review source. Excel Review checks every existing row for Collection Id, Id, Anchor Index, Anchor Id, Type, Lead Length, Lead Length provider, bearing, pitch, Owner, and Guys. A populated display value may satisfy its corresponding raw Lead Length, bearing, or pitch field. `Anchor` rows do not set the comm `DG` control.

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

When an attachment direction is entered approximately, AutoProposed associates it with the nearest span bearing within 45 degrees (one adjacent compass sector). This keeps a proposed forespan from becoming geometry-only when, for example, a visual `SE` reference is physically closer to an `E` span.

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
- ignored Excel Review findings

Fresh non-empty imported geometry, owners, wires, power, Environment, and source references replace their previous imported equivalents. Empty replacement cells retain the prior known calculator value only after the old and new entities match. Rows omitted by the update are not recreated unless they contain manual user work. This prevents removed spans from generating empty endpoint comm rows. Logical duplicate comm rows are still reconciled to one physical row. Derived values are cleared and recalculated.

`excelReviewSource` is the exception: it always stores the newly selected workbook exactly as imported, including Collection, Span, Span.Wire, Equipment, Anchor, Anchor.Guys, Make Ready, and Make Ready.Comm Transfers. This lets Excel Review report a blank or missing value even when the calculator retains an older value to avoid destructive data loss.

Every update writes a collapsed `[PoleCalc Update Data]` group to the browser console. It includes entity counts, reconciliation diagnostics, and a field-level before/after table for poles, spans, Proposed rows, comms, and power.

## Compact AutoProposed JSON

`Export Proposed` recalculates first and downloads the compact Web-to-plugin contract defined by `schemas/autoproposed.schema.json` and `AUTOPROPOSED_CONTRACT.md`.

All height and span-length values are whole inches. Bearings are decimal degrees.

Example:

```json
{
  "sizes": {
    "messenger": 0.242,
    "fiber": {
      "144": 0.51
    }
  },
  "owner": "Wecom",
  "poles": [
    {
      "id": "P01",
      "spans": [
        {
          "to": "P02",
          "kind": "F",
          "bearing": 90,
          "length": 1500,
          "hoa": 264,
          "fiber": 144,
          "endDrop": -12,
          "nextHoa": 252
        },
        {
          "to": "P03 UG",
          "kind": "O",
          "bearing": 0,
          "length": 960,
          "ug": true
        }
      ],
      "moves": [
        {
          "owner": "CTL",
          "from": 216,
          "to": 228,
          "dg": true
        }
      ]
    }
  ]
}
```

### Sizes

- `sizes.messenger`: positive decimal messenger diameter, or `null` when the payload has no aerial fiber requiring it.
- `sizes.fiber`: map from fiber count to positive decimal diameter.
- Only aerial spans that actually export a `fiber` count create a required size entry.

### Pole fields

- `id`: calculator pole identity.
- `terminalHoa`: terminal standalone Proposed height in inches.
- `spans`: outgoing geometry/proposal entries owned by that pole.
- `moves`: local communication movements.

### Span fields

- `to`: linked pole identity.
- `kind`: `F`, `B`, or `O`.
- `bearing`: degrees in `[0, 360)` when known.
- `length`: whole inches when known.
- `hoa` and `fiber`: explicit aerial proposal.
- `overlash: true`: the proposal is a fiber overlash and must reuse the
  existing messenger for the exported job owner instead of creating another
  messenger.
- `endDrop` and `nextHoa`: optional endpoint proposal values.
- `ug: true`: geometry-only underground relation; it cannot include aerial proposal fields.
- Geometry without `ug`, `hoa`, or `fiber` is a reference-only span and creates no proposed attachment.

### Movement fields

- `owner`: normalized display owner, including `CTL` abbreviation.
- `from` and `to`: whole-inch attachment heights.
- `service: true`: service/drop movement.
- `dg: true`: move matching down guy.

### UG and PCO safety

The safety policy is applied inside `compact-autoproposed.js` before download:

- A normal pole may export local spans, movements and terminal HOA.
- A PCO pole removes local movements, terminal HOA and local aerial proposal fields while retaining geometry-only spans without `ug`.
- A fully UG pole removes local movements, terminal HOA and local aerial proposal fields while retaining geometry-only spans with `ug: true`.
- UG takes priority when UG and PCO are both active.
- A blocked pole with no remaining geometry is omitted.
- A neighboring normal pole may still keep its own aerial proposal toward a PCO pole.

Filename example:

```text
EXCEL_Wecom-SUPRAZ1.1_2026-07-13.xlsx
-> Wecom-SUPRAZ1.1_AutoProposed.json
```

## Debug JSON

Export Debug is intentionally verbose and is available only in the local app
copy (it is hidden and blocked on the public GitHub Pages build). It contains:

- Full current state
- Entity counts
- Multiple-wire groups for the same owner identity
- Latest Update Data diagnostics
- One trace per SpanComm
- Selected remote comm
- Local and remote half-movements
- Imported, expected, stored, and displayed midspans
- Current flagging status and message
- Latest in-memory Auto Calculate trace, including the prior automatic plan,
  baseline groups, candidate order, Proposed/comm plan for every evaluated
  candidate, validation result, ranking decision, stop reason, and selected
  distribution

Use Debug JSON when a visible midspan does not match the expected endpoint movements or when duplicate/stale imported data is suspected.
