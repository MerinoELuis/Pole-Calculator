# Pole Span MR Calculator

Pole Span MR Calculator is a static web app for reviewing poles, spans, communications, midspans, clearances and make ready from field Excel exports. It runs directly on GitHub Pages with plain HTML, CSS and JavaScript.

The calculator imports raw pole data, lets users edit existing comm heights, proposes new attachment heights by span, recalculates midspans affected by changes at either end of a span, and generates one Make Ready block per pole.

## Main Features

- Imports poles from the `Collection` sheet.
- Imports pole-to-pole relationships from the `Span` sheet.
- Imports wires, owners, attachment heights and midspans from `Span.Wire`.
- Imports `Attachment Size` references from the `Make Ready` sheet for proposed export data.
- Shows all poles in one editable workspace.
- Draws a relative pole map from span length and bearing, with every connected and disconnected pole visible.
- Provides a collapsible left-side pole index for quick navigation without occupying workspace width.
- Calculates max heights against Low Power.
- Recalculates midspans when comms move up or down on either connected pole.
- Runs Auto Calculate in repeated passes until proposed and HOA changes stop moving.
- Separates real midspan comms from `REF` comms that only reference a connected span.
- Validates power-comm, comm-comm and bolt-bolt clearances.
- Generates Make Ready by pole from comm movements and proposed attachments.
- Exports/imports full JSON save points.
- Exports a proposed JSON package intended for a future O-Calc plugin.

## Saving Work

`Save` writes a full calculator JSON named from the job, for example `EXCEL_JOB_Pole_Calculator.json`. In browsers that support the File System Access API, the first save asks for a file path and later saves overwrite that same file. The selected file handle is remembered when the browser allows it, so `Save` and `Load` can reuse the same file later. `Load` opens a `.json` save file and restores the calculator state.

`Update Data` imports a newer raw Excel file for the same job and merges it over the current workspace. Imported field data is refreshed, while matching user work such as HOA changes, proposed values, notes, service-drop/DG checks and manually added proposed spans is preserved.

During an update, comm rows are matched first by their complete imported identity and then by span, pole and owner. This allows a saved HOA movement to follow the same physical comm when a newer Excel changes its `Wire Id`, while stale duplicate rows are discarded before midspans are recalculated.

If there are unsaved changes and the page is closed, the browser shows its native leave-page warning. Browsers do not allow custom Save/Cancel buttons during tab close, so the app keeps `Ctrl+S` and the visible `Save` button as the supported save path.

## Workflow

1. Import the raw field Excel file.
2. Review imported poles, spans, power and comms.
3. Edit `Low Power`, existing comm heights, or proposed heights by span.
4. Review recalculated midspans, flagging and clearances.
5. Review the generated Make Ready for each pole.
6. Export JSON to save progress or export Proposed JSON for downstream O-Calc work.
7. Import the saved JSON later to continue.

## Excel Sheets Used

### Collection

Used to create poles and read general pole data:

- Pole ID or pole name.
- Pole height / type.
- Low Power.
- Tip and Circumference for the Pole Type Check tab.

Low Power is matched flexibly. The app accepts headers containing `Low Power Attachment`, plus fallbacks such as `Lowest Power` or `Low Power`.

### Span

Used to create graph connections between poles:

- Span ID.
- Source pole.
- Connected pole.
- Span length.
- Bearing angle.
- Span type.
- Environment.

Bearing is converted to a cardinal direction so span relationships are easier to read.

## Relative Pole Map

The first known pole in each connected group is used as the local `0,0` origin. Every connected pole is projected from the imported span length and bearing, where `0` is north, `90` east, `180` south and `270` west. Reciprocal Fore Span and Back Span records are drawn as one physical connection so the same line is not duplicated.

Disconnected groups and individual poles are arranged beside the main network so every imported pole remains visible. A dashed connection indicates that length or bearing was missing and a deterministic visual fallback was required. The map supports zoom, fit, drag-to-pan, pole selection and span selection. The collapsible left index remains available for filtered navigation.

### Span.Wire

Used to create comm and power rows by span:

- `Owner`.
- `Attachment Height.display`.
- `Mid Span Height.display`.
- `Wire Id`.
- `Size`.
- `Construction`.
- `Insulator`.

The visible Owner/Comm value comes from the `Owner` column.

For INTEC, `Size` values containing `Self-Supporting Fiber` are treated as POF. They remain visible in the Existing Comm Movements table with a `POF` badge, but they do not define the Top Comm reference and are not moved by Auto Calculate.

### Anchor.Guys

Used to pre-check the `DG` box on existing comm movements. The importer matches:

- `Id` with the pole ID.
- `Owner` with the comm owner.
- `Attachment Height.display` with the comm `Existing HOA`.

When all three match, Make Ready movement lines can include `with DG`.

`Service Drop` is shown for INTEC projects and hidden for Metronet projects. Wecom remains an INTEC proposed owner, not a separate project profile. `DG` remains available for every profile because it depends on the imported anchor/guy relationship, not on the service-drop workflow.

### Make Ready

Used as imported reference data, not as the final generated Make Ready. The app reads:

- `Attachment Size`, such as `6.6M 24CT Fiber (E/W)`.
- Attachment type.
- Attachment height.
- Proposed midspan.
- Reference notes.

The Proposed JSON export preserves the raw attachment size and also breaks it into messenger, fiber and direction fields.

For INTEC jobs, detected fiber counts such as `12CT`, `24CT`, `72CT` and `144CT` create a `Fiber` configuration section. The user can enter one messenger diameter and a separate fiber diameter for every detected fiber count. These values are saved with the job and exported in the top-level `attachmentSizes` object of the AutoProposed JSON.

## Calculation Rules

### Existing Comm Movements

Each existing comm can have a `HOA Change`. When a height changes at one end of a span, the span midspan is recalculated with half of that movement.

Example:

- If a comm moves from `20'` to `19'`, the movement is `-12"` and the midspan drops `6"`.
- If the comm on the other pole moves from `20'` to `21'`, the movement is `+12"` and the midspan rises `6"`.
- If both ends move, both effects are applied to the midspan.

Comms without their own imported midspan can show as `REF`. That means they belong to the connected span, while the real midspan comes from the other end.

### Proposed by Span

`Proposed by Span` is only used for spans that have real midspan data or manually added proposed spans. It avoids repeating backspans and duplicate physical connections.

End Drop is calculated from the local Proposed value and the Next Pole Proposed value when available.

### Clearances

Editable clearance values control these checks:

- `Pole · Power-comms`.
- `Pole · Comm-comm`.
- `Pole · Bolt-bolt`.
- `Midspan · Power-comm`.
- `Midspan · Comm-comm`.

Between comms on the pole:

- Different owners use `Pole · Comm-comm`.
- The same owner uses `Pole · Bolt-bolt`.

Proposed attachments also check against existing attachment points. A proposed height may reuse an exact existing HOA only when that existing comm was moved away. It cannot land inside the bolt-bolt restricted zone around existing attachment points.

### Low Power at Midspan

The app calculates `Max Height at MS` from the lowest power midspan on the span minus `Midspan · Power-comm`.

For INTEC, when a proposed adjustment is specifically required by Low Power at midspan, the Make Ready can include:

`Ensure min 30" to low power at midspan.`

For Metronet, the proposed midspan is not automatically adjusted down for this rule and that Make Ready line is not generated. The clearance violation remains visible so it can be solved by lowering the affected comms.

## Rule Table

| Rule | Default | Applies To | Result |
| --- | --- | --- | --- |
| `Pole · Power-comms` | `40"` | comms and proposed on the pole | Defines `Max Height on Pole`. |
| `Pole · Comm-comm` | `12"` | different-owner comms on the pole | Prevents comms from being too close on the pole. |
| `Pole · Bolt-bolt` | `4"` | same-owner comms, proposed-to-proposed, and existing bolt points | Prevents bolts from being too close. |
| `Midspan · Power-comm` | `30"` | comms and proposed at midspan | Defines `Max Height at MS`. |
| `Midspan · Comm-comm` | `4"` | comms and proposed at midspan | Keeps vertical separation between cables. |
| `Environment` | by span | comms and proposed at midspan | Prevents values below the environment minimum. |
| Comm order | by heights | existing comms on the same span | Prevents cables from crossing between pole and midspan. |
| `Top Comm` / `Low Comm` | configurable | proposed on pole and midspan | Forces proposed to stay on the configured side of existing comms. |

For `Proposed` on the pole:

- It must always keep `Pole · Comm-comm` from existing comms.
- It may share the exact same height as an existing attachment only when that comm has been moved away.
- Separate proposed attachments on the same pole use `Pole · Bolt-bolt`.

## Proposed JSON Export

The `Export Proposed` button creates a `.json` package for downstream O-Calc automation. It includes:

- A compact `poles[]` list grouped by Pole ID.
- A top-level `attachmentSizes` catalog with the configured messenger size and one size per detected fiber count.
- For each pole, sibling `proposed[]`, `attachments[]` and `spans[]` collections.
- Proposed, End Drop and Next Pole Proposed values grouped under the pole that owns the proposal.
- Outgoing spans with usable length, bearing or direction data.
- Comm-only Make Ready lines grouped under the same pole.
- Make Ready `Attachment Size` data needed by O-Calc, including messenger, fiber and direction.

The proposed export intentionally omits internal app IDs such as `spanId` and avoids repeating environment or Low Power clearance fields that O-Calc does not need for placing the proposal.

`Export Debug` creates a separate diagnostic JSON. It contains the full calculator state, multiple-wire groups for the same owner, the result of the latest Excel reconciliation and one midspan trace per comm showing the imported value, local half-movement, selected remote comm, remote half-movement, expected result and displayed result.

## Main Files

- `index.html`
- `css/styles.css`
- `js/app.js`
- `js/state.js`
- `js/height-utils.js`
- `js/excel-import.js`
- `js/excel-export.js`
- `js/json-export.js`
- `js/calculations.js`
- `js/mr-logic.js`
- `js/validations.js`
- `js/floating-calculator.js`
- `libs/xlsx.full.min.js`

## Local Use

Open `index.html` in a browser or publish the folder through GitHub Pages.

The app does not require a backend, custom server, Node.js, npm, build step or framework.
