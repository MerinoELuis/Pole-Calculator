# Pole Span MR Calculator

Pole Span MR Calculator is a static web app for reviewing poles, spans, communications, midspans, clearances and make ready from field Excel exports. It runs directly on GitHub Pages with plain HTML, CSS and JavaScript.

The calculator imports raw pole data, lets users edit existing comm heights, proposes new attachment heights by span, recalculates midspans affected by changes at either end of a span, and generates one Make Ready block per pole.

## Main Features

- Imports poles from the `Collection` sheet.
- Imports pole-to-pole relationships from the `Span` sheet.
- Imports wires, owners, attachment heights and midspans from `Span.Wire`.
- Imports `Attachment Size` references from the `Make Ready` sheet for proposed export data.
- Shows all poles in one editable workspace.
- Calculates max heights against Low Power.
- Recalculates midspans when comms move up or down on either connected pole.
- Runs Auto Calculate in repeated passes until proposed and HOA changes stop moving.
- Separates real midspan comms from `REF` comms that only reference a connected span.
- Validates power-comm, comm-comm and bolt-bolt clearances.
- Generates Make Ready by pole from comm movements and proposed attachments.
- Exports/imports full JSON save points.
- Exports a proposed JSON package intended for a future O-Calc plugin.

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
- For each pole, the related `spans[]` with span label, other pole, length, bearing and direction.
- A clear `proposed` object inside each span with Proposed, End Drop, Next Pole Proposed, O-CALC MS, MS Proposed and Adjusted Final MS.
- Comm-only Make Ready lines grouped under the same pole.
- Make Ready `Attachment Size` data needed by O-Calc, including messenger, fiber and direction.

The proposed export intentionally omits internal app IDs such as `spanId` and avoids repeating environment or Low Power clearance fields that O-Calc does not need for placing the proposal.

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
