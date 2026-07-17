# Pole Span MR Calculator

Pole Span MR Calculator is a static web app for reviewing poles, spans, communications, midspans, clearances and make ready from field Excel exports. It runs directly on GitHub Pages with plain HTML, CSS and JavaScript.

The calculator imports raw pole data, lets users edit existing comm heights, proposes new attachment heights by span, recalculates midspans affected by changes at either end of a span, and generates one Make Ready block per pole.

## Documentation

- [Architecture](docs/ARCHITECTURE.md): runtime modules, data flow, recalculation lifecycle, and extension points.
- [Data Model](docs/DATA_MODEL.md): AppState, entity ownership, identities, physical spans, and REF semantics.
- [Business Rules](docs/BUSINESS_RULES.md): formulas, matching priorities, clearances, Proposed, Auto Calculate, and Make Ready.
- [Import and Export](docs/IMPORT_EXPORT.md): accepted Excel fields, Save/Load, Update Data, AutoProposed, and Debug JSON.
- [Error Handling](docs/ERROR_HANDLING.md): error categories, warning codes, recovery, and diagnostic workflow.
- [Public API Reference](docs/API_REFERENCE.md): supported global modules and cross-module methods.
- [Test Plan](docs/TEST_PLAN.md): regression matrix, browser smoke tests, and automation still needed.
- [Excel Review](docs/EXCEL_REVIEW.md): HOA and Final PLA / MR audit stages, statuses, matching rules, and exclusions.

## Main Features

- Imports poles from the `Collection` sheet.
- Imports pole-to-pole relationships from the `Span` sheet.
- Imports wires, owners, attachment heights and midspans from `Span.Wire`.
- Imports `Attachment Size` references from the `Make Ready` sheet for proposed export data.
- Shows all poles in one editable workspace.
- Provides a collapsible left-side pole index for quick navigation without occupying workspace width.
- Calculates max heights against Low Power.
- Recalculates midspans when comms move up or down on either connected pole.
- Runs Auto Calculate in repeated passes until proposed and HOA changes stop moving.
- Separates real midspan comms from `REF` comms that only reference a connected span.
- Validates power-comm, comm-comm and bolt-bolt clearances.
- Adjusts each Proposed midspan only against comms on the same physical pole-to-pole connection.
- Generates Make Ready by pole from comm movements and proposed attachments.
- Saves and loads complete JSON job files.
- Exports a proposed JSON package intended for a future O-Calc plugin.
- Audits the imported workbook in an `Excel Review` tab after raw import or Update Data.

## Saving Work

`Save` writes a full calculator JSON named from the job, for example `JOB_Pole_Calculator.json`. The leading `EXCEL_` label and a trailing import date such as `_2026-07-13` are removed from the suggested name. In browsers that support the File System Access API, the first save of the current page session asks for a file path and later saves overwrite that same file. Reloading the page or importing a new raw Excel clears the active save destination, so the next `Save` asks for a new path instead of overwriting the previous job.

`Load` always opens the `.json` file picker and restores the complete calculator state. When supported by the browser, it starts from the location of the last loaded or saved JSON. A loaded JSON becomes the active destination for later saves during that session.

`Update Data` imports a newer raw Excel file for the same job and merges it over the current workspace. A populated value from the new workbook refreshes the calculator; an empty replacement keeps the prior known value when the imported entity matches. Omitted comm rows remain when they contain imported HOA/midspan baselines or manual work, while blank endpoint helpers are discarded when the physical connection already exists. Matching HOA changes, proposed values, notes, service-drop/DG checks and manually added proposed spans are preserved. DevTools logs a reconciliation summary and field-level change table for every update.

During an update, comm rows are matched first by their complete imported identity and then by span, pole and owner. This allows a saved HOA movement to follow the same physical comm when a newer Excel changes its `Wire Id`, while stale duplicate rows are discarded before midspans are recalculated.

Pole identity uses the Collection name as the visible ID. Trailing `STEEL`, `UG`, and `PCO` tokens are ignored only while matching, so variants such as `P01-LX339927`, `P01-LX339927 STEEL`, and `P01-LX339927 STEEL UG` resolve to one physical pole instead of creating duplicate cards.

If there are unsaved changes and the page is closed, the browser shows its native leave-page warning. Browsers do not allow custom Save/Cancel buttons during tab close, so the app keeps `Ctrl+S` and the visible `Save` button as the supported save path.

## Workflow

1. Import the raw field Excel file.
2. Review imported poles, spans, power and comms.
3. Edit `Low Power`, existing comm heights, or proposed heights by span.
4. Review recalculated midspans, flagging and clearances.
5. Review the generated Make Ready for each pole.
6. Use `Save` to preserve progress or `Export Proposed` for downstream O-Calc work.
7. Use `Load` later to continue from the saved JSON.

## Excel Review

`Excel Review` is a read-only workbook audit built from the imported source rows and the current calculated state. It runs automatically after `Import Raw Excel` and `Update Data` without changing the active tab. `Re-run Review` recalculates the job and replaces the previous review results.

Every Collection row receives an HOA Review and, when final work exists, a Final PLA / MR Review. HOA Review requires exactly one Fore Span. One Back Span passes, zero produces a warning, and more than one is an error. It also checks reciprocal span relationships, Linked Collection, Environment, and INTEC wire rules. Final Review compares PLA/MRE fields, Proposed heights, final Proposed midspans, generated Make Ready structure, and structured comm transfers. It does not rerun calculator clearances or Pole Type Check.

Review results stay in natural pole/sequence order. The pole name itself is the navigation link back to its Calculator card; the rest of the review row only expands or collapses review details.

Make Ready comparison removes duplicate source instructions, recognizes CenturyLink, CTL and TELCO as the same owner, and accepts PLA/model-only slack, anchor and split-anchor instructions as supplemental work. Those model instructions are not required to have a Calculator equivalent.

Known job exceptions can be marked `Ignore` from an expanded pole review. They remain visible in gray for traceability but stop affecting review badges and summary totals; `Restore` turns the finding back on. Ignored findings are retained in the saved job JSON and through Update Data.

A terminal pole can keep a Proposed attachment even when it has no outgoing span or midspan. Its Span cell remains empty, and that Proposed is available to the preceding pole as `Next Pole Proposed` for End Drop calculation.

When `Add Proposed Span` targets an imported connection, the first Proposed reuses that physical span and its Power, length, bearing and environment data. Additional attachments remain separate Proposed rows but reference the same physical span instead of creating an unrelated blank connection.

## Excel Sheets Used

### Collection

Used to create poles and read general pole data:

- Pole ID or pole name.
- Pole height / type.
- Low Power.
- Tip and Circumference for the Pole Type Check tab.
- Critical height discrepancies are highlighted on the Pole Type Check tab and beside the affected pole in the calculator index.

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

For INTEC, `Size` values containing `Self-Supporting Fiber` are treated as POF. They remain visible in the Existing Comm Movements table with a `POF` badge, but they do not define the Top Comm reference and are not moved by Auto Calculate.

### Equipment

Imported as a read-only Power Equipment section on each pole. The calculator includes power-owned:

- Streetlights.
- Transformers.
- Risers.

Communication-owned risers are excluded. Equipment heights can lower `Max Height on Pole`. INTEC applies the normal Pole Power-comms clearance to the equipment's lowest physical height. MidAm applies its dedicated streetlight bracket and uncovered drip-loop rules to streetlights, while transformers and power risers use Pole Power-comms clearance.

### Anchor.Guys

Used to pre-check the `DG` box on existing comm movements. The importer matches:

- `Id` with the pole ID.
- `Owner` with the comm owner.
- `Attachment Height.display` with the comm `Existing HOA`.

When all three match, Make Ready movement lines can include `with DG`.

For `Metronet > MidAm`, utility guys must be `1/2\"` and communication guys must be `3/8\"`. Guy sizes are audited, but guy attachment heights do not impose a comm clearance rule.

`Service Drop` is shown for INTEC projects and hidden for Metronet projects. Wecom remains an INTEC proposed owner, not a separate project profile. `DG` remains available for every profile because it depends on the imported anchor/guy relationship, not on the service-drop workflow.

`Transfer to New Pole` is one manual option per physical comm at a pole, even when that comm is related to several spans. It changes the normal movement into one transfer instruction at the entered HOA Change. INTEC also provides `Re-sag Service Drop` per span: for a checked service drop below `15'6"`, the calculator validates that span at `15'6"` and adds the corresponding re-sag instruction without changing the imported midspan baseline.

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

MidAm is an explicit exception for Back Span rows that contain their own measured `Span.Wire` midspan. The row remains labeled `REF` so it is not proposed twice, but its displayed midspan is recalculated from movements at both physical endpoints. Reciprocal Fore/Back rows are matched by their pole pair even when their Span IDs and Wire IDs differ.

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

`Metronet > MidAm` also exposes:

- `Pole · Streetlight bracket-comm` (`20\"`).
- `Pole · Streetlight drip loop-comm` (`12\"`, treating drip loops as uncovered).

The pole maximum is the lowest ceiling produced by Low Power and supported Power Equipment. MidAm crossing defaults are `23'6\"` for railroad, `15'6\"` for truck/vehicular areas, `9'6\"` for pedestrian areas, and `14'` over water not suitable for sailboats.

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
- `docs/ARCHITECTURE.md`
- `docs/DATA_MODEL.md`
- `docs/BUSINESS_RULES.md`
- `docs/IMPORT_EXPORT.md`
- `docs/ERROR_HANDLING.md`
- `docs/API_REFERENCE.md`
- `docs/TEST_PLAN.md`

## Local Use

Open `index.html` in a current browser or publish the folder through GitHub Pages. Chrome and Edge provide the complete Save/Load path through the File System Access API; other browsers fall back to JSON downloads and regular file inputs.

The app does not require a backend, custom server, Node.js, npm, build step or framework.

For syntax verification during development, Node.js may be used as an optional tool without becoming a runtime dependency:

```powershell
Get-ChildItem js -Filter *.js | ForEach-Object { node --check $_.FullName }
```
