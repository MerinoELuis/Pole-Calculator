# Implementation decisions

These decisions record approved behavior so future changes do not need to rediscover it from conversation history.

## DEC-001 — UG has priority over PCO

When a pole is both UG and PCO, UG controls local export and Make Ready behavior. Every local exported span is geometry-only with `ug: true`.

## DEC-002 — Block local actions on UG and PCO poles

A completely UG or PCO pole does not export local `moves` or `terminalHoa`. Local proposal fields (`hoa`, `fiber`, `endDrop`, `nextHoa`) are removed. A PCO span remains geometry-only without `ug`; a UG span includes `ug: true`.

A neighboring normal pole may still export its own aerial proposal toward a PCO destination. Only the PCO pole's local entry is blocked.

## DEC-003 — Keep Other Pole HOA internal

`Other Pole HOA` is calculation data used by the midspan logic. It is hidden from the Existing Comm Movements table but is not deleted from AppState or calculation results.

## DEC-004 — Restart span colors per pole

Every pole starts its visible span sequence at:

1. `span-color-0` — turquesa
2. `span-color-1` — naranja
3. `span-color-2` — azul oscuro
4. `span-color-3` — rosa
5. `span-color-4` — verde

Hidden spans do not consume a visible color position. Midspan rows, including blank or `REF` rows, must use the same color as their paired span.

## DEC-005 — Stable DOM identities

Rendered tables are annotated with stable attributes such as `data-pole-id`, `data-span-id`, `data-column` and `data-component`. UI behavior should prefer these identifiers over visible English labels or row-position inference. Position matching remains only as a fallback for legacy markup.

## DEC-006 — Feet and inches input

The floating calculator displays `Feet' Inches"`. A single apostrophe represents feet. Two consecutive apostrophes entered manually are normalized to a double quote for inches without moving the cursor unexpectedly.

## DEC-007 — One UI refresh coordinator

Post-render UI behavior is coordinated by `js/ui/runtime.js`. New table or annotation modules should expose a synchronous `refresh(root)` method and should not create their own MutationObserver.

## DEC-008 — One test command

`npm test` is the required local and CI command. Runtime remains plain HTML/CSS/JavaScript; Node is development tooling only and does not add a GitHub Pages build step.
