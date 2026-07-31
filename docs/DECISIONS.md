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
6. `span-color-5` — rojo
7. `span-color-6` — verde oscuro
8. `span-color-7` — azul claro
9. `span-color-8` — salmón
10. `span-color-9` — amarillo

Hidden spans do not consume a visible color position. The sequence follows the
natural imported Span Index order and repeats only after ten visible physical
spans. Midspan rows, including blank or `REF` rows, use their explicit
`spanId` and must always have the same color as their paired span.

## DEC-005 — Stable DOM identities

Rendered tables are annotated with stable attributes such as `data-pole-id`, `data-span-id`, `data-column` and `data-component`. UI behavior should prefer these identifiers over visible English labels or row-position inference. Position matching remains only as a fallback for legacy markup.

## DEC-006 — Feet and inches input

The floating calculator displays `Feet' Inches"`. A single apostrophe represents feet. Two consecutive apostrophes entered manually are normalized to a double quote for inches without moving the cursor unexpectedly.

## DEC-007 — One UI refresh coordinator

Post-render UI behavior is coordinated by `js/ui/runtime.js`. New table or annotation modules should expose a synchronous `refresh(root)` method and should not create their own MutationObserver.

## DEC-008 — One test command

`npm test` is the required local and CI command. Runtime remains plain HTML/CSS/JavaScript; Node is development tooling only and does not add a GitHub Pages build step.

## DEC-009 — Pole compliance has priority over Midspan

Auto Calculate compares candidate arrangements by result category before movement cost. `SAFE` is best. When no SAFE arrangement exists, a pole-compliant `BEST_AVAILABLE` arrangement with Midspan violations is always preferred over an arrangement that fixes Midspan but leaves any pole violation.

## DEC-010 — Retain the best partial aerial arrangement

A candidate is no longer discarded only because one or more violations remain. The solver retains the best evaluated arrangement as `BEST_AVAILABLE` or `CRITICAL`, reports the remaining issue count/shortfall, and lets the operator decide the final construction resolution.

## DEC-011 — Auto Calculate supports TOP COMM and LOW COMM

TOP COMM targets `Top Comm + Comm-comm`; LOW COMM targets `Low Comm - Comm-comm`. The rule defines the ideal Proposed position, while the solver minimizes remaining violations, moved comm groups, total movement and distance from that ideal.

## DEC-012 — UG and PCO remain operator decisions

Auto Calculate may recommend manual UG/PCO review for INTEC, but it never activates either option. Poles already marked UG or PCO are skipped by the aerial solver.

## DEC-013 — TOP COMM uses progressive search and selective retries

TOP COMM tries the direct Proposed position first, the pole ceiling when direct space is unavailable, and exact Midspan-driven movement boundaries before opening the wider fallback search. The first SAFE result is final. A pole-compliant BEST_AVAILABLE result after all direct progressive candidates is also final, so the wider scan is reserved for cases that still violate the pole. Connected poles are not processed in unconditional global passes; a pole is retried only when a later connected change creates an opportunity to remove automatic movements or improve a remaining violation.

## DEC-014 — TOP COMM Midspan recovery starts at the pole ceiling

When the current TOP COMM arrangement has a Midspan violation, the first recovery candidate places Proposed at Max Height on Pole and arranges comms downward from their highest legal positions. The stack may reuse an existing bolt height but avoids the clearance exclusion around every other existing bolt point. Pole and upper-clearance compliance remain mandatory; if the remaining Midspan profile cannot improve, the highest equally safe Proposed is retained.

## DEC-015 — Automatic retries rebuild from imported geometry

A previous automatic arrangement is replaceable output, not the geometric baseline for the next retry. The solver clears local `AUTO` Proposed and HOA Change values before generating candidates, preserves manual values, and creates a propagated Midspan floor only when a real minimum or lock exists. This prevents automatic results from recursively lowering Top Comm toward zero.
