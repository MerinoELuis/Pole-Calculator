# Compact AutoProposed contract

The formal structural contract is `schemas/autoproposed.schema.json`. This document records business rules that JSON Schema cannot fully express.

## Units

- Heights: whole inches.
- Span lengths: whole inches.
- Bearings: decimal degrees from `0` inclusive to `360` exclusive.
- Messenger and fiber sizes: positive decimal values.

## Processing safety

1. Recalculate the complete calculator state before export.
2. Build compact spans and movement candidates.
3. Apply the pole-level UG/PCO safety policy.
4. Validate required messenger and fiber sizes only for aerial spans that actually use fiber.
5. Download the compact payload.

## Pole rules

- Normal pole: may contain `terminalHoa`, `spans` and `moves`.
- PCO pole: remove local `moves` and `terminalHoa`; retain geometry-only local spans without `ug`.
- Completely UG pole: remove local `moves` and `terminalHoa`; retain geometry-only local spans with `ug: true`.
- UG takes priority when UG and PCO are both active.
- A blocked pole that has no remaining geometry is omitted from `poles`.
- A neighboring normal pole may keep its own aerial proposal toward a PCO pole.

## Span rules

- `ug: true`: do not include `hoa`, `fiber`, `endDrop` or `nextHoa`.
- `hoa` plus `fiber`: explicit aerial proposal.
- Geometry without `ug`, `hoa` or `fiber`: reference only; create no proposed attachment.
- `kind` is `F`, `B` or `O`.

## Movement rules

- Every movement requires `owner`, `from` and `to`.
- `service: true` identifies a service/drop movement.
- `dg: true` requests matching down-guy movement.
- Transfer-only internal state is normalized into the movement destination and is not exported as a separate `transfer` property.

## Fixtures

Reusable examples are under `tests/fixtures/`, and exact expected Web output is under `tests/expected/`.
