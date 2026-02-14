# Quality Checklist

Run this checklist before final output.

## Anchor Quality

- `context` is unique enough in the source text.
- `prefix + placeholder + suffix` reconstructs original local phrase.
- `placeholder` excludes fixed units/titles/symbols that should remain static.

## Semantic Quality

- Company/person/date/amount fields are templatized when appropriate.
- Legal references and mandatory boilerplate are not incorrectly variablized.
- Optional clauses are modeled with `mode: "paragraph"` only when whole clause toggling is needed.

## UI Mapping Quality

- `type` is practical for user input.
- `formatFn` is conservative (`none` by default).
- `select`/`radio` options are complete and mutually exclusive.

## Consistency Quality

- Tags are unique and deterministic.
- Similar entities across clauses use consistent naming strategy.
- No missing required fields.
