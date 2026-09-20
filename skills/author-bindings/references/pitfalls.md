# Authoring pitfalls

## Sparse labels need `fill: true`

Group labels often sit only on the first column or row of a block:

```text
header row:  2050   (blank) (blank)  2075
measure row: Base   Alt     Gap      Base
```

Without `fill: true` on `column_header` / `row_label`, blank label cells raise
`bind_resolution_failed` and resolution `ok=False` even when
`validate_series_bindings` looked fine. That is the usual cause of
`partial_bind_failure`.

## Mixed measures vs a short public key

Do not bind one rectangle across mixed measure columns when the public key is
only `(SCENARIO, TIME_PERIOD)` (or similar). Columns collide or only a subset
resolves. Shard per measure, **or** add `MEASURE` to the key (typical for
internals that triangulate the whole block). See
[assets/measure-shards.example.yaml](../assets/measure-shards.example.yaml).

## Shared vs unique compute names / input ids

Share `output.compute.name` or an input series `id` **only** to merge
complementary slices of one public series (for example Gap milestone columns
that should become one `compute_gap_milestones`).

Uniquify distinct scenario / engine paths. Sharing a name across those paths
merges definitions at export and can leave most paths **unreachable** in the
exported library even though every shard still looks valid in YAML.

## Do not confuse `constant: {}` with `bind.kind: constant`

- `constant: {}` names a reader-only **leaf**.
- `bind.kind: constant` fills a coordinate without reading a cell.

Structural blanks (`INDEX`/`MATCH` padding, NPV window overflow, separator
rows) belong in `BLANK_RANGES`, not constants.

## Anti-pattern: geometry-first YAML

Do **not**:

- dump unbound burndown rows as one series per printed line
- expand a catalog into four files and “clean up later”
- stamp `#724` candidates as declarations

Those paths almost exclusively create merge/re-key junk. Author known tables
(semantic families, often matrices) correctly the first time. Burndown is a
**coverage worklist**, not a generator.

A custom workbook-specific script that upserts many series is fine when those
series are chosen semantically. A generic bulk emitter is not.
