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

## Grab-bag pins need both address axes

A non-rectangular `data_range` with no year header or row label is unique
only as `(row, column)`. Prefer `bind.kind: row_index` plus
`bind.kind: column_letter` over identity `value_map`s (`2: 2`, `B: B`).
One axis alone still collides when two cells share a row or a column.
A compact range map such as `2: "2:4"` stamps one key onto every row.

## Do not confuse `constant: {}` with `bind.kind: constant`

- `constant: {}` names a reader-only **leaf**.
- `bind.kind: constant` fills a coordinate without reading a cell.

Structural blanks (`INDEX`/`MATCH` padding, NPV window overflow, separator
rows) belong in `BLANK_RANGES`, not constants.

## Extract-time domains are iterative

`bindings candidates` is a pre-extract worklist of A1 leaves, not a series
generator. One scalar `input` with `domain` (or a non-blank `constant`) per
address is enough to extract. Choosing `input` versus `constant` is the
semantic commitment the address table did not force.

Re-run candidates after widening a selector domain. A too-narrow domain can
make extract succeed while omitting a later `INDIRECT` or `OFFSET`.
`attach_domains` covers leaves that do not feed dynamic refs; it does not
discover those hidden cells.

`constant` does not type a blank cell. A `greater_than` / `not_equal`
relation without its partner series fails closed (`SeriesRelationError`)
instead of emitting a guard.

## Extract-only domain pins need `validation.catalog: false`

`OFFSET` / `INDEX` / `INDIRECT` inference reads a `CellType` per cell. A
constant that exists only for that extract-time pin is not a public tensor.
Set `validation.catalog: false` so inverted-tree omits it from the catalog and
`data.py`. Do not stamp identity `ROW`/`COL` keys just to satisfy catalog
uniqueness. A catalog-skipped series that uniquely owns an on-graph formula
cell fails closed.

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
