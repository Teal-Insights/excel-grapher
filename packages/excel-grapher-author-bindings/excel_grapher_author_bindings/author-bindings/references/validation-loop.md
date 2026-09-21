# Validation loop

`validate_series_bindings` / `excel-grapher bindings validate` can succeed while
resolution `ok=False`. Codegen requires `ok=True`. Run this loop until audit is
clean, then use burndown only for leftover holes.

```bash
uv run excel-grapher bindings candidates WORKBOOK --bindings BINDINGS
uv run excel-grapher bindings validate WORKBOOK --bindings BINDINGS
uv run excel-grapher bindings undomained WORKBOOK --bindings BINDINGS
uv run excel-grapher bindings audit WORKBOOK --bindings BINDINGS
uv run excel-grapher bindings burndown WORKBOOK --bindings BINDINGS
```

## Extract-time domains

`OFFSET` / `INDEX` / `INDIRECT` arguments are inferred while the graph is
built. Declare those domains on the sidecar and pass
`DynamicRefConfig.from_bindings`. Do not use `--use-cached-dynamic-refs` or a
Python `CONSTRAINTS` table for this step.
`validate_bindings_workbook` defaults `use_cached_dynamic_refs` to True; the
CLI defaults it to False. Call the library with `use_cached_dynamic_refs=False`
when you mean the declared-domain path.

1. Author an extraction root first (usually an `output` series). Empty
   `series: []` shards have no targets. `bindings candidates --target ADDR`
   works before any sidecar exists.
2. `bindings candidates` prints A1 leaves that feed dynamic-ref arguments and
   have no domain. Defined names are resolved: the list contains
   `Lookups!A1`, not the name `BASE`.
3. For each address, add a minimal series: `input` plus `domain`, or
   `constant` when the cached value is non-blank. Re-run candidates until the
   list is empty. Cells behind an unresolved dynamic ref do not appear until
   the controlling domain is wide enough to reach them. A domain that excludes
   the cached selector (for example `enum: [0]` while the workbook stores `1`)
   lets extract succeed and **hides** the downstream `INDIRECT`. Re-run
   candidates after every domain edit on a dynamic-ref argument.
4. `bindings validate` without `--use-cached-dynamic-refs`.
5. `bindings undomained` lists remaining graph leaves (plain inputs, lookup
   values, header pins). Those do not block extract. Domain-only edits on
   them are applied with `graph.attach_domains` and do not need a second
   extract. The CLI rebuilds the graph; that is safe when the new domains do
   not feed a dynamic-ref argument.
6. `greater_than` / `not_equal` name a partner **series id**. Both series
   must already be in the sidecar or `from_bindings` raises
   `SeriesRelationError` and candidates cannot compile. Matching key field
   names are required. The compiled guard is `GreaterThanCell` /
   `NotEqualCell` of the partner address.

`INDEX` / `MATCH` over a static rectangle is often typed from range geometry.
Those selectors do not show up as candidates. An empty candidate list does
not mean the workbook has no `INDEX`.

A blank dynamic-ref leaf is not pinned by `constant` / `from_workbook`
(`None` compiles no `CellType`). Give it an explicit `domain`, or list the
cell in `BLANK_RANGES`. `bindings candidates` does not read `BLANK_RANGES`;
subtract those addresses yourself. Structural pads still appear on the
worklist until you do.

`--strict` exits 1 while any candidate remains.

Optional Python checks after a clean audit:

```python
from excel_grapher.series_bindings import (
    derive_constant_series,
    derive_input_series,
    derive_internal_series,
    derive_output_series,
    load_series_bindings,
    validate_series_bindings,
)

bindings = load_series_bindings("bindings")
report = validate_series_bindings(graph, bindings, workbook=workbook_path)
derive_input_series(graph, bindings, workbook=workbook_path)
derive_output_series(graph, bindings, workbook=workbook_path)
derive_internal_series(graph, bindings, workbook=workbook_path)
derive_constant_series(graph, bindings, workbook=workbook_path)
```

## `bindings audit`

Resolves `input`, `output`, `internal`, and `constant` the same way codegen
does. Exit 1 on errors; `--strict` also fails on warnings.

Tier-1 findings:

- `bind_resolution_failed`
- `partial_bind_failure`
- `empty_public_series` (warning)
- `resolution_not_ok`
- `sparse_label_without_fill`
- `duplicate_internal_cell_binding`
- `duplicate_formula_cell_binding` (output vs internal unique-address ownership)

## `bindings burndown`

Coverage residual **inside the current bound graph closure**: formula nodes on
the graph built from bound `data_range` targets that are not covered by input,
output, or internal bindings (optional `--exempt` file of reviewed addresses).
This is **not** a full workbook walk. Empty shards mean an empty graph and a
zero unbound count.

Prints formula-node count, unbound count, collapsed A1 rectangles, per-sheet
totals, and weak layout **hints** (`scalar` / `series` / `matrix`).

Those ranges are a **coverage worklist**, not candidate bindings. Do not emit
one YAML series per printed row.

`--strict` exits 1 when unbound cells remain. `--max-rows` limits printed row
detail and appends `... (truncated)` when it cuts off.

## `bindings upsert`

Surgical write of **one** series after schema, id, occupancy, shard, and
resolution checks against a temporary copy. The destination tree is written
only when those checks succeed. Use `--replace` to update an existing id.
Never a catalog overwrite of four files.

Library `upsert_series_binding(..., use_cached_dynamic_refs=True)` matches
`validate_bindings_workbook`. The CLI flag `--use-cached-dynamic-refs` defaults
to False, like the other `bindings` commands.

Agents may write a workbook-specific loop that calls upsert many times for
semantic families. That is not a geometry dump.
