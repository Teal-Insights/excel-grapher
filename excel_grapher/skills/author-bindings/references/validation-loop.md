# Validation loop

`validate_series_bindings` / `excel-grapher bindings validate` can succeed while
resolution `ok=False`. Codegen requires `ok=True`. Run this loop until audit is
clean, then use burndown only for leftover holes.

```bash
uv run excel-grapher bindings validate WORKBOOK --bindings BINDINGS
uv run excel-grapher bindings audit WORKBOOK --bindings BINDINGS
uv run excel-grapher bindings burndown WORKBOOK --bindings BINDINGS
```

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

Coverage residual: formula nodes not covered by input, output, or internal
bindings (optional `--exempt` file of reviewed addresses). Prints formula-node
count, unbound count, collapsed A1 rectangles, per-sheet totals, and weak
layout **hints** (`scalar` / `series` / `matrix`).

Those ranges are a **coverage worklist**, not candidate bindings. Do not emit
one YAML series per printed row.

`--strict` exits 1 when unbound cells remain.

## `bindings upsert`

Surgical write of **one** series after schema, id, occupancy, shard, and
resolution checks. Use `--replace` to update an existing id. Never a catalog
overwrite of four files.

Agents may write a workbook-specific loop that calls upsert many times for
semantic families. That is not a geometry dump.
