# Binding conventions

## Directions

Directions are mutually exclusive on one series:

| Block | Meaning |
| --- | --- |
| `input: {}` | Editable graph leaf (required `compute_*` argument). Do not emit removed `input.setter`. |
| `output.compute.name` | Public output. Name must match `compute_[a-z][a-z0-9_]*`. |
| `internal: {}` | Formula-cell triangulation. No public `compute_*`. |
| `constant: {}` | Reader-only **leaf** named for `data.py` / defaulted `compute_*` kwargs. |

Empty `input: {}` marks an editable leaf.

Do not confuse the `constant` **direction** with `bind.kind: constant` (a fixed
structure scalar that does not read a cell).

`domain` / `from_workbook` classifies leaves for dynamic-ref inference; a
constant binding is what *names* them. Fail closed: constant leaves in
`constants.bindings.yaml`, mutable leaves in `inputs.bindings.yaml`.

## Layouts

Use implemented layouts (`IMPLEMENTED_LAYOUTS`):

- `scalar` — one-cell parameters. Keyless (`key: []`) with `PARAMETER` in `series_context` when useful.
- `series` — true 1-D rows or columns. Legacy `row_series` is accepted and normalized to `series`.
- `matrix` — prefer this for a **semantic family** whose rows share a label dimension (scenario, country, indicator, …) and whose headers share a key (usually `TIME_PERIOD`).

Keep scalar / series for true one-cell parameters, true 1-D public concepts, and
rows that are distinct public APIs. Do not stamp one `series` per burndown row.

Measure: `OBS_VALUE` + `bind.kind: data_cell`.

## Keys and dimension identity

Keys are for record matching. Every dimension gets an explicit `id`. `concept`
is the SDMX-style category. They match unless:

- two dimensions share a concept (`PROJECTION_PERIOD` / `REFERENCE_PERIOD` both on `TIME_PERIOD`), or
- an internal formula's operands vary independently (`REF_AREA` vs `COUNTERPART_REF_AREA`).

Without distinct ids those clusters are skipped as `operand_level_variation_unsupported`.

Record fields, key entries, `series_context`, and internals-refactor parameters
use the effective id. Parameter names come from the effective id
(`PROJECTION_PERIOD` → `projection_period`).

## Unique-address ownership

A formula cell has one owner among `output` and `internal` series. An internal
matrix must not cover cells already owned by a public output series (and the
reverse). `input` / `constant` may pair with a formula owner on a **graph leaf**.
Audit reports `duplicate_internal_cell_binding` and `duplicate_formula_cell_binding`.

## Graph coverage

An interior hole and a leading or trailing off-graph section are the same case.
Legality depends on the series direction and what each `data_range` cell is,
not on where the gap sits in the rectangle.

A cell is one of:

- **On-graph value.** The closure reaches it, and it is not a formula.
- **On-graph formula.** The closure reaches it, and it has a formula.
- **Outside the closure.** This extraction never reaches it.
- **Structural blank.** A formula in the closure names it, but it is listed in
  `blank_ranges`, so it is omitted from the graph. A series rectangle may cover
  it; the key is then removed from that series. Do not author an input or
  constant whose only job is that blank. Do not put user-fillable cells in
  `blank_ranges`.

| Cell in `data_range` | Input (leaf) | Input (`mode: override`) | Constant | Internal | Output |
| --- | --- | --- | --- | --- | --- |
| On-graph value | Allowed. This is what an input binds. | Allowed alongside formulas. | Allowed. The range needs at least one, or `no_leaf_constant_targets`. | Allowed. Stays baked. `leaf_in_formula_series` is a warning; double-bind as an input or constant only when callers must supply the value. | Same as internal. |
| On-graph formula | Must fix: `non_leaf_input_overlap`. Narrow to value cells, or use `mode: override`. | Allowed. The range needs at least one, or `no_formula_override_targets`. | Must fix: `non_leaf_constant_overlap`. | Allowed. The range needs at least one, or `no_formula_internal_targets`. | Allowed. Export fails when the range has no on-graph formula (`data_range has no graph formula cells`). |
| Outside the closure | Allowed. `partial_graph_overlap` is a warning. Do not narrow `data_range` to clear it. The cell stays on the published series. | Allowed. Same warning. It is not an override target. | Allowed. Same warning. The published constant still includes it. | Allowed. Same warning. `layout: series` drops the cell from the helper. `layout: matrix` keeps it in the rectangle. | Same as internal. |
| Structural blank | Coverage is allowed. The key leaves the series. It does not count as the required value or formula. | Same. | Same. It does not satisfy `no_leaf_constant_targets`. | Same. It does not satisfy `no_formula_internal_targets`. | Same. It does not satisfy the export formula requirement. |

`partial_graph_overlap` and `leaf_in_formula_series` do not fail resolution.
Do not narrow `data_range` to clear them.

Two exceptions:

- A constant with `validation.intersect_graph_leaves: false` may lie entirely
  outside the closure. That is a lookup table, not a structural blank.
- A series with `axis_labels` must keep every remaining cell inside the
  closure. An outside-closure header cell fails export.

`bindings burndown` and `bindings audit` build the graph from the current
sidecar's `data_range` targets, so the worklist is residual **inside the bound
closure**, not every formula on every sheet. Widening a shard past the previous
end requires widening extraction targets / overlapping internals so those cells
stay in the dependency graph.
