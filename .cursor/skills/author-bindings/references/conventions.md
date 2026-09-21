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

`CONSTRAINTS` `Literal[...]` classifies leaves; a constant binding is what
*names* them. Fail closed: constant leaves in `constants.bindings.yaml`,
mutable leaves in `inputs.bindings.yaml`.

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

Bound `data_range` cells must be on-graph. `bindings burndown` and `bindings
audit` build that graph from the current sidecar's `data_range` targets, so
the worklist is residual **inside the bound closure**, not every formula on
every sheet. Widening a shard past the previous end requires widening
extraction targets / overlapping internals so those cells stay in the
dependency graph. Structural blanks are `blank_ranges`, not inputs/constants.
Do not put user-fillable cells in `blank_ranges`.
