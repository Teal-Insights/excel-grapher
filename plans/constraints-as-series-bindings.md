---
status: proposal
tracking: https://github.com/Teal-Insights/excel-grapher/issues/188
spike: excel_grapher/series_bindings/domains.py, tests/unit/series_bindings/test_domains.py
written against: 0e2569f (v20.6.0)
updated: 2026-09-19
---

# Constraints are series bindings

A dynamic-ref constraint is a statement about the values a cell can hold. A
series binding is a statement about what a range of cells *is*. Every
constraint we have ever written sits on a cell that a binding already names,
and the binding already knows more about that cell than the constraint does
(direction, dtype, and, for `input` series, a `domain`). So the proposal is:

> **Stop maintaining a second sidecar.** Promote `domain` to a series-level
> field of the bindings manifest, let `constant` series pin their template
> values, and compile the manifest into the `CellTypeEnv` that dynamic-ref
> inference and `cycle_report` already consume. `constraints.py` becomes a
> derived artifact, then goes away.

## 1. Evidence

### The three constraint kinds already exist in the manifest

`constraints.py` values are `Literal[...]`, `Annotated[int, Between]`, and
`Annotated[float, RealBetween]`. Schema 1.13.0 added `input.domain` with
`enum`, `between`, and `real_between`, defined with the same closed-interval
semantics, and codegen already enforces them on `compute_*` / `Model`
arguments (`require_input_domain`). Two authoring surfaces, one meaning.

### Every committed constraint lands on a bound series

Mapping each `constraints.py` key onto the bindings sidecar of the same
fixture (`expand_bound_series_addresses`, then `normalize_cell_type_env_key`):

| Fixture | Constraint keys | On a bound series | Direction of the series |
| --- | ---: | ---: | --- |
| `tiny_dsa` | 32 | 32 | 8 `constant`, 24 `input` |
| `tiny_dsa_labelled` | 33 | 33 | 3 `constant`, 25 `input`, 5 `internal` |

The constraint *kind* is a function of the direction:

- `constant` series carry singleton `Literal[<template value>]`. That is
  exactly `DynamicRefConfig.from_constraints_and_workbook` with a
  `FromWorkbook()` marker, hand-expanded.
- `input` series carry an `enum` / `between` / `real_between` domain, which is
  the `input.domain` shape.
- The five `internal` cells (`Engine!C5:G5` in the labelled fixture, formulas
  `=Inputs!$B$3`, `=C5+1`, ...) carry `Literal[1]`..`Literal[5]`. An env entry
  on a formula cell short-circuits type analysis (`_enter_cell` consults
  `leaf_env` before reading the formula), so this is a typed-over-formula
  *pin*, the same idea as `input.mode: override`.

### The spike compiles bindings to an identical env

`excel_grapher/series_bindings/domains.py` reads a loaded manifest, pins
`constant` cells to cached workbook values, and turns `input.domain` (or
`value_map` needles) into annotations for `constraints_to_cell_type_env`.
With `input.domain` added to the eight `tiny_dsa` input series in memory, the
result is `==` to the env built from the fixture's `constraints.py`
(`test_tiny_dsa_bindings_compile_to_constraints_py_env`). Nothing about
dynamic-ref inference has to change to consume it.

## 2. Manifest format (schema 1.18.0)

```yaml
series:
  - id: shock_year
    data_range: Inputs!B21
    input: {}
    domain: { between: { min: 1, max: 5 } }

  - id: country_profile_names
    data_range: Inputs!A10:A12
    constant: {}                       # domain: { from_workbook: true } implied

  - id: engine_year_labels
    data_range: Engine!C5:G5
    internal: {}
    domain: { from_workbook: true }    # pin formula cells to template values
```

Rules:

- **`domain` is series-level.** `input.domain` stays accepted and is
  normalized to the series-level key at load time (the same mechanism that
  maps `row_series` to `series`). One declaration feeds both consumers:
  `compute_*` argument checks and dynamic-ref inference.
- **Kinds:** `enum`, `between`, `real_between`, plus `from_workbook: true`,
  which gives each cell in the range its own singleton domain from the cached
  value. Exactly one kind per `domain`.
- **Range is the series.** `data_range` minus `exclude_rows` /
  `exclude_columns` is the constrained range; the sidecar never lists cells.
  A range that needs two domains is two series.
- **Direction defaults.** `constant` implies `from_workbook`. `input` uses
  `domain`, else the values of `value_map` (the workbook needles; the map keys
  remain the public `compute_*` domain). `internal`, `output`, and
  `input.mode: override` contribute nothing unless `domain` is explicit, and
  an explicit domain on a formula cell is documented as a pin.
- **Validation** keeps the dtype rules from 1.13.0, adds "`from_workbook`
  cell has no cached value" as an error, and reports a domain on a cell that
  is not in the graph as a warning.
- **Relations** (`GreaterThanCell`, `NotEqualCell`) are used by no fixture
  and stay out of the manifest until a workbook needs them.

## 3. Lookup: `SeriesDomainIndex`

The index is built from the loaded manifest, so its size is O(#series), not
O(#cells):

```
sheet -> [(rect, excludes, series_id, domain spec), ...]
```

- `domain_for(key)` normalizes the key, picks the sheet bucket, finds the
  covering rectangle, and memoizes the `CellType` per address. Rectangles per
  sheet number in the tens to low hundreds, so a linear scan is fine; an
  interval tree is a drop-in later.
- `from_workbook` values come from the graph's node values when the graph was
  built with `load_values=True`, otherwise from one streaming read of the
  sheet on first miss.
- The index implements `Mapping[str, CellType]`. `CellTypeEnv` is already a
  `Mapping` alias and `dynamic_refs.py` only does `in` and `.get`, so
  inference needs no change. Only `guard.seed_cell_type_env` iterates the env;
  it gets an explicit `items()` over the expanded ranges.
- `DependencyGraph.domains: SeriesDomainIndex | None` replaces
  `cell_type_env`, which today is a full `dict(...)` copy made in
  `create_dependency_graph` and copied again by `subgraph`. `cycle_report`
  reads through it. Pickle and JSON store a handle (`bindings` path relative to
  the workbook plus `bindings_canonical_sha256`); loading re-reads the
  manifest lazily. Nothing lands on `Node`.

## 4. Workflow

1. **Pre-extract.** `list_dynamic_ref_constraint_candidates(workbook,
   targets, dynamic_refs=DynamicRefConfig.from_bindings(bindings, workbook))`
   lists leaves feeding `OFFSET` / `INDEX` / `INDIRECT` with no domain. The
   author adds a `constant` series or an `input` series with `domain` and
   builds. Same loop as today; the answer goes into the sidecar.
2. **Extract.** The builder consumes the index. Only cells that feed dynamic
   refs are ever materialized as `CellType`.
3. **Enrich.** `undomained_leaves(graph, bindings)` returns graph leaves with
   no domain. The author writes or updates shards, then
   `graph.attach_domains(bindings)` swaps the index; later lookups see the new
   entries. No second extract, because a domain on a cell that does not feed a
   dynamic-ref argument cannot change graph structure. A domain that *does*
   feed one is an extraction input, and step 1 exists to catch those first.
4. **CLI.** `bindings validate` and `bindings viz` derive the config from the
   sidecar when `--constraints` is absent.

## 5. Compatibility

The package is on 20.x and #153 declined removing `DynamicRefConfig` without
a `feat!` migration. Three phases:

- **A, additive (minor).** Schema 1.18.0 with series-level `domain`.
  `DynamicRefConfig.from_bindings(bindings, workbook, *, limits)`. CLI derives
  from bindings by default. When `--constraints` is also given, the union is
  used with `constraints.py` winning per key, and a warning lists keys the
  sidecar already covers.
- **B, migration.** Add `domain` to the committed fixtures and local corpus
  bindings, make `constraints` optional in `corpus.toml`, update the user
  guide, delete the fixture `constraints.py` files.
- **C, breaking (`feat!`).** Remove `grapher/constraints.py`, `--constraints`,
  `from_constraints_and_workbook`, and `FromWorkbook`. Keep
  `DynamicRefConfig.from_constraints(dict)` as the low-level Python path the
  compiler targets, dropping the dead `constraints_data` argument.

## 6. Resolved questions from #188

| Question | Answer |
| --- | --- |
| Authoring format | YAML, inside the bindings manifest. No Python DSL. |
| Index for lazy lookup | Per-sheet rectangle list from the manifest, memoized per address. |
| Range vs cell | Ranges only; the series is the range. |
| Breaking change | Phases A and B are additive; C is one `feat!`. |
| Who attaches the sidecar | `create_dependency_graph(..., dynamic_refs=DynamicRefConfig.from_bindings(...))` at build time; `graph.attach_domains(bindings)` for enrichment. |

## 7. Follow-up issues (file once agreed)

1. Schema 1.18.0: series-level `domain`, `from_workbook`, `input.domain`
   normalization, validation rules.
2. `DynamicRefConfig.from_bindings`, CLI derivation, precedence warning.
3. `SeriesDomainIndex`, lazy `Mapping` env, `DependencyGraph.domains`,
   `domain_for`, `attach_domains`, pickle handle; drop `cell_type_env` copies.
4. `undomained_leaves` and a `bindings` subcommand that prints them.
5. Fixture, corpus, and docs migration; delete fixture `constraints.py`.
6. `feat!` removal of the `constraints.py` surface.

The spike ships items 2's core (`dynamic_refs_from_bindings`) against the
current schema and is the acceptance test for item 1.
