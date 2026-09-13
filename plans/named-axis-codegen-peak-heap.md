# Reducing peak heap during named-axis code generation

Date: 2026-09-11. Branch: `migration/named-axis-codegen`.
Status: plan; nothing below is implemented.

## Objective and scope

Lower the peak heap of the **generator process** while
`generate_inverted_tree_modules` (`excel_grapher/exporter/inverted_tree/emit.py`)
plans and emits a named-axis package, without changing the generated modules.
This is the memory of `plan_inverted_tree` plus `emit_named_modules`, measured
with the caller's `DependencyGraph` already resident. It is not the memory of
the generated package at import or evaluation time; that is covered by
[the public-code report](named-axis-public-code-report.md) and is unchanged here.

"Safely" means every change is observably neutral:

- Generated `api.py`, `internals.py`, `data.py`, `runtime.py`, `excel.py`,
  `tensor.py`, `provenance.py`, and `__init__.py` are byte-identical before
  and after, on every committed fixture and on the LIC DSF sandbox package.
- `CODEGEN_FINGERPRINT` (`named_codegen_fingerprint`) is byte-identical.
- The same `InvertedTreeExportError` messages are raised for the same inputs.
- The caller's graph is not mutated and no new module-level cache grows with
  workbook size.

Peak heap is a secondary objective on this branch. The 10 MB source gate,
parity, and readability in [the optimization plan](named-axis-public-code-optimization.md)
remain the primary gates; no change here may trade them away.

## Baseline and diagnosis

### What is known

No measurement of generator memory exists. The migration report records
746.1 s of generation for the LIC DSF workload and the evidence JSON records
`timings: {}`; `scripts/measure_named_export.py` times generation but does not
sample memory. The LIC DSF inputs (`sandbox/lic-dsf/`) are not in this
container, so the numbers below come from a synthetic workload built with the
test helpers (`tests/unit/exporter/inverted_tree/helpers.py`): one sheet,
60 formula series and 6 input series over 40 years (2,400 formula cells,
240 input cells; recurrences with seeds, growing-window `SUM`s, lagged `IF`s,
fixed-window `AVERAGE`s; 7,155 dependence edges). Tracemalloc with
`gc.collect()` at every phase boundary; `capture_dependency_provenance=True`
as in the test harness.

| Phase | Retained after (MB) | Phase peak (MB) | Notes |
| --- | ---: | ---: | --- |
| `create_dependency_graph` | 12.9 | 17.9 | Graph, ASTs, edge sets, guards, provenance |
| `_refuse_invalid_bindings` | 12.9 | 14.4 | Transient only |
| `build_catalog` | 14.4 | 15.4 | `KeyPoint`s, `_cell_indices`, schedule maps |
| `collect_catalog_edges` | 15.6 | 15.7 | 7,155 `DependenceEdge` (104 B each) + cell strings |
| `collect_all_deps` | 15.8 | 15.9 | `SeriesDeps` keeps a second tuple of the same edges |
| `build_scc_map`, `plan_scc`, `assert_subgraph_bound` | 15.8 | 16.1 | Transient only |
| `emit_named_internals` | 16.6 | 16.8 | `_emit_cache` ref lists (367 KiB), address strings |
| `emit_named_api` | 16.6 | 16.7 | |
| `emit_named_data` | 16.7 | 17.6 | Defaults, fingerprint payload |
| `build_runtime_modules` | 16.9 | 26.1 | Fixed ~9 MB `ast` spike embedding the runtime |

Releasing objects at the end attributes the retained total: graph 9.4 MB,
catalog 2.2 MB, edges and deps 1.4 MB, everything else about 0.2 MB. A second
run at 150 formula series and 15 inputs over 50 years (7,500 formula cells,
8,250 bound cells, 22,337 edges; one-frame tracemalloc) scales linearly:

| Workload | Graph (MB) | Catalog (MB) | Edges and deps (MB) | Emission adds (MB) | `ru_maxrss` at end (MB) |
| --- | ---: | ---: | ---: | ---: | ---: |
| 60x40 (2,640 bound cells, 7,155 edges) | 9.4 | 2.2 | 1.4 | 0.9 | 187 |
| 150x50 (8,250 bound cells, 22,337 edges) | 32.7 | 8.0 | 4.3 | 3.3 | 372 |

Per unit that is roughly 3.6 to 4 KB per graph node, 850 to 970 B per bound
cell in the catalog, and 190 to 200 B per dependence edge. Scaled to the LIC
DSF workload (about 130,000 bound cells across 1,938 series, per the
migration report), the codegen-owned structures are in the low hundreds of
megabytes and the graph is larger again; the exact split must be measured,
not extrapolated. Process RSS runs several times the traced Python heap in
both runs (tracemalloc overhead, allocator slack, and the workbook reader),
so RSS and traced peak must be reported separately.

Two further observations from the same run:

- **Address strings are not shared.** 41,012 strings (2.0 MB) allocated in
  `excel_grapher/core/address_keys.py` survive for about 2,640 distinct
  addresses; 137,923 strings (6.7 MB) for about 8,300 at the larger size. `format_cell_key`, `resolve_cell_ref`, and `as_canonical`
  (`address_keys.py:221`, a `NewType` over `str`) mint a new `str` each time,
  and the graph, catalog, edge, and schedule maps each hold their own copies.
- **Domains are rebuilt per reference.** Under cProfile, 60 formula series
  triggered 4,403 `BoundSeries.tensor_domain` evaluations, 2,847
  `coordinate_cells`, and 2,707 `required_coordinates`, each building fresh
  `Axis`/`Domain` objects and dictionaries (`catalog.py:196`, `:239`, `:252`);
  8,802 `Domain.__post_init__` calls cost 0.8 s of 7.5 s. These are transient,
  but they set the high-water mark inside a series and dominate allocator churn.
  `emit_named_data` evaluates `series.required_coordinates` once per coordinate
  inside a generator expression (`named_emit.py:1169`), which is quadratic in
  the domain size.

### Lifetimes today

`plan_inverted_tree` returns `catalog`, `deps`, and `scc_map`. `CatalogEdges`
is local and freed, but every `SeriesDeps.edges` tuple keeps the same
`DependenceEdge` objects alive through emission, and `index_maps`,
`affine_maps`, `aligned_ids`, `lookup_ids`, `lagged_ids`, and `keyed_ids` are
retained as well. Emission reads only `param_ids`, `is_scan`, and
`scan_direction` (`named_emit.py:395`, `:483`, `:734`; nothing in
`ast_emit.py` reads `ctx.deps`). `BoundSeries._emit_cache` (`catalog.py:164`)
accumulates one `CellRef` list per host cell (`ast_emit.py:266`) and is never
cleared, so all series' caches coexist until the catalog dies. In
`scripts/measure_named_export.py`, `generate_package` keeps the inventory's
`catalog`, `deps`, and `scc_map` alive while it runs a second full
`plan_inverted_tree` inside `generate_inverted_tree_modules`, doubling the
planning footprint on the measured path.

`_classify_formula_holes` calls `_stored_formula_addresses` (`catalog.py:1039`)
once per formula series with holes; each call re-opens the workbook and
iterates every row of every touched sheet. This is churn and time rather than
retained memory, but on a 5.5 MB `.xlsm` with 1,548 formula series it is a
repeated allocation spike and a plausible share of the 746 s.

## Safety contract

- Test-driven: each change starts with a failing test that pins the current
  generated bytes or the current lifetime (for example, "no `DependenceEdge`
  is reachable from the object passed to `emit_named_modules`"), then the
  change, then refactoring while green.
- A golden-output harness compares every generated module byte-for-byte
  against a pre-change snapshot for `tiny_dsa`, the unit fixtures that call
  `generate_inverted`, and the synthetic scaled workload. On the LIC DSF
  sandbox, compare the module SHA-256 values against
  `plans/named-axis-public-code-evidence.json` (`modules.*.sha256`).
- No change to `named_codegen_fingerprint` output; a streamed hash must
  reproduce the exact JSON bytes of the current `json.dumps(...,
  sort_keys=True, separators=(",", ":"))` payload and is tested for equality.
- Memory claims come from fresh-process measurements with both tracemalloc
  peak and `ru_maxrss`, reported separately, with sample counts.

## Work sequence

### 1. Make generator memory measurable

- [ ] Add `--profile-generation` to `scripts/measure_named_export.py`: run
  generation in a fresh subprocess, report per-phase tracemalloc retained and
  peak plus `ru_maxrss` at the boundaries of `build_catalog`,
  `collect_catalog_edges`, `collect_all_deps`, `build_scc_map`, the `plan_scc`
  loop, `assert_subgraph_bound`, `emit_named_internals`, `emit_named_api`,
  `emit_named_data`, and `build_runtime_modules`, and an object census
  (`DependenceEdge`, `KeyPoint`, `Statement`, `Domain`, `str` from
  `address_keys.py`). Record the graph's own footprint separately with
  `scripts/measure_graph_memory.py` on the pinned graph pickle.
- [ ] Add a synthetic scaled workload generator (`scripts/` or a test helper)
  parameterized by series count, horizon, scenario sheets, and formula mix,
  so scaling can be measured without the sandbox. Record 60x40, 200x60, and
  one multi-sheet keyed configuration.
- [ ] Record the LIC DSF baseline on the pinned inputs (`--reconcile-bindings
  --standalone`), with the same input hashes as the evidence report.

Deliverable: a baseline table attributing peak and retained bytes to phases
and object classes, on both synthetic and LIC DSF workloads.

### 2. Shorten lifetimes without changing output

- [ ] Drop planning-only fields before emission. After `plan_scc` and
  `assert_subgraph_bound`, `generate_inverted_tree_modules` passes emission a
  projection of `SeriesDeps` with `edges=()` and empty index and access maps,
  or `plan_inverted_tree` gains a `for_emission` return. `collect_all_deps`
  and `plan_inverted_tree` keep their full result for
  `tests/unit/exporter/inverted_tree/test_dependence_edges.py` and the
  inventory path. Expected: every `DependenceEdge` freed before emission.
- [ ] Scope `_emit_cache`: clear the host's cache at the end of
  `_semantic_body` (the `positions` and `("refs", cell)` entries are
  host-local by construction), or move the cache into `EmitContext` per
  series. Expected: emission retained grows by one series, not all.
- [ ] Fix the measurement path: in `generate_package`, `del` the inventory
  `catalog`, `deps`, `scc_map` (or move the inventory into a helper) before
  `generate_inverted_tree_modules` runs.
- [ ] Emit with a sink. Give `emit_named_modules` an optional writer so
  `internals.py`, `api.py`, and `data.py` can be written as they complete
  instead of all module strings coexisting alongside their `"\n".join`
  inputs. Keep `dict[str, str]` as the default return for `CodeGenerator`.
  Expected: peak lower by roughly the total generated size (about 10 MB for
  LIC DSF, 20 MB counting the join copies).

Deliverable: retained-after-planning and peak-during-emission figures from
step 1, before and after, with byte-identical modules.

### 3. Share what is duplicated

- [ ] Intern canonical addresses at the boundary. `as_canonical`,
  `canonical_address`, and `format_cell_key` return `sys.intern`ed plain
  `str` (the `NewType` is erased at runtime, so identity semantics are
  unchanged). Alternatively intern in `build_catalog` and `_DepCollector._emit`
  only, so the graph is untouched. Measure the string census before and
  after; the synthetic run suggests an order of magnitude fewer address
  strings.
- [ ] Slim `KeyPoint`s. `_key_point` (`catalog.py:806`) allocates one
  `KeyPoint`, one `items` tuple, and one `(field, value)` pair per key field
  for every bound cell (about 2.7 allocations per cell in the profile).
  Store the field names once per series and keep only the values tuple per
  cell, or share the pair tuples across cells with equal `(field, value)`;
  `__getitem__` and `as_mapping` keep their contract.
- [ ] Reduce per-cell schedule maps. `ScheduleIndex` keeps `coord_of`,
  `partition_of`, `axis_of`, `index_by_coord`, and `statement_id_by_coord`
  keyed by address; evaluate storing per-series arrays indexed by catalog
  position with a single address-to-(series, index) map, provided every
  reader (`schedule_coord`, `schedule_partition`, `schedule_axis_coord`) keeps
  its contract and `partition_catalog` still refreshes statements.

Deliverable: object census deltas and retained bytes per bound cell and per
edge, with byte-identical modules.

### 4. Flatten transient spikes and churn

- [ ] Cache `tensor_domain`, `coordinate_cells`, and `required_coordinates`
  on `BoundSeries` (lazy, computed once; `__post_init__` already resets
  per-instance caches so `dataclasses.replace` in `partition_catalog` stays
  correct). Hoist `series.required_coordinates` out of the per-coordinate
  generator in `emit_named_data` (`named_emit.py:1169`). Expected: fewer
  `Domain` constructions by orders of magnitude and a lower in-series
  high-water mark; also a large share of generation time.
- [ ] Stream `named_codegen_fingerprint`: hash entry by entry with the exact
  JSON framing the current payload uses, instead of building one list of
  dictionaries and one payload string covering all provenance (about 130,000
  coordinate-to-cell entries for LIC DSF). Test the digest for equality with
  the current implementation on every fixture.
- [ ] Compute `_provenance_source`'s explicit literal only when no rectangle
  or grid description wins, and bound `_grid_source`'s `3**len(fields)`
  assignment loop to build one mapping set at a time.
- [ ] Collect stored-formula candidates for all series first and call
  `_stored_formula_addresses` once per workbook (one open, one streamed pass
  per touched sheet), preserving the per-series hole classification.
- [ ] Check the `plan_scc` loop: `collect_dependence_edges` filters the full
  edge tuple once per SCC. Pre-bucket edges by consumer (already available as
  `CatalogEdges.by_consumer`) so each SCC reads only its members' edges. This
  is time rather than peak, but it removes one full-edge-list copy per SCC.

Deliverable: phase peak table from step 1 before and after, generation time,
and byte-identical modules.

### 5. Re-measure and record

- [ ] Re-run step 1 on the synthetic workloads and on LIC DSF. Report
  retained after planning, peak during emission, `ru_maxrss`, generation
  seconds, and the object census, with the change list that produced each
  delta.
- [ ] Run `uv run ruff check .`, `uv run ruff format --check .`,
  `uv run ty check`, `uv run pytest`, the Tiny-DSA CLI smoke check, and the
  public-computation contract tests; run `scripts/named_differential.py`
  against the regenerated LIC DSF package to confirm the module hashes and
  parity are unchanged.
- [ ] Save evidence in `plans/` beside this file and update this plan's status.

## Out of scope

- `DependencyGraph` construction and its retained size (nodes, ASTs, edge
  sets, guards, provenance). It is the largest resident structure during
  generation but belongs to the grapher; this plan only ensures codegen adds
  nothing to it and measures it separately.
- Import-time or evaluation memory of the generated package.
- Parallel or out-of-process emission.
- Any change that alters generated source, even whitespace.

## Checkpoints

After step 1, attribute the LIC DSF generator peak to graph, catalog, edges,
and emission transients; reorder steps 2 to 4 by measured share. After step
2, confirm no `DependenceEdge` survives into emission and the measured path no
longer plans twice. After steps 3 and 4, report bytes per bound cell and per
edge and the number of `Domain` constructions. If a step cannot reach
byte-identical output, stop it and record why rather than relaxing the
golden-output check.
