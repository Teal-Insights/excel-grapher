# Dead / obsolete code audit

Method: `vulture`, a package-wide reachability walk from the public API
(`excel_grapher/**/__init__.py` `__all__`, the CLI, and the wholesale-shipped
`inverted_tree/runtime.py`), a full `pytest --cov` run (81.7% line coverage),
and manual review of every candidate. "Test-only" means the symbol is exercised
by `tests/` but nothing in the package, CLI, docs, or scripts reaches it.

Sizes are approximate line counts of the removable code, not counting the tests
that would go with it.

## Tier 1: obsolete leftovers of removed representations (high confidence)

### 1. Address-keyed `compute_*` codegen and its docstring subsystem (~2,200 lines)

`CodeGenerator.generate()` was removed in #805 and input setters in #840, but the
library-level codegen that served that representation is still exported from
`excel_grapher.series_bindings` and `excel_grapher.exporter`:

| Module | Lines | Note |
| --- | --- | --- |
| `series_bindings/compute_codegen.py` | 451 | Emits `from excel_grapher.runtime.cache import EvalContext, xl_cell` and `ctx = make_context(inputs)`; `make_context` no longer exists anywhere, so emitted code cannot run. |
| `series_bindings/reader_index.py` | 365 | Only consumer was `compute_codegen`. `ReaderFallbackReason`, `reader_index_as_discovery_dicts` referenced nowhere. |
| `series_bindings/output_helper_index.py` | 228 | Only consumer was `compute_codegen`. |
| `series_bindings/docstrings.py` | 393 | Only consumer was `compute_codegen`. User guide says the callbacks "are not wired into `generate_modules()`". |
| `series_bindings/docstring_renderers.py` | 492 | Same. |
| `series_bindings/codegen_literals.py` | ~120 of 151 | `emit_setter_type_alias_lines`, `emit_compute_preamble_lines`, `setter_input_annotation` are setter/compute leftovers. |
| `series_bindings/setter_input_types.py` | 30 | `SetterInput` alias; `Layout`/`SeriesInput` can move to `input_coerce.py`. |
| `evaluator/name_utils.py` | 50 of 90 | `address_to_python_name`, `excel_func_to_python` were the address-keyed name manglers. |

Also: `ConstantSeries.reader_name`, `workflow.reader_names`, `groups._reader_name`
/ `grouped_public_names`, the `ConstantDirection.reader` schema property, and the
`row_series` / top-level `setter` / `input.setter` / `input.reader` stripping in
`normalize.normalize_series_entry`. The project is greenfield; 13 test files still
pass `input: {"setter": ...}` and rely on the stripping.

Docs to update when removed: `user_guide/06-export.qmd` "Series docstring
callbacks", `user_guide/08-contributing.qmd:70`, the `AGENTS.md` line "Series-binding
codegen defaults to the `google` renderer".

### 2. `embed.emit_runtime` cache scaffold modes (~450 lines)

The only production caller is `inverted_tree/standalone.py`, which passes an
explicit `modules=` list and `include_offset_table=False,
include_dep_tracking=False, include_operators_fastpath=False`. Everything behind
the defaults is exercised only by tests:

- `embed.py`: `_RUNTIME_MODULES`, `_EXPORT_RUNTIME_MODULES`, `_CORE_MODULES`,
  `_ALL_MODULES`/`_ALL_MODULE_NAMES`, `_SLIM_CACHE_EVAL_*`, `runtime_cache_seed_symbols`,
  the whole `include_dep_tracking` branch, and the `include_offset_table` parameter
  (100% unused per vulture).
- `runtime/cache_eval_slim.py` (93 lines, never executed by tests either).
- `runtime/cache.py`: `circular_safe_cache` (referenced nowhere), `xl_memoize` (test-only).
- Tests/fixtures that exist only for this: `tests/unit/exporter/test_dep_tracking_emit_runtime_scaffold.py`,
  `tests/unit/exporter/test_emit_runtime.py`, `tests/fixtures/dep_tracking_baseline/`,
  the `SLIM_CACHE_EVAL_SCAFFOLD_LINE_BUDGET` / `count_dep_tracking_lines` helpers in
  `tests/integration/utils/parity_harness.py`.

### 3. Rung / fused-loop scheduling in `inverted_tree/schedule.py` and `deps.py` (~700 lines)

After the named-axis migration (#837) `emit.py` says "Every series is emitted as
named coordinate code; recurrences are demand-driven". `plan_scc(...)` is still
called from `plan_inverted_tree`, but its `SccPlan` return value is discarded; the
only surviving effect is the `InvertedTreeExportError` raised by
`assert_distance_zero_legal` inside `plan_fused_scc`. Everything that computes the
rung and the fused loop is now vestigial:

- `schedule.py`: `SccPlan`, `Rung`, `FusedPlan`, `FusedRegion`, `plan_fused_scc`
  (105), `_fuse_regions`, `_contiguous_domain`, `has_residual_may_cycle`,
  `residual_body_order` (test-only), `IndexSet` (212, only reached from the dead
  `deps.plan_indices` cluster), `indices_to_source` + the five `_*_source` helpers,
  `IndexSourceIntern` + `_INDEX_INTERN`, `_REPEAT_MIN`, `_FAMILY_MIN`.
  `emit.py`'s `force_rung: Literal[2, 3] | None` parameter is documented as
  "accepted for compatibility and ignored".
- `deps.py`: `plan_indices` (104, test-only), `_propagate_param_indices`, `_scc_is_nested`,
  `_topo_units`, `_topo_sort`, `predecessor_closure`, `formula_closure`,
  `collect_all_dependence_edges` (all unreachable from production).
  Also referenced nowhere: `iter_cross_sheet_addresses`, `eval_host_selector` (52),
  `ref_window_corners`, `shift_range_corners`, `offset_expr_exclude_addresses`,
  `refine_access_classes`, `lagged_ids`/`keyed_ids` fields on `SeriesDeps`.
- `catalog.py`: `SeriesHole`, `build_schedule_index` (91), `partition_catalog`
  (test-only), `bound_addresses`, `is_time_series`, `hole_indices`,
  `require_key_point_for`, `require_binds_for`.
- `access.py`: `AxisKind`.

Replace the `plan_scc` loop in `plan_inverted_tree` with a direct
`assert_distance_zero_legal` call and the rest can go.

### 4. Legacy serialization paths (~200 lines)

- `grapher/graph.py` `__getstate__` / `__setstate__` (80 lines). `__reduce_ex__`
  now pickles via `graph_pickle`; the state-dict path is reached only by
  `tests/unit/grapher/test_graph_serialization.py` crafting a pickle by hand.
- `grapher/cache.py:479-488` schema-7 `normalized_formula` fallback. The cache
  loader rejects any stored `schema_version != 9`, so this branch is unreachable.
- `core/formula_ast_json.py` `legacy_a1` handling for `{"t": "cell", "v": "A1"}`.
  No producer emits that shape and no test exercises it.
- `export_runtime/tensor.py` `from_legacy` / `to_legacy` adapters (test-only) and
  `YearSeries` / `ScenarioSeries` (test-only; nothing in the package or emitted code
  constructs them).
- `grapher/node.py` `_finalize_cell_legacy` (23) is the `Node(sheet=, column=, row=)`
  constructor path; 3 package call sites vs 105 in tests. Low priority.

## Tier 2: referenced nowhere at all (safe deletes)

| Location | Symbol | Lines |
| --- | --- | --- |
| `grapher/dynamic_refs.py:3020` | `_infer_choose_numeric_domain` | 47 |
| `grapher/dynamic_refs.py:384,407` | `_expand_sheet_qualified_range`, `_strip_optional_sheet_prefix` | 29 |
| `grapher/dynamic_refs.py:2233,3781` | `_domain_with_max`, `_narrowed_branch_is_infeasible` (never executed) | 35 |
| `series_bindings/workflow.py:137` | `output_binding_covered_addresses` | 27 |
| `exporter/inverted_tree/named_emit.py:1379` | `inventory_named_emission` (only `scripts/measure_named_export.py`) | 43 |
| `exporter/inverted_tree/emit.py:138` | `_HOLE_DOC_LABELS` | 7 |
| `exporter/export_runtime/errors.py:53` | `raise_if_sentinel` | 5 |
| `exporter/export_runtime/values.py:30` | `_convergence_delta` (duplicate of `runtime/cache.py`) | 20 |
| `exporter/embed.py:128` | `_ImportCollector.visit_AnnAssign` | 3 |
| `grapher/resolver.py:403` | `qualify_cell_ref` | 3 |
| `grapher/compression.py:499` | `ensure_snapshot` | 3 |
| `grapher/parser.py:60,521,522` | `_QUOTED_SHEET`, `_needs_quoting`, `_format_ref` ("historical internal names") | 3 |
| `grapher/builder.py:99,104` | `_logger`, `_VOLATILE_DYNAMIC_REF_FUNCS` | 4 |
| `core/cell_types.py:15,248` | `CellKind.DATE`, `_normalize_cell_address` alias | 2 |
| `core/addressing.py:152` | `_split_sheet_qualified_address` alias | 1 |
| `series_bindings/output_helper_index.py:23` | `OutputHelperFallbackReason` | 1 |
| `grapher/type_analysis_cache.py` | `TypeAnalysisCacheStats.disabled` never read | 1 |
| `examples/micro_workbooks/~$ffv2.xlsx` | committed Excel lock file | file |

## Tier 3: production code that exists only for tests

These are exported or module-level names whose only callers are tests. Either
delete them together with their tests, or move the logic into `tests/utils/`.

| Location | Symbol | Lines |
| --- | --- | --- |
| `grapher/solver_mcve.py` | whole module (`load_solver_mcve` etc.); consumers are `tests/unit/grapher/test_solver_mcve.py` and `scripts/analyze_solver_mcve.py` | 253 |
| `core/formula_shape.py` | `resolve_address_leaf`, `encode_address_leaf`, `iter_address_holes`, `_encode_*_axis`, `FormulaShapeSummary`, `summarize_normalized_formulas`, `summarize_formula_shapes`, `clear_shape_parse_cache`, `shape_parse_cache_info` | 248 |
| `grapher/parser.py` + `core/formula_normalization.py` | regex path: `parse_range_refs` (86), `normalize_formula` (23), `normalize_excel_formula` (27), `PreparedFormula`/`prepare_formula` (24). Docstring already calls it "transitional; not the AST render dialect". | 160 |
| `grapher/lightweight_viz.py` | `lightweight_viz_flat` + `LightweightVizFlat` (test-only); `_force_directed_xy` (96), `_grid_xy`, `_default_bfs_target_ranks`, `_node_for_viz_subgraph`, `_induced_dependency_subgraph`, `estimate_serialized_json_bytes` never executed | 290 |
| `grapher/guard.py` | `evaluate_guard`, `_eval_guard_atom`, `_UNKNOWN` (tests + `scripts/analyze_solver_mcve.py`); `guard_intern_pool_size`, `clear_guard_intern_pool` | 65 |
| `core/formula_ast.py:338` | `iter_resolved_cell_keys` | 22 |
| `core/operators_reference.py:20` | `broadcast_pair` | 20 |
| `core/operators.py:280` | `OPERATOR_TABLE` (only the parity test against `inverted_tree/runtime.OPERATOR_TABLE`; keep if the parity test is wanted) | 15 |
| `grapher/range_compression/index.py` | `find_dependents`, `find_precedents` (plus never-executed `_find_*_range`) | 40 |
| `grapher/range_compression/ref_parser.py:135` | `parse_cell_refs_with_abs` | 7 |
| `grapher/dynamic_refs.py:184,267` | `_apply_constraint_to_schema`, `DynamicRefConfig.from_constraints_and_workbook` | 75 |
| `grapher/node.py:138,143` | `_derived_fields_cache_info`, `_derived_fields_cache_clear` | 6 |
| `evaluator/evaluator.py:208,214` | `clear_caches`, `ast_cache_info` | 8 |
| `exporter/inverted_tree/runtime.py` | `require_length`, `live_measure`, `demand_instance` (shipped into generated packages, so arguably public; nothing emitted calls them) | 22 |
| `exporter/semantic_viz.py:53,65` | `semantic_viz_renderer`, `spread_rank_centers` | 50 |
| `series_bindings/docstrings.py:111` | `run_series_docstring_callback` | 6 |
| `core/cell_types.py:29` | `IntIntervalDomain = IntervalDomain` "backwards-compatible alias" (37 test uses) | 1 |

## Tier 4: repository clutter

- `plans/`: 17 files, 300 KB of evidence snapshots for merged PRs. Only
  `inverted-tree-scheduling.md` (cited by 3 tests) and
  `named-axis-chart-error-contract.json` (default for `scripts/named_differential.py`)
  are referenced. `downstream-named-axis.patch` is a 46 KB patch against a
  *different* repository (`src/codegen_cache.py`).
- `scripts/measure_formula_ast_intern.py`, `measure_formula_shapes.py`,
  `measure_leaf_store.py`, `measure_graph_memory.py`: one-off measurements for
  closed issues (#517, #550, #579). `measure_formula_shapes.py` and
  `analyze_solver_mcve.py` are the only things keeping the Tier 3 `formula_shape`
  summaries and `solver_mcve` alive.
- `AGENTS.md:29-30` and `great-docs.yml:379` reference `scripts/great_docs_build.py`
  and `scripts/great_docs_preview.py`, which do not exist.
- `README_files/lightweight_viewer.png` is unreferenced (the user guide uses
  `user_guide/_files/lightweight_viewer.png`).
- `examples/micro_workbooks/build_*_workbook.py`, `demo_taco_index.py`,
  `demo_cross_sheet_taco_index.py` are not referenced by tests or docs;
  `demo_taco_index.py` is the only reason `matplotlib` is a dev dependency.
  `nbformat` and `jupyter` dev dependencies are imported nowhere.
- `evaluator/__init__.py` `__getattr__` lazy import of `FormulaEvaluator` ("keep the
  package importable while modules are developed") and the
  `_LAZY_WORKFLOW_EXPORTS` `__getattr__` in `series_bindings/__init__.py`: both
  modules import eagerly without a cycle, so plain imports work.
- Unused parameters flagged by `ruff --select ARG`: `emit.generate_inverted_tree_modules(force_rung)`,
  `embed.emit_runtime(include_offset_table)`, `embed._format_emitted_symbol_source(symbol)`,
  `dynamic_refs._enumerate_assignments(limits)`, `_enumerate_value_assignments(limits)`,
  `infer_index_targets(bounds)`, `graph_consistency._keys_for_range(graph)`,
  `parser._parse_ref_or_range_token(current_sheet)`, `range_compression.build._stream_expected_deps(ref_idx)`,
  `ref_parser._parse_range_text(default_sheet)`, `resolver._eval_indirect_formula_to_range(get_cell_value)`,
  `target_expansion._require_sheet(target)`, `standalone._rewrite_imports(tree)`,
  `formula_ast._try_parse_local_whole_row(original)`.
- 14 `ERA001` commented-out code blocks in tests (`test_functions.py`,
  `test_dynamic_refs.py`, `test_perf_optimizations.py`,
  `test_type_analysis_cache_integration.py`).

## Not dead, but worth knowing

- `excel_grapher/runtime/*` (evaluator) and `exporter/export_runtime/*` (emitted
  code) deliberately duplicate `xl_choose`, `xl_address`, `xl_row`, etc. with
  sentinel-returning vs raising signatures, and `inverted_tree/runtime.py` carries a
  third `OPERATOR_TABLE`. That is the parity design in `AGENTS.md`, not cruft.
- `core/operators_fastpath.py` and `evaluator/shape_eval.py` `_compile_*` helpers
  show as unreachable or unexecuted only because they are loaded through
  `try/except ImportError` and dispatch tables.
- Many `export_runtime` worksheet functions are never executed by tests. They are
  emitted on demand and are missing test workbooks, not dead.

## Suggested order

1. Tier 2 (mechanical, no behavior change).
2. Tier 1.3 (`plan_scc` → direct legality check) and Tier 1.2 (`emit_runtime` defaults).
3. Tier 1.1 with a docs pass, since it removes public exports.
4. Tier 3 case by case; Tier 4 housekeeping any time.
