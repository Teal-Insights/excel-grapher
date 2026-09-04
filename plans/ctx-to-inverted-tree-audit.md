# ctx → inverted-tree feature audit (issue 662)

Pin: `3fe65d3` (12.7.2) plus this change. Probes live in
`tests/unit/exporter/inverted_tree/test_ctx_parity_audit.py`.

**Gate for flipping the library default** (all three):

1. Local workbook pool (#656 §12) green on tiny-DSA, Q-CRAFT, **and LIC-DSF**.
   tiny-DSA is committed; Q-CRAFT 796/796 via #659; LIC-DSF has not been through
   the gate. **Default stays `ctx`.**
2. #651 landed (shared comparison / operator core).
3. This table decided; `port` rows shipped or linked.

**Deprecation sequence** (after the gate, not this release): default flip +
`DeprecationWarning` on `paradigm="ctx"` for one release, then removal. Documented
in `user_guide/06-export.qmd`.

Legend: **port** = inverted-tree should grow the feature; **bindings-equivalent**
= already true at sidecar/catalog load, or a documented caller recipe;
**drop** = ctx-only by design.

| Feature | ctx | inverted-tree | decision |
| --- | --- | --- | --- |
| `set_<series>` input setters, `input.mode` for non-leaf inputs | yes | inputs are arguments | **drop** — by design (#597). Override-mode cells should be bound as `input`. |
| `read_<series>` / `input.reader` | yes | none | **drop** — import `data.py` or pass the value. |
| `input.domain` on setters | yes | yes (#666) | **port shipped** — `require_input_domain` on `compute_*` / `_run_N` arguments. |
| `make_context` / `inputs=` overlay | yes | none | **drop** — by design. |
| `output.compute.helper` | yes | every output is a leaf-closure function | **drop**. |
| Tidy `Records` / `as_records` | yes | tuples only | **bindings-equivalent** — documented recipe in `06-export.qmd`. |
| Concept-based naming (#379) | partial | series `id` | **port** — keep [#379](https://github.com/Teal-Insights/excel-grapher/issues/379); applies to both paradigms. |
| Series-keyed object façade (#593) | planned | none | **port** — retarget [#593](https://github.com/Teal-Insights/excel-grapher/issues/593) at inverted-tree (`compute_*` wrapper, not `EvalContext`). |
| `constant` direction `read_*` | yes | imported from `data` | **port shipped** — [#663](https://github.com/teal-insights/excel-grapher/issues/663) (`data.CONSTANT_X` / `data.overrides`). |
| `CONSTANTS` `MappingProxyType` (#582); sparse leaf store (#578) | yes | `data.py` tuples | **drop** — dense catalog-order tuples are the inverted-tree store. |
| Range-watch export invalidation (#585) | yes | n/a | **drop** — no mutable `EvalContext`; pass new arguments. |
| Complementary shards / list `data_range` (#591) | yes | catalog concatenates `series_data_ranges` after merge | **bindings-equivalent**. Probe: sheet_name keys and four-cell catalog; a formula may still fail closed if it gathers two non-adjacent members of that series (same as any series). |
| `sheet_name`, `value_map`, `datetime` / `bool` dtypes | yes | catalog via `resolve_key_domain`; bool measured; datetime leaves emitted | **bindings-equivalent** (keys) + **port shipped** (`datetime` literals in `data.py`). |
| `exclude_rows` / `exclude_columns` | yes | yes (#600) | done. |
| `layout: matrix` | yes | yes (#599, #612, #638) | done. |
| Named ranges as `data_range` | yes | yes (`expand_data_range`) | **bindings-equivalent**. Probe green vs `FormulaEvaluator`. |
| Named ranges in formulas | expanded to A1 before parse | same, then cell-ref lowering | **bindings-equivalent** when the expansion is a bound cell. |
| Whole-column / whole-row refs | yes | covering series + used-range expansion | **port shipped** ([#667](https://github.com/Teal-Insights/excel-grapher/issues/667)). Unbound used-range cells still fail closed. |
| Cross-sheet / 3-D ranges | yes | `sheet_order` between endpoints; `xl_sum` of bound cells | **port shipped** ([#667](https://github.com/Teal-Insights/excel-grapher/issues/667)). |
| Unions / `SUM` of a range | yes | `xl_sum` / `xl_sumproduct` of covering series (`take` for a window) | **port shipped** ([#667](https://github.com/Teal-Insights/excel-grapher/issues/667)). `SUM(IF(…))` stays fail-closed ([#483](https://github.com/Teal-Insights/excel-grapher/issues/483)). |
| `INDIRECT` | `DynamicRefConfig` + runtime | graph may resolve edges; AST still `INDIRECT` → fail closed | **port** as graph-derived access (INDEX/OFFSET pattern from #656) when a corpus needs it. |
| Array formulas / spill (#284), `SUMPRODUCT`, `SUM(IF(…))` (#483) | yes / partial | `SUMPRODUCT` of covering series; spill / `SUM(IF)` fail closed | **port** mechanical `SUMPRODUCT` shipped ([#667](https://github.com/Teal-Insights/excel-grapher/issues/667)); spill / `SUM(IF)` remain [#284](https://github.com/Teal-Insights/excel-grapher/issues/284) / [#483](https://github.com/Teal-Insights/excel-grapher/issues/483). |
| Excel function coverage | `core/` + `export_runtime/` library | generic `xl_<name>` only when `runtime.py` defines it; `IF`/`CHOOSE`/`INDEX`/`MATCH`/`OFFSET`/`EXP`/`SUM`/`SUMPRODUCT` special-cased | **port** incrementally, fail closed otherwise. Do not embed the ctx library. |
| Error semantics (#326) | `XlErrorException` raise-only | `XlError` at operators; series members store `err.code` measures | **drop** as a ctx convention. Inverted-tree's measure+raise split is intentional (#597 A7). [#326](https://github.com/Teal-Insights/excel-grapher/issues/326) remains ctx-export work. |
| Pass-1 semantic internals (#595) | yes | inverted-tree helpers replace them | **drop**. Pass-2 LLM naming stays a pipeline concern. |
| `--smoke-test` in `bindings validate` | setter + Records compute smoke | skipped → **shipped**: calls `compute_*` with `data.py` defaults | **port shipped**. |
| Projection / `embed` / viz | separate | unaffected | **drop** from this audit. |

## Port follow-ups (not shipped here)

1. **`INDIRECT` graph-derived access** — [#668](https://github.com/Teal-Insights/excel-grapher/issues/668).

Shipped: #666 (`input.domain`), #667 (range aggregates / remaining refs). Already filed: #663 (constants), #593 (façade), #379 (concept names), #284 (spill), #483 (`SUM(IF)`).
