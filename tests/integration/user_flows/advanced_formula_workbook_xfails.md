# `advanced_formula_workbook` — known xfail gaps

Tracked in `advanced_formula_workbook_parity.py` (`XFAIL_*` sets and `DOWNSTREAM_CHAIN_CASES`).
Tests: `test_advanced_formula_workbook_parity.py`.

---

## `'Product Lookup'!K16` (NUMBERVALUE Demo)

**Issue:** https://github.com/Teal-Insights/excel-grapher/issues/264

Listed in `XFAIL_EVAL_CODEGEN` and `XFAIL_BINDING_EVAL`. The cell uses `NUMBERVALUE(TEXT(INDEX(...)))` to format a looked-up list price; the evaluator returns `#N/A` while standalone codegen and modular binding export return `3999.0`. Spot subgraph parity excludes this address from the passing set and `test_spot_subgraph_eval_codegen_known_gaps` / `test_spot_binding_eval_known_gaps` document the mismatch. Fix requires aligning `NUMBERVALUE`, `TEXT`, and error propagation with Excel when the INDEX/MATCH path succeeds.

---

## `'Product Lookup'!K24` (SUMPRODUCT analytics)

**Issue:** https://github.com/Teal-Insights/excel-grapher/issues/265

Listed in `XFAIL_EVAL_CODEGEN` and `XFAIL_BINDING_EVAL`. The formula `SUMPRODUCT(($E$5:$E$19>1000)*1)` should count products priced above 1000; the evaluator returns `0.0` and codegen returns `1.0` at the default workbook SKU. This is the same underlying SUMPRODUCT criteria semantics issue that breaks the `lookup_sku_to_sumproduct` downstream chain. Fixing boolean/criteria `SUMPRODUCT` in the shared runtime should clear both parity xfails and the chain xfail.

---

## `'Product Lookup'!K18` (partial graph overlap)

**Issue:** https://github.com/Teal-Insights/excel-grapher/issues/268

Listed in `XFAIL_BINDING_EVAL` only. The cell has no formula (empty leaf) but sits inside the `lookup_panel` binding `data_range`, so modular `compute_lookup_panel` emits `OBS_VALUE=0` while the evaluator returns `None`. Validation reports `partial_graph_overlap` for `lookup_panel` (12 cells skipped during unique-key checks). Tightening the binding range to formula-only rows (e.g. `K6:K17`) or excluding non-formula leaves from the series should resolve this without runtime changes.

---

## `'Product Lookup'!K19` (partial graph overlap)

**Issue:** https://github.com/Teal-Insights/excel-grapher/issues/266

Listed in `XFAIL_BINDING_EVAL` only. Like K18, this is an empty cell included in the `lookup_panel` output range, producing binding `0` vs evaluator `None`. It is part of the same `partial_graph_overlap` hygiene warning on the Product Lookup panel shard. Trimming the YAML `data_range` to match actual formula leaves is the expected fix.

---

## `lookup_sku_to_sumproduct` (downstream chain)

**Issue:** https://github.com/Teal-Insights/excel-grapher/issues/267

`DOWNSTREAM_CHAIN_CASES` entry with `xfail=True` and reason *"SUMPRODUCT analytics returns #VALUE! after lookup SKU change"*. After `set_lookup_sku` to `PRD-001`, the first hop (`compute_lookup_panel` / Product Name) passes, but `compute_sumproduct_analytics` for Software Revenue Potential returns `#VALUE!` instead of the expected `0.0`. Exercised by `test_downstream_propagation_chain[lookup_sku_to_sumproduct]` as a strict xfail. Likely fixed together with K24 once SUMPRODUCT criteria and any dependent range refs are correct.
