# Inverted-tree shape tests

Tier 1 distilled workbooks live in `test_shape_*.py` and the `_CORPUS`
oracle in `test_design_properties.py`. They run in every orientation and
forced rung.

## Distillation rule

When a workbook in the local pool (`tests/fixtures/local/`, #656) fails
the corpus gate, the **first-diverging series** in the statement-graph
topological order becomes a Tier 1 toy **before** the fix lands. Add it
to `_CORPUS` so the next regression cannot hide behind a missing shape.

First distilled pool shape: the country-table lookup from #654
(`INDEX(block, MATCH(country, names), col)` / sliding window / OFFSET),
in `test_shape_a26_index_block.py` as `country_table`.

#667 distilled shapes: `SUM` of a bound series (`range_sum` in
`test_shape_a27_range_aggregates.py`), plus whole-column / whole-row,
cross-sheet ranges, and `SUMPRODUCT` of covering series. `SUM(IF(...))`
stays fail-closed (#483).

## Local pool

`tests/fixtures/local/corpus.toml` lists pool workbooks. Workbooks
themselves are gitignored. Opt in with `pytest -m local_corpus`. A
missing workbook `pytest.skip`s with its path.
