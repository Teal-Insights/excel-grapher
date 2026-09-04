# Tiny DSA fixture

Committed inverted-tree canary workbook with `OFFSET` / `INDEX` / `INDIRECT`
that need `DynamicRefConfig` constraints.

Smoke-check bindings with the colocated constraints module:

```bash
uv run excel-grapher bindings validate \
  tests/fixtures/inverted_tree/tiny_dsa/tiny-dsa.xlsx \
  --bindings tests/fixtures/inverted_tree/tiny_dsa/bindings \
  --constraints tests/fixtures/inverted_tree/tiny_dsa/constraints.py \
  --paradigm inverted_tree --smoke-test
```

`constraints.py` exposes `CONSTRAINTS: Mapping[str, type]`, the same contract
as `tests/fixtures/local/corpus.toml` entries.
