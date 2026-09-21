# Tiny DSA fixture

Committed inverted-tree canary workbook with `OFFSET` / `INDEX` / `INDIRECT`.
Series `domain` (and `constant` `from_workbook` pins) compile to the
`CellTypeEnv` dynamic-ref inference consumes.

Smoke-check bindings from the sidecar:

```bash
uv run excel-grapher bindings validate \
  tests/fixtures/inverted_tree/tiny_dsa/tiny-dsa.xlsx \
  --bindings tests/fixtures/inverted_tree/tiny_dsa/bindings \
  --smoke-test
```

Graph building derives domains from the sidecar.
