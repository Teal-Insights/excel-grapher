# Test suite layout

Tests are grouped **by kind first**, then by package mirror:

- [`tests/unit/`](unit/) — narrow, refactor-friendly tests (`core/`, `evaluator/`, `grapher/`, `exporter/`, plus root `test_package_boundaries.py`).
- [`tests/integration/`](integration/) — behavior and contract tests (same subdirs, plus `examples/`, `utils/`).

Shared helpers live in [`tests/utils/`](utils/) or a `utils/` subdir near the tests they serve.

## Unit vs integration

- **`unit`**: Narrow tests over a single component or small synthetic graph. Safe to refactor heavily; prefer these for fast feedback.
- **`integration`**: Behavior-level checks—workbooks, evaluator↔Excel cache parity, codegen integration, example contracts, smoke coverage of the public API, Excel automation helpers.
