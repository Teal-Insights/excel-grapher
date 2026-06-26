# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

<!-- version list -->

## v1.0.1 (2026-06-26)

### Bug Fixes

- **series-bindings**: Make setter docstrings layout-aware
  ([#302](https://github.com/Teal-Insights/excel-grapher/pull/302),
  [`bba3a1e`](https://github.com/Teal-Insights/excel-grapher/commit/bba3a1ec856a4c2b1197b3f7332703f59d01e110))

### Continuous Integration

- Skip CI on version bump
  ([`eb433b2`](https://github.com/Teal-Insights/excel-grapher/commit/eb433b27babee558853e0338985ec70a4fb29627))

### Documentation

- Clarify use_cached_dynamic_refs warning (#138)
  ([#301](https://github.com/Teal-Insights/excel-grapher/pull/301),
  [`7797ecb`](https://github.com/Teal-Insights/excel-grapher/commit/7797ecb51234c93d63e8d2478cfb5da4fd3bb9cb))


## v1.0.0 (2026-06-25)

### Bug Fixes

- **series-bindings**: Apply key dtype coercion after record normalization
  ([#294](https://github.com/Teal-Insights/excel-grapher/pull/294),
  [`2053688`](https://github.com/Teal-Insights/excel-grapher/commit/20536884106b368213a1d662327d91ba8617c7cf))

- **series-bindings**: Customize coercion error messages by layout
  ([#294](https://github.com/Teal-Insights/excel-grapher/pull/294),
  [`2053688`](https://github.com/Teal-Insights/excel-grapher/commit/20536884106b368213a1d662327d91ba8617c7cf))

- **series-bindings**: Reject duplicate composite keys in setter batches
  ([#294](https://github.com/Teal-Insights/excel-grapher/pull/294),
  [`2053688`](https://github.com/Teal-Insights/excel-grapher/commit/20536884106b368213a1d662327d91ba8617c7cf))

### Chores

- Update deprecated action ([#296](https://github.com/Teal-Insights/excel-grapher/pull/296),
  [`176b0a3`](https://github.com/Teal-Insights/excel-grapher/commit/176b0a3dce294d5def1610b252432de8561da2fe))

- **cursor**: Add conventional commit rule
  ([#294](https://github.com/Teal-Insights/excel-grapher/pull/294),
  [`2053688`](https://github.com/Teal-Insights/excel-grapher/commit/20536884106b368213a1d662327d91ba8617c7cf))

### Code Style

- **series-bindings**: Fix ruff and ty check issues
  ([#294](https://github.com/Teal-Insights/excel-grapher/pull/294),
  [`2053688`](https://github.com/Teal-Insights/excel-grapher/commit/20536884106b368213a1d662327d91ba8617c7cf))

### Continuous Integration

- Fix build command ([#297](https://github.com/Teal-Insights/excel-grapher/pull/297),
  [`776914c`](https://github.com/Teal-Insights/excel-grapher/commit/776914cccaf09042a9659448cbd01208eb265da1))

### Features

- Automate version bumps with semantic-release (closes #291)
  ([#292](https://github.com/Teal-Insights/excel-grapher/pull/292),
  [`2a3a8bd`](https://github.com/Teal-Insights/excel-grapher/commit/2a3a8bd3cccc75b32dd45ebb6798ebfde021ce07))

- **series-bindings**: Add flexible setter input coercion (closes #243)
  ([#294](https://github.com/Teal-Insights/excel-grapher/pull/294),
  [`2053688`](https://github.com/Teal-Insights/excel-grapher/commit/20536884106b368213a1d662327d91ba8617c7cf))

- **series-bindings**: Add matrix Layout and macro_matrix DataFrame test
  ([#294](https://github.com/Teal-Insights/excel-grapher/pull/294),
  [`2053688`](https://github.com/Teal-Insights/excel-grapher/commit/20536884106b368213a1d662327d91ba8617c7cf))

- **series-bindings**: Align SeriesInput type hints with DataFrame support
  ([#294](https://github.com/Teal-Insights/excel-grapher/pull/294),
  [`2053688`](https://github.com/Teal-Insights/excel-grapher/commit/20536884106b368213a1d662327d91ba8617c7cf))

- **series-bindings**: Extend setter smoke tests and add DataFrame example
  ([#294](https://github.com/Teal-Insights/excel-grapher/pull/294),
  [`2053688`](https://github.com/Teal-Insights/excel-grapher/commit/20536884106b368213a1d662327d91ba8617c7cf))

### Testing

- **series-bindings**: Add coercion parity tests and document input shapes
  ([#294](https://github.com/Teal-Insights/excel-grapher/pull/294),
  [`2053688`](https://github.com/Teal-Insights/excel-grapher/commit/20536884106b368213a1d662327d91ba8617c7cf))
