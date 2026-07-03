# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

<!-- version list -->

## v2.2.0 (2026-07-03)

### Chores

- Relicense project under MIT ([#330](https://github.com/Teal-Insights/excel-grapher/pull/330),
  [`a82aadc`](https://github.com/Teal-Insights/excel-grapher/commit/a82aadce07570ead3142e48d88b3e83a17de25ee))

### Features

- **series_bindings**: Add optional view-level groups for export sequencing (#308)
  ([#332](https://github.com/Teal-Insights/excel-grapher/pull/332),
  [`5d2e610`](https://github.com/Teal-Insights/excel-grapher/commit/5d2e6103b588ec158306d5eb6d0364f75921c39c))


## v2.1.3 (2026-07-03)

### Bug Fixes

- Enforce raise-only export boundary
  ([#327](https://github.com/Teal-Insights/excel-grapher/pull/327),
  [`8991987`](https://github.com/Teal-Insights/excel-grapher/commit/899198771ad9e5c37d359abaa40d21d7f9bb31de))

- Satisfy export runtime type check
  ([#327](https://github.com/Teal-Insights/excel-grapher/pull/327),
  [`8991987`](https://github.com/Teal-Insights/excel-grapher/commit/899198771ad9e5c37d359abaa40d21d7f9bb31de))

### Code Style

- Sort export runtime imports ([#327](https://github.com/Teal-Insights/excel-grapher/pull/327),
  [`8991987`](https://github.com/Teal-Insights/excel-grapher/commit/899198771ad9e5c37d359abaa40d21d7f9bb31de))

### Testing

- Update if codegen boundary assertion
  ([#327](https://github.com/Teal-Insights/excel-grapher/pull/327),
  [`8991987`](https://github.com/Teal-Insights/excel-grapher/commit/899198771ad9e5c37d359abaa40d21d7f9bb31de))


## v2.1.2 (2026-07-03)

### Bug Fixes

- **exporter**: Emit TypeAlias for SeriesInput in exported code
  ([#325](https://github.com/Teal-Insights/excel-grapher/pull/325),
  [`514d3cb`](https://github.com/Teal-Insights/excel-grapher/commit/514d3cb0240d5ecd8423faf3176c7dd58c536de2))


## v2.1.1 (2026-07-03)

### Bug Fixes

- Cache repeated dynamic ref expansions
  ([#324](https://github.com/Teal-Insights/excel-grapher/pull/324),
  [`e92f979`](https://github.com/Teal-Insights/excel-grapher/commit/e92f979b9f8cbc131cd345f321d17300215ae46e))

- Stop caching dynamic-ref mask spans in dep cache
  ([#324](https://github.com/Teal-Insights/excel-grapher/pull/324),
  [`e92f979`](https://github.com/Teal-Insights/excel-grapher/commit/e92f979b9f8cbc131cd345f321d17300215ae46e))

### Code Style

- Apply ruff format to builder and dynamic refs tests
  ([#324](https://github.com/Teal-Insights/excel-grapher/pull/324),
  [`e92f979`](https://github.com/Teal-Insights/excel-grapher/commit/e92f979b9f8cbc131cd345f321d17300215ae46e))


## v2.1.0 (2026-07-03)

### Features

- **exporter**: Emit list_setters/list_computes discovery helpers
  ([#323](https://github.com/Teal-Insights/excel-grapher/pull/323),
  [`85bef61`](https://github.com/Teal-Insights/excel-grapher/commit/85bef61936ba40ff87c7834736d8ce31841357aa))

- **series-bindings**: Add empty_measure and matrix DataFrame setter input
  ([#323](https://github.com/Teal-Insights/excel-grapher/pull/323),
  [`85bef61`](https://github.com/Teal-Insights/excel-grapher/commit/85bef61936ba40ff87c7834736d8ce31841357aa))

- **series-bindings**: Empty_measure knob and matrix DataFrame setter ergonomics
  ([#323](https://github.com/Teal-Insights/excel-grapher/pull/323),
  [`85bef61`](https://github.com/Teal-Insights/excel-grapher/commit/85bef61936ba40ff87c7834736d8ce31841357aa))

### Refactoring

- **series-bindings**: Tighten empty_measure setter ergonomics
  ([#323](https://github.com/Teal-Insights/excel-grapher/pull/323),
  [`85bef61`](https://github.com/Teal-Insights/excel-grapher/commit/85bef61936ba40ff87c7834736d8ce31841357aa))


## v2.0.0 (2026-07-02)

### Bug Fixes

- **exporter**: Raise NA() as an error literal to preserve parity
  ([#318](https://github.com/Teal-Insights/excel-grapher/pull/318),
  [`9c9c9b7`](https://github.com/Teal-Insights/excel-grapher/commit/9c9c9b7e2ab6fddf182bf054255408330752d837))

### Documentation

- **parity**: Correct list of thunked error-consuming functions
  ([#318](https://github.com/Teal-Insights/excel-grapher/pull/318),
  [`9c9c9b7`](https://github.com/Teal-Insights/excel-grapher/commit/9c9c9b7e2ab6fddf182bf054255408330752d837))

### Features

- Add export runtime scaffolding ([#318](https://github.com/Teal-Insights/excel-grapher/pull/318),
  [`9c9c9b7`](https://github.com/Teal-Insights/excel-grapher/commit/9c9c9b7e2ab6fddf182bf054255408330752d837))

- Inline operators in exported codegen
  ([#318](https://github.com/Teal-Insights/excel-grapher/pull/318),
  [`9c9c9b7`](https://github.com/Teal-Insights/excel-grapher/commit/9c9c9b7e2ab6fddf182bf054255408330752d837))

- Migrate exported range consumers onto lazy Range
  ([#318](https://github.com/Teal-Insights/excel-grapher/pull/318),
  [`9c9c9b7`](https://github.com/Teal-Insights/excel-grapher/commit/9c9c9b7e2ab6fddf182bf054255408330752d837))

- Pythonic exported runtime — lazy ranges and raise-based errors
  ([#318](https://github.com/Teal-Insights/excel-grapher/pull/318),
  [`9c9c9b7`](https://github.com/Teal-Insights/excel-grapher/commit/9c9c9b7e2ab6fddf182bf054255408330752d837))

- Raise Excel errors as exceptions in exported code
  ([#318](https://github.com/Teal-Insights/excel-grapher/pull/318),
  [`9c9c9b7`](https://github.com/Teal-Insights/excel-grapher/commit/9c9c9b7e2ab6fddf182bf054255408330752d837))

### Performance Improvements

- **exporter**: Bind operands once in array-operator guards
  ([#318](https://github.com/Teal-Insights/excel-grapher/pull/318),
  [`9c9c9b7`](https://github.com/Teal-Insights/excel-grapher/commit/9c9c9b7e2ab6fddf182bf054255408330752d837))

- **exporter**: Only guard operands that can evaluate to arrays
  ([#318](https://github.com/Teal-Insights/excel-grapher/pull/318),
  [`9c9c9b7`](https://github.com/Teal-Insights/excel-grapher/commit/9c9c9b7e2ab6fddf182bf054255408330752d837))

### Refactoring

- **export-runtime**: Drop dead code in export runtime
  ([#318](https://github.com/Teal-Insights/excel-grapher/pull/318),
  [`9c9c9b7`](https://github.com/Teal-Insights/excel-grapher/commit/9c9c9b7e2ab6fddf182bf054255408330752d837))

- **export-runtime**: Use canonical sheet-name quoting in OFFSET
  ([#318](https://github.com/Teal-Insights/excel-grapher/pull/318),
  [`9c9c9b7`](https://github.com/Teal-Insights/excel-grapher/commit/9c9c9b7e2ab6fddf182bf054255408330752d837))

### Testing

- Align dep-tracking baseline assertion with regenerated fixture
  ([#318](https://github.com/Teal-Insights/excel-grapher/pull/318),
  [`9c9c9b7`](https://github.com/Teal-Insights/excel-grapher/commit/9c9c9b7e2ab6fddf182bf054255408330752d837))

- Satisfy export runtime hook checks
  ([#318](https://github.com/Teal-Insights/excel-grapher/pull/318),
  [`9c9c9b7`](https://github.com/Teal-Insights/excel-grapher/commit/9c9c9b7e2ab6fddf182bf054255408330752d837))


## v1.2.0 (2026-07-02)

### Documentation

- **user_guide**: Grouped-row matrix geometry semantics and decision table
  ([#321](https://github.com/Teal-Insights/excel-grapher/pull/321),
  [`e71d121`](https://github.com/Teal-Insights/excel-grapher/commit/e71d121b307fb7e7b9559e68b13e9d90dfe730cb))

### Features

- **series_bindings**: Grouped-row matrix geometry (schema 1.5.0)
  ([#321](https://github.com/Teal-Insights/excel-grapher/pull/321),
  [`e71d121`](https://github.com/Teal-Insights/excel-grapher/commit/e71d121b307fb7e7b9559e68b13e9d90dfe730cb))

### Refactoring

- **grapher**: Make TACO index a derived artifact, not graph state
  ([`6370794`](https://github.com/Teal-Insights/excel-grapher/commit/637079425e6402aff6a6b5d1ec53becb66dd380d))


## v1.1.0 (2026-06-27)

### Features

- **exporter**: Emit list_setters/list_computes discovery helpers
  ([#305](https://github.com/Teal-Insights/excel-grapher/pull/305),
  [`93b71d4`](https://github.com/Teal-Insights/excel-grapher/commit/93b71d4f39f7388b3d5d8cd8b9a1ef89ca4b4399))


## v1.0.2 (2026-06-26)

### Bug Fixes

- **evaluator**: Track runtime deps so dynamic-ref shifts invalidate correctly
  ([#303](https://github.com/Teal-Insights/excel-grapher/pull/303),
  [`acf9301`](https://github.com/Teal-Insights/excel-grapher/commit/acf9301dde364352434998dcf803f8c8b5cb4814))

### Continuous Integration

- Use conventional commit for releases
  ([`cdc155a`](https://github.com/Teal-Insights/excel-grapher/commit/cdc155a3dc90e463a85cb0b6dcf1b92cc0a72708))

### Refactoring

- Add series-binding helper-block emitter and include_helpers flag
  ([#299](https://github.com/Teal-Insights/excel-grapher/pull/299),
  [`3c2a37d`](https://github.com/Teal-Insights/excel-grapher/commit/3c2a37d8f792a06be2af78de656d22173129b6ec))

- Emit series-binding coercion into a dedicated _api_helpers module
  ([#299](https://github.com/Teal-Insights/excel-grapher/pull/299),
  [`3c2a37d`](https://github.com/Teal-Insights/excel-grapher/commit/3c2a37d8f792a06be2af78de656d22173129b6ec))

### Testing

- Assert raw-emitted Ruff cleanliness instead of post-fix
  ([#299](https://github.com/Teal-Insights/excel-grapher/pull/299),
  [`3c2a37d`](https://github.com/Teal-Insights/excel-grapher/commit/3c2a37d8f792a06be2af78de656d22173129b6ec))


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
