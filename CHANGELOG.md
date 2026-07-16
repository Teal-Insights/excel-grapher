# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

<!-- version list -->

## v3.15.0 (2026-07-16)

### Bug Fixes

- **series-bindings**: Tighten reader export parity after review
  ([#413](https://github.com/Teal-Insights/excel-grapher/pull/413),
  [`a997795`](https://github.com/Teal-Insights/excel-grapher/commit/a997795b63aaf1c4301b821ecb5b665a2998a7bd))

### Code Style

- Ruff-format series bindings readers module test
  ([#413](https://github.com/Teal-Insights/excel-grapher/pull/413),
  [`a997795`](https://github.com/Teal-Insights/excel-grapher/commit/a997795b63aaf1c4301b821ecb5b665a2998a7bd))

### Features

- **series-bindings**: Migrate formula bodies onto read_* via _readers
  ([#413](https://github.com/Teal-Insights/excel-grapher/pull/413),
  [`a997795`](https://github.com/Teal-Insights/excel-grapher/commit/a997795b63aaf1c4301b821ecb5b665a2998a7bd))

### Testing

- **exporter**: Stop requiring unused xl_eval in projected internals
  ([#413](https://github.com/Teal-Insights/excel-grapher/pull/413),
  [`a997795`](https://github.com/Teal-Insights/excel-grapher/commit/a997795b63aaf1c4301b821ecb5b665a2998a7bd))


## v3.14.1 (2026-07-15)

### Bug Fixes

- Canonicalize multi-cell edge endpoints and cache schema
  ([`367b510`](https://github.com/Teal-Insights/excel-grapher/commit/367b51082b43fab5bb8aa749573feab428cac89a))

### Code Style

- Ruff format node extent constructor
  ([`367b510`](https://github.com/Teal-Insights/excel-grapher/commit/367b51082b43fab5bb8aa749573feab428cac89a))

### Refactoring

- Drop unused UnionKey import after row-shim removal
  ([`367b510`](https://github.com/Teal-Insights/excel-grapher/commit/367b51082b43fab5bb8aa749573feab428cac89a))

### Testing

- Migrate row-node tests to union nodes
  ([`367b510`](https://github.com/Teal-Insights/excel-grapher/commit/367b51082b43fab5bb8aa749573feab428cac89a))


## v3.14.0 (2026-07-15)

### Bug Fixes

- **series-bindings**: Treat overlapping data_ranges as ambiguous
  ([#412](https://github.com/Teal-Insights/excel-grapher/pull/412),
  [`5333fbf`](https://github.com/Teal-Insights/excel-grapher/commit/5333fbf4a21d8c5fae7111bb6ac36b45fbb41efa))

### Features

- **series-bindings**: Reverse address map for reader call forms
  ([#412](https://github.com/Teal-Insights/excel-grapher/pull/412),
  [`5333fbf`](https://github.com/Teal-Insights/excel-grapher/commit/5333fbf4a21d8c5fae7111bb6ac36b45fbb41efa))

- **series-bindings**: Reverse address map for reader call forms (#409)
  ([#412](https://github.com/Teal-Insights/excel-grapher/pull/412),
  [`5333fbf`](https://github.com/Teal-Insights/excel-grapher/commit/5333fbf4a21d8c5fae7111bb6ac36b45fbb41efa))


## v3.13.1 (2026-07-15)

### Bug Fixes

- **grapher**: Fail closed on non-arithmetic OFFSET extent ops
  ([#411](https://github.com/Teal-Insights/excel-grapher/pull/411),
  [`9a8ab6e`](https://github.com/Teal-Insights/excel-grapher/commit/9a8ab6ef172138008de5a579770880529c753d95))

- **grapher**: OFFSET named ranges with COUNTA(...)+n no longer collapse to 1×1
  ([#411](https://github.com/Teal-Insights/excel-grapher/pull/411),
  [`9a8ab6e`](https://github.com/Teal-Insights/excel-grapher/commit/9a8ab6ef172138008de5a579770880529c753d95))

- **grapher**: Resolve OFFSET named ranges with arithmetic extents
  ([#411](https://github.com/Teal-Insights/excel-grapher/pull/411),
  [`9a8ab6e`](https://github.com/Teal-Insights/excel-grapher/commit/9a8ab6ef172138008de5a579770880529c753d95))


## v3.13.0 (2026-07-15)

### Bug Fixes

- **series-bindings**: Align read_* emission with discovery exports
  ([#407](https://github.com/Teal-Insights/excel-grapher/pull/407),
  [`17de304`](https://github.com/Teal-Insights/excel-grapher/commit/17de304d09ffc6be451c0fa20bff416f52d7c52b))

- **test**: Avoid ruff B009 in reader range assertion
  ([#407](https://github.com/Teal-Insights/excel-grapher/pull/407),
  [`17de304`](https://github.com/Teal-Insights/excel-grapher/commit/17de304d09ffc6be451c0fa20bff416f52d7c52b))

- **test**: Silence ty on xl_range.cell access in reader tests
  ([#407](https://github.com/Teal-Insights/excel-grapher/pull/407),
  [`17de304`](https://github.com/Teal-Insights/excel-grapher/commit/17de304d09ffc6be451c0fa20bff416f52d7c52b))

### Features

- **series-bindings**: Emit read_* duals of set_* setters
  ([#407](https://github.com/Teal-Insights/excel-grapher/pull/407),
  [`17de304`](https://github.com/Teal-Insights/excel-grapher/commit/17de304d09ffc6be451c0fa20bff416f52d7c52b))

- **series-bindings**: Emit read_* duals of set_* setters (#404)
  ([#407](https://github.com/Teal-Insights/excel-grapher/pull/407),
  [`17de304`](https://github.com/Teal-Insights/excel-grapher/commit/17de304d09ffc6be451c0fa20bff416f52d7c52b))

### Testing

- **exporter**: Bump dep-tracking baseline for list_readers discovery
  ([#407](https://github.com/Teal-Insights/excel-grapher/pull/407),
  [`17de304`](https://github.com/Teal-Insights/excel-grapher/commit/17de304d09ffc6be451c0fa20bff416f52d7c52b))


## v3.12.0 (2026-07-15)

### Bug Fixes

- Type optional NumPy import for ty check
  ([#408](https://github.com/Teal-Insights/excel-grapher/pull/408),
  [`ffafd94`](https://github.com/Teal-Insights/excel-grapher/commit/ffafd94c32cb21e0ac8f65f4f4ed0f8c2b7cd6fa))

### Documentation

- **export**: Clarify INDEX/OFFSET ref vs value contract
  ([#406](https://github.com/Teal-Insights/excel-grapher/pull/406),
  [`b099df7`](https://github.com/Teal-Insights/excel-grapher/commit/b099df77d3e203a7e3bcab9576d9550179a76dde))

### Features

- Make NumPy an optional fast extra
  ([#408](https://github.com/Teal-Insights/excel-grapher/pull/408),
  [`ffafd94`](https://github.com/Teal-Insights/excel-grapher/commit/ffafd94c32cb21e0ac8f65f4f4ed0f8c2b7cd6fa))

- Make NumPy an optional fast extra (#403)
  ([#408](https://github.com/Teal-Insights/excel-grapher/pull/408),
  [`ffafd94`](https://github.com/Teal-Insights/excel-grapher/commit/ffafd94c32cb21e0ac8f65f4f4ed0f8c2b7cd6fa))

### Testing

- Keep NumPy-free CI green without re-syncing ops
  ([#408](https://github.com/Teal-Insights/excel-grapher/pull/408),
  [`ffafd94`](https://github.com/Teal-Insights/excel-grapher/commit/ffafd94c32cb21e0ac8f65f4f4ed0f8c2b7cd6fa))


## v3.11.0 (2026-07-15)

### Bug Fixes

- **test**: Satisfy ty check for compute_all calls in unpack tests
  ([#402](https://github.com/Teal-Insights/excel-grapher/pull/402),
  [`2e53f32`](https://github.com/Teal-Insights/excel-grapher/commit/2e53f32abcf59b2c92d6cb3d47af88fe494855ad))

### Features

- **exporter**: Optional return-line unpacking in codegen
  ([#402](https://github.com/Teal-Insights/excel-grapher/pull/402),
  [`2e53f32`](https://github.com/Teal-Insights/excel-grapher/commit/2e53f32abcf59b2c92d6cb3d47af88fe494855ad))

### Refactoring

- **evaluator**: Lazy range cleanup (#336 Phase 4)
  ([#401](https://github.com/Teal-Insights/excel-grapher/pull/401),
  [`cd88838`](https://github.com/Teal-Insights/excel-grapher/commit/cd88838a8cccaf6c6a9237b40405f93a7749f651))

- **exporter**: Hoist return temps during formula AST emission
  ([#402](https://github.com/Teal-Insights/excel-grapher/pull/402),
  [`2e53f32`](https://github.com/Teal-Insights/excel-grapher/commit/2e53f32abcf59b2c92d6cb3d47af88fe494855ad))

### Testing

- **exporter**: Broaden unpack_return coverage and docs
  ([#402](https://github.com/Teal-Insights/excel-grapher/pull/402),
  [`2e53f32`](https://github.com/Teal-Insights/excel-grapher/commit/2e53f32abcf59b2c92d6cb3d47af88fe494855ad))


## v3.10.1 (2026-07-15)

### Performance Improvements

- **runtime**: Stream AVERAGEIF via Grid.at_flat pairing
  ([#399](https://github.com/Teal-Insights/excel-grapher/pull/399),
  [`36d0ce6`](https://github.com/Teal-Insights/excel-grapher/commit/36d0ce6823f605fbba6a1754ad967c912fb3419e))


## v3.10.0 (2026-07-15)

### Features

- **evaluator**: Cell-wise AND/OR over lazy Range (#397)
  ([#400](https://github.com/Teal-Insights/excel-grapher/pull/400),
  [`74e4252`](https://github.com/Teal-Insights/excel-grapher/commit/74e425211d8dd8b2eceac32cc1ef14993808397c))


## v3.9.0 (2026-07-15)

### Bug Fixes

- **core**: Drop redundant cast in Range flatten walk
  ([#396](https://github.com/Teal-Insights/excel-grapher/pull/396),
  [`c07b157`](https://github.com/Teal-Insights/excel-grapher/commit/c07b157297592b8f6dccd54f8f95229ee285c29b))

- **evaluator**: Excel COUNTIF skip, AST precheck exemptions, demote AND/OR
  ([#396](https://github.com/Teal-Insights/excel-grapher/pull/396),
  [`c07b157`](https://github.com/Teal-Insights/excel-grapher/commit/c07b157297592b8f6dccd54f8f95229ee285c29b))

### Features

- **evaluator**: Lazy Grid aggregates for SUM/SUMPRODUCT (#336 Phase 3)
  ([#396](https://github.com/Teal-Insights/excel-grapher/pull/396),
  [`c07b157`](https://github.com/Teal-Insights/excel-grapher/commit/c07b157297592b8f6dccd54f8f95229ee285c29b))


## v3.8.0 (2026-07-15)

### Bug Fixes

- **evaluator**: Reuse materialized arrays on operator fastpath miss
  ([#394](https://github.com/Teal-Insights/excel-grapher/pull/394),
  [`e7a5c0f`](https://github.com/Teal-Insights/excel-grapher/commit/e7a5c0fcacc2e65f79c669c56ee6a0c85fb1c7f9))

- **evaluator**: Type/embed polish for Phase 2 operator maps
  ([#394](https://github.com/Teal-Insights/excel-grapher/pull/394),
  [`e7a5c0f`](https://github.com/Teal-Insights/excel-grapher/commit/e7a5c0fcacc2e65f79c669c56ee6a0c85fb1c7f9))

### Features

- **evaluator**: Lazy Grid maps for binary/unary operators (#336 Phase 2)
  ([#394](https://github.com/Teal-Insights/excel-grapher/pull/394),
  [`e7a5c0f`](https://github.com/Teal-Insights/excel-grapher/commit/e7a5c0fcacc2e65f79c669c56ee6a0c85fb1c7f9))

- **evaluator**: Route binary ops through shared Grid maps (#336 Phase 2)
  ([#394](https://github.com/Teal-Insights/excel-grapher/pull/394),
  [`e7a5c0f`](https://github.com/Teal-Insights/excel-grapher/commit/e7a5c0fcacc2e65f79c669c56ee6a0c85fb1c7f9))


## v3.7.0 (2026-07-14)

### Bug Fixes

- **series-bindings**: Accept numpy scalars in measure dtype checks
  ([#386](https://github.com/Teal-Insights/excel-grapher/pull/386),
  [`c42017c`](https://github.com/Teal-Insights/excel-grapher/commit/c42017c47dbe3a4de0921f9be75cc36b5d9a71af))

- **series-bindings**: Harden measure dtype enforcement and cleanup
  ([#386](https://github.com/Teal-Insights/excel-grapher/pull/386),
  [`c42017c`](https://github.com/Teal-Insights/excel-grapher/commit/c42017c47dbe3a4de0921f9be75cc36b5d9a71af))

### Features

- **series-bindings**: Enforce measure dtype in generated setters
  ([#386](https://github.com/Teal-Insights/excel-grapher/pull/386),
  [`c42017c`](https://github.com/Teal-Insights/excel-grapher/commit/c42017c47dbe3a4de0921f9be75cc36b5d9a71af))

- **series-bindings**: Narrow setter input annotations by measure dtype
  ([#386](https://github.com/Teal-Insights/excel-grapher/pull/386),
  [`c42017c`](https://github.com/Teal-Insights/excel-grapher/commit/c42017c47dbe3a4de0921f9be75cc36b5d9a71af))


## v3.6.0 (2026-07-14)

### Bug Fixes

- **evaluator**: Scalar boundary for lazy Range (#336 Phase 1)
  ([#389](https://github.com/Teal-Insights/excel-grapher/pull/389),
  [`0d63e45`](https://github.com/Teal-Insights/excel-grapher/commit/0d63e450d85ffc27c438f0afc8761b6ea6d1a9b9))

### Features

- **evaluator**: Lazy-by-default range resolution (#336 Phase 1)
  ([#389](https://github.com/Teal-Insights/excel-grapher/pull/389),
  [`0d63e45`](https://github.com/Teal-Insights/excel-grapher/commit/0d63e450d85ffc27c438f0afc8761b6ea6d1a9b9))

### Refactoring

- **evaluator**: Explicit eager/grid/VALUE range arg policy
  ([#389](https://github.com/Teal-Insights/excel-grapher/pull/389),
  [`0d63e45`](https://github.com/Teal-Insights/excel-grapher/commit/0d63e450d85ffc27c438f0afc8761b6ea6d1a9b9))


## v3.5.0 (2026-07-14)

### Bug Fixes

- **evaluator**: Cast lazy Range and lookup returns for ty
  ([#388](https://github.com/Teal-Insights/excel-grapher/pull/388),
  [`c9925a0`](https://github.com/Teal-Insights/excel-grapher/commit/c9925a03324c8a46c6482315e3891e53f8c3c354))

### Features

- **evaluator**: Lazy Range for lookup consumers
  ([#388](https://github.com/Teal-Insights/excel-grapher/pull/388),
  [`c9925a0`](https://github.com/Teal-Insights/excel-grapher/commit/c9925a03324c8a46c6482315e3891e53f8c3c354))

- **evaluator**: Lazy Range for lookup consumers (#336)
  ([#388](https://github.com/Teal-Insights/excel-grapher/pull/388),
  [`c9925a0`](https://github.com/Teal-Insights/excel-grapher/commit/c9925a03324c8a46c6482315e3891e53f8c3c354))

### Refactoring

- **core**: Unify ExcelRange as shared geometry type
  ([#388](https://github.com/Teal-Insights/excel-grapher/pull/388),
  [`c9925a0`](https://github.com/Teal-Insights/excel-grapher/commit/c9925a03324c8a46c6482315e3891e53f8c3c354))

### Testing

- **evaluator**: Budget and selective-access coverage for lazy lookups
  ([#388](https://github.com/Teal-Insights/excel-grapher/pull/388),
  [`c9925a0`](https://github.com/Teal-Insights/excel-grapher/commit/c9925a03324c8a46c6482315e3891e53f8c3c354))


## v3.4.3 (2026-07-14)

### Bug Fixes

- **addressing**: Accept ExcelRangeGeometry protocol for export embed
  ([#383](https://github.com/Teal-Insights/excel-grapher/pull/383),
  [`454c25a`](https://github.com/Teal-Insights/excel-grapher/commit/454c25a3a6abc5d6a2ca32e7743c2b97386e7aa4))

- **export_runtime**: Avoid CoreCellValue alias in embedded resolver
  ([#383](https://github.com/Teal-Insights/excel-grapher/pull/383),
  [`454c25a`](https://github.com/Teal-Insights/excel-grapher/commit/454c25a3a6abc5d6a2ca32e7743c2b97386e7aa4))

### Refactoring

- Remove dead type/lint suppressions and narrow export ExcelRange bridging
  ([#383](https://github.com/Teal-Insights/excel-grapher/pull/383),
  [`454c25a`](https://github.com/Teal-Insights/excel-grapher/commit/454c25a3a6abc5d6a2ca32e7743c2b97386e7aa4))

- Typing audit — drop dead suppressions, narrow ExcelRange casts
  ([#383](https://github.com/Teal-Insights/excel-grapher/pull/383),
  [`454c25a`](https://github.com/Teal-Insights/excel-grapher/commit/454c25a3a6abc5d6a2ca32e7743c2b97386e7aa4))


## v3.4.2 (2026-07-13)

### Bug Fixes

- **core**: Unify same-sheet range normalization on single sheet prefix
  ([#382](https://github.com/Teal-Insights/excel-grapher/pull/382),
  [`6442ade`](https://github.com/Teal-Insights/excel-grapher/commit/6442adef6561a9300da22325607c11c2d920edae))

- **core**: Unify same-sheet ranges on single sheet prefix
  ([#382](https://github.com/Teal-Insights/excel-grapher/pull/382),
  [`6442ade`](https://github.com/Teal-Insights/excel-grapher/commit/6442adef6561a9300da22325607c11c2d920edae))

- **grapher**: Harden single-prefix range dep extraction
  ([#382](https://github.com/Teal-Insights/excel-grapher/pull/382),
  [`6442ade`](https://github.com/Teal-Insights/excel-grapher/commit/6442adef6561a9300da22325607c11c2d920edae))

- **grapher**: Mask range spans before cell-ref parse
  ([#382](https://github.com/Teal-Insights/excel-grapher/pull/382),
  [`6442ade`](https://github.com/Teal-Insights/excel-grapher/commit/6442adef6561a9300da22325607c11c2d920edae))

- **grapher**: Refuse unmasked ranges in parse_cell_refs
  ([#382](https://github.com/Teal-Insights/excel-grapher/pull/382),
  [`6442ade`](https://github.com/Teal-Insights/excel-grapher/commit/6442adef6561a9300da22325607c11c2d920edae))

- **test**: Silence ty invalid-argument for CodeGenerator(None)
  ([#382](https://github.com/Teal-Insights/excel-grapher/pull/382),
  [`6442ade`](https://github.com/Teal-Insights/excel-grapher/commit/6442adef6561a9300da22325607c11c2d920edae))

### Code Style

- Ruff format and import tidy for range single-prefix
  ([#382](https://github.com/Teal-Insights/excel-grapher/pull/382),
  [`6442ade`](https://github.com/Teal-Insights/excel-grapher/commit/6442adef6561a9300da22325607c11c2d920edae))

- Ruff format parse_cell_refs signature
  ([#382](https://github.com/Teal-Insights/excel-grapher/pull/382),
  [`6442ade`](https://github.com/Teal-Insights/excel-grapher/commit/6442adef6561a9300da22325607c11c2d920edae))

### Refactoring

- **core**: Share colon split and canonicalize range ends
  ([#382](https://github.com/Teal-Insights/excel-grapher/pull/382),
  [`6442ade`](https://github.com/Teal-Insights/excel-grapher/commit/6442adef6561a9300da22325607c11c2d920edae))


## v3.4.1 (2026-07-13)

### Bug Fixes

- **evaluator**: Re-emit circular-reference warning on memoized re-evaluate
  ([#381](https://github.com/Teal-Insights/excel-grapher/pull/381),
  [`b14b682`](https://github.com/Teal-Insights/excel-grapher/commit/b14b6820e46a0b97114ffa4b2800f789565e1962))

### Refactoring

- **parity**: Dedupe live parity onto workbook compare helper
  ([#380](https://github.com/Teal-Insights/excel-grapher/pull/380),
  [`42e7085`](https://github.com/Teal-Insights/excel-grapher/commit/42e70856d278f7b156229d5680c27a15ffb69622))

### Testing

- **exporter**: Refresh dep-tracking baseline after circular-warning runtime
  ([#381](https://github.com/Teal-Insights/excel-grapher/pull/381),
  [`b14b682`](https://github.com/Teal-Insights/excel-grapher/commit/b14b6820e46a0b97114ffa4b2800f789565e1962))

- **parity**: Assert Excel error codes in excel_workbook_parity
  ([#380](https://github.com/Teal-Insights/excel-grapher/pull/380),
  [`42e7085`](https://github.com/Teal-Insights/excel-grapher/commit/42e70856d278f7b156229d5680c27a15ffb69622))


## v3.4.0 (2026-07-11)

### Features

- **series_bindings**: Per-dimension dtype for same-concept dimensions
  ([#378](https://github.com/Teal-Insights/excel-grapher/pull/378),
  [`829ac70`](https://github.com/Teal-Insights/excel-grapher/commit/829ac70e8146a372541e72450211e84b6560e125))

- **series_bindings**: Separate dimension id from concept (schema 1.8.0)
  ([#378](https://github.com/Teal-Insights/excel-grapher/pull/378),
  [`829ac70`](https://github.com/Teal-Insights/excel-grapher/commit/829ac70e8146a372541e72450211e84b6560e125))


## v3.3.0 (2026-07-09)

### Features

- **bindings**: Support internal series declarations for formula-cell key triangulation
  ([#373](https://github.com/Teal-Insights/excel-grapher/pull/373),
  [`a8d1ae3`](https://github.com/Teal-Insights/excel-grapher/commit/a8d1ae309e4930b71f2d8bffcff044515bd658d7))

### Refactoring

- **series_bindings**: Dedupe derive helpers and document internal series
  ([#373](https://github.com/Teal-Insights/excel-grapher/pull/373),
  [`a8d1ae3`](https://github.com/Teal-Insights/excel-grapher/commit/a8d1ae309e4930b71f2d8bffcff044515bd658d7))


## v3.2.0 (2026-07-08)

### Documentation

- **series_bindings**: Address PR 371 review feedback for input.mode override
  ([#371](https://github.com/Teal-Insights/excel-grapher/pull/371),
  [`db55ce6`](https://github.com/Teal-Insights/excel-grapher/commit/db55ce6907d932e88f742568c47de2cb0713d78e))

### Features

- **series_bindings**: Add input.mode override for formula cell setters
  ([#371](https://github.com/Teal-Insights/excel-grapher/pull/371),
  [`db55ce6`](https://github.com/Teal-Insights/excel-grapher/commit/db55ce6907d932e88f742568c47de2cb0713d78e))

- **series_bindings**: Input.mode override for formula cell setters
  ([#371](https://github.com/Teal-Insights/excel-grapher/pull/371),
  [`db55ce6`](https://github.com/Teal-Insights/excel-grapher/commit/db55ce6907d932e88f742568c47de2cb0713d78e))

### Refactoring

- **export**: Address PR review nits for three-layer wrappers
  ([#370](https://github.com/Teal-Insights/excel-grapher/pull/370),
  [`74b447d`](https://github.com/Teal-Insights/excel-grapher/commit/74b447de01820bd7c463f2aa560c0ae1b0c25907))

- **export**: Move worksheet functions to core with thin wrappers
  ([#370](https://github.com/Teal-Insights/excel-grapher/pull/370),
  [`74b447d`](https://github.com/Teal-Insights/excel-grapher/commit/74b447de01820bd7c463f2aa560c0ae1b0c25907))


## Unreleased

### Features

- **series_bindings**: Add `input.mode: override` (schema 1.6.0) for public setters on user-editable formula cells ([#371](https://github.com/Teal-Insights/excel-grapher/pull/371))

### Changed

- **series_bindings**: Leaf-mode input bindings now **error** on non-leaf `data_range` overlap (`non_leaf_input_overlap`) instead of warning and silently dropping formula cells. Manifests that relied on the old warn-and-drop behavior must either narrow `data_range` to graph leaves or declare `input.mode: override`.

## v3.1.0 (2026-07-07)

### Features

- **export**: Raise-only boundary for embedded runtime helpers (#326)
  ([#368](https://github.com/Teal-Insights/excel-grapher/pull/368),
  [`c700b80`](https://github.com/Teal-Insights/excel-grapher/commit/c700b805d79327c9f9171d19af74c50fbae42617))

- **export**: Wrap runtime calls with raise_if_sentinel at codegen boundary
  ([#368](https://github.com/Teal-Insights/excel-grapher/pull/368),
  [`c700b80`](https://github.com/Teal-Insights/excel-grapher/commit/c700b805d79327c9f9171d19af74c50fbae42617))

### Refactoring

- **export**: Use runtime boundary wrappers instead of codegen wrap
  ([#368](https://github.com/Teal-Insights/excel-grapher/pull/368),
  [`c700b80`](https://github.com/Teal-Insights/excel-grapher/commit/c700b805d79327c9f9171d19af74c50fbae42617))

### Testing

- **export**: Document shadowing invariants and expand boundary coverage
  ([#368](https://github.com/Teal-Insights/excel-grapher/pull/368),
  [`c700b80`](https://github.com/Teal-Insights/excel-grapher/commit/c700b805d79327c9f9171d19af74c50fbae42617))


## v3.0.1 (2026-07-06)

### Performance Improvements

- Speed up optimal projection manifests
  ([#350](https://github.com/Teal-Insights/excel-grapher/pull/350),
  [`9133cae`](https://github.com/Teal-Insights/excel-grapher/commit/9133caef3a687bf2da4476469c2c26332db75615))

### Testing

- Cover projection manifest ordering guards
  ([#350](https://github.com/Teal-Insights/excel-grapher/pull/350),
  [`9133cae`](https://github.com/Teal-Insights/excel-grapher/commit/9133caef3a687bf2da4476469c2c26332db75615))

- Guard projection metadata copy fields
  ([#350](https://github.com/Teal-Insights/excel-grapher/pull/350),
  [`9133cae`](https://github.com/Teal-Insights/excel-grapher/commit/9133caef3a687bf2da4476469c2c26332db75615))

- Guard shared preparsed ast projection copy
  ([#350](https://github.com/Teal-Insights/excel-grapher/pull/350),
  [`9133cae`](https://github.com/Teal-Insights/excel-grapher/commit/9133caef3a687bf2da4476469c2c26332db75615))


## v3.0.0 (2026-07-06)

### Refactoring

- **series_bindings**: Drop planned metadata and consolidate types
  ([#349](https://github.com/Teal-Insights/excel-grapher/pull/349),
  [`d9ba8fd`](https://github.com/Teal-Insights/excel-grapher/commit/d9ba8fd3d5d2acf917b141c4f7244b1117599d15))


## v2.5.2 (2026-07-06)

### Bug Fixes

- **grapher**: Repair local force subgraph selection
  ([#345](https://github.com/Teal-Insights/excel-grapher/pull/345),
  [`5d824bf`](https://github.com/Teal-Insights/excel-grapher/commit/5d824bf26d65f935748eba6d9cb57e5b7dcf86ab))

- **grapher**: Repair local force subgraph selection and add tests
  ([#345](https://github.com/Teal-Insights/excel-grapher/pull/345),
  [`5d824bf`](https://github.com/Teal-Insights/excel-grapher/commit/5d824bf26d65f935748eba6d9cb57e5b7dcf86ab))

### Code Style

- Fix import ordering in local force subgraph tests
  ([#345](https://github.com/Teal-Insights/excel-grapher/pull/345),
  [`5d824bf`](https://github.com/Teal-Insights/excel-grapher/commit/5d824bf26d65f935748eba6d9cb57e5b7dcf86ab))

### Refactoring

- Minor grapher cleanup (dead code, sha256, version, viz API)
  ([#348](https://github.com/Teal-Insights/excel-grapher/pull/348),
  [`208a56f`](https://github.com/Teal-Insights/excel-grapher/commit/208a56f7aad90d16207498423a2544299f57e23d))

- **test**: Move local force subgraph oracle into test helpers
  ([#345](https://github.com/Teal-Insights/excel-grapher/pull/345),
  [`5d824bf`](https://github.com/Teal-Insights/excel-grapher/commit/5d824bf26d65f935748eba6d9cb57e5b7dcf86ab))

### Testing

- **grapher**: Replace local force oracle with regression asserts
  ([#345](https://github.com/Teal-Insights/excel-grapher/pull/345),
  [`5d824bf`](https://github.com/Teal-Insights/excel-grapher/commit/5d824bf26d65f935748eba6d9cb57e5b7dcf86ab))


## v2.5.1 (2026-07-06)

### Bug Fixes

- Preserve quoted apostrophe sheet refs in formula normalization
  ([#346](https://github.com/Teal-Insights/excel-grapher/pull/346),
  [`d4d05a2`](https://github.com/Teal-Insights/excel-grapher/commit/d4d05a2b50c2a5b3f9bb5395abb0cf2da4d81b65))

- **evaluator**: Import xl_isblank from runtime after shim removal
  ([#347](https://github.com/Teal-Insights/excel-grapher/pull/347),
  [`37677be`](https://github.com/Teal-Insights/excel-grapher/commit/37677be0734d2af64b94b117612a72e0b9863a49))

### Refactoring

- Consolidate sheet-qualified address parsing
  ([#343](https://github.com/Teal-Insights/excel-grapher/pull/343),
  [`4ce69d7`](https://github.com/Teal-Insights/excel-grapher/commit/4ce69d72d87ba423929c4f8fb348c345334f11bd))

- Remove unused LocalForceSubgraph API
  ([#342](https://github.com/Teal-Insights/excel-grapher/pull/342),
  [`d3f261e`](https://github.com/Teal-Insights/excel-grapher/commit/d3f261e8d78ca6bea1f23c560edae7eb58b6140f))

- **evaluator**: Collapse function shims into explicit registry
  ([#347](https://github.com/Teal-Insights/excel-grapher/pull/347),
  [`37677be`](https://github.com/Teal-Insights/excel-grapher/commit/37677be0734d2af64b94b117612a72e0b9863a49))

### Testing

- Cover remaining apostrophe address parsers and fix call sites
  ([#343](https://github.com/Teal-Insights/excel-grapher/pull/343),
  [`4ce69d7`](https://github.com/Teal-Insights/excel-grapher/commit/4ce69d72d87ba423929c4f8fb348c345334f11bd))

- Xfail graph build with quoted apostrophe sheet refs in formulas
  ([#343](https://github.com/Teal-Insights/excel-grapher/pull/343),
  [`4ce69d7`](https://github.com/Teal-Insights/excel-grapher/commit/4ce69d72d87ba423929c4f8fb348c345334f11bd))


## v2.5.0 (2026-07-05)

### Documentation

- Add Cursor Cloud setup instructions to AGENTS.md
  ([#340](https://github.com/Teal-Insights/excel-grapher/pull/340),
  [`47b078a`](https://github.com/Teal-Insights/excel-grapher/commit/47b078aa48bc6f837651e409f44d7f86a09ba026))

### Features

- **grapher**: Opt-in AST pre-parsing during graph extraction
  ([#341](https://github.com/Teal-Insights/excel-grapher/pull/341),
  [`511492a`](https://github.com/Teal-Insights/excel-grapher/commit/511492ac5b77b797fc9053fab74c2fe25440d6a5))

### Testing

- Address PR 341 review feedback on preparsed formulas
  ([#341](https://github.com/Teal-Insights/excel-grapher/pull/341),
  [`511492a`](https://github.com/Teal-Insights/excel-grapher/commit/511492ac5b77b797fc9053fab74c2fe25440d6a5))


## v2.4.1 (2026-07-04)

### Performance Improvements

- **operators**: Fast batch coercion for numeric-string compare arrays
  ([#339](https://github.com/Teal-Insights/excel-grapher/pull/339),
  [`28e532c`](https://github.com/Teal-Insights/excel-grapher/commit/28e532c1f2c0fb9c8a42d9a556c68cb4df72b041))

### Refactoring

- **operators**: Address PR review feedback for numeric-string fastpath
  ([#339](https://github.com/Teal-Insights/excel-grapher/pull/339),
  [`28e532c`](https://github.com/Teal-Insights/excel-grapher/commit/28e532c1f2c0fb9c8a42d9a556c68cb4df72b041))

### Testing

- **exporter**: Refresh dep-tracking baseline after coercion helper
  ([#339](https://github.com/Teal-Insights/excel-grapher/pull/339),
  [`28e532c`](https://github.com/Teal-Insights/excel-grapher/commit/28e532c1f2c0fb9c8a42d9a556c68cb4df72b041))


## v2.4.0 (2026-07-04)

### Code Style

- **evaluator**: Fix ruff and docstring issues in AST cache PR
  ([#338](https://github.com/Teal-Insights/excel-grapher/pull/338),
  [`3724cfa`](https://github.com/Teal-Insights/excel-grapher/commit/3724cfa4f2ebddc037ef8e8f268fd4cbf31ab488))

### Features

- **evaluator**: Cache parsed formula ASTs keyed by normalized_formula
  ([#338](https://github.com/Teal-Insights/excel-grapher/pull/338),
  [`3724cfa`](https://github.com/Teal-Insights/excel-grapher/commit/3724cfa4f2ebddc037ef8e8f268fd4cbf31ab488))


## v2.3.0 (2026-07-04)

### Code Style

- Sort imports in evaluator math functions
  ([#334](https://github.com/Teal-Insights/excel-grapher/pull/334),
  [`5103a4a`](https://github.com/Teal-Insights/excel-grapher/commit/5103a4a9ef2a00e2cfb24f61e2e393f722ef57a2))

### Features

- **evaluator**: Implement EXP function
  ([#334](https://github.com/Teal-Insights/excel-grapher/pull/334),
  [`5103a4a`](https://github.com/Teal-Insights/excel-grapher/commit/5103a4a9ef2a00e2cfb24f61e2e393f722ef57a2))

### Testing

- **exp**: Strengthen coverage and add dynamic-ref domain inference
  ([#334](https://github.com/Teal-Insights/excel-grapher/pull/334),
  [`5103a4a`](https://github.com/Teal-Insights/excel-grapher/commit/5103a4a9ef2a00e2cfb24f61e2e393f722ef57a2))

- **parity**: Add live Excel harness and ABS/EXP excel parity tests
  ([#334](https://github.com/Teal-Insights/excel-grapher/pull/334),
  [`5103a4a`](https://github.com/Teal-Insights/excel-grapher/commit/5103a4a9ef2a00e2cfb24f61e2e393f722ef57a2))


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
