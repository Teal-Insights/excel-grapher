# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

<!-- version list -->

## v14.2.5 (2026-09-06)

### Bug Fixes

- **grapher**: Sound INDEX shape cache and unify FormulaShape parse
  ([#728](https://github.com/Teal-Insights/excel-grapher/pull/728),
  [`2fa5a92`](https://github.com/Teal-Insights/excel-grapher/commit/2fa5a92c5b93e70729f0e8d01112a29d5b5272c2))


## v14.2.4 (2026-09-06)

### Performance Improvements

- **grapher**: Speed up LIC-DSF-scale extract (#716)
  ([#719](https://github.com/Teal-Insights/excel-grapher/pull/719),
  [`fd5379c`](https://github.com/Teal-Insights/excel-grapher/commit/fd5379c8037069dd5c3bfde0a2508f6a2cb39f11))


## v14.2.3 (2026-09-06)

### Performance Improvements

- **grapher**: Reuse env membership, named-range regex, and sheet bounds
  ([#718](https://github.com/Teal-Insights/excel-grapher/pull/718),
  [`524a622`](https://github.com/Teal-Insights/excel-grapher/commit/524a62240b8810320d34a2227fc1a503f103490d))


## v14.2.2 (2026-09-06)

### Performance Improvements

- **grapher**: Cache repeated provenance walks
  ([#717](https://github.com/Teal-Insights/excel-grapher/pull/717),
  [`7137692`](https://github.com/Teal-Insights/excel-grapher/commit/71376922238447e38e28e70479c4fc4cf08faa95))


## v14.2.1 (2026-09-06)

### Bug Fixes

- **export**: Lower sparse MATCH/INDEX next-non-blank idiom
  ([#714](https://github.com/Teal-Insights/excel-grapher/pull/714),
  [`bb5bcc2`](https://github.com/Teal-Insights/excel-grapher/commit/bb5bcc21c6662cefe3536c8193a3a87915cc06b0))


## v14.2.0 (2026-09-06)

### Features

- **export**: Emit multi-owner MATCH/INDEX windows positionally
  ([#711](https://github.com/Teal-Insights/excel-grapher/pull/711),
  [`2d529e8`](https://github.com/Teal-Insights/excel-grapher/commit/2d529e873fd803cb7e69d67c9c022a3d37a1ccea))

### Testing

- **export**: Add multi-owner INDEX/MATCH inverted-tree regression
  ([#711](https://github.com/Teal-Insights/excel-grapher/pull/711),
  [`2d529e8`](https://github.com/Teal-Insights/excel-grapher/commit/2d529e873fd803cb7e69d67c9c022a3d37a1ccea))


## v14.1.1 (2026-09-06)

### Bug Fixes

- **export**: Keep on-graph leaves in series formula catalog
  ([#709](https://github.com/Teal-Insights/excel-grapher/pull/709),
  [`01d17da`](https://github.com/Teal-Insights/excel-grapher/commit/01d17da20454056e32d8bbe5db15c17bc75ab50d))

### Testing

- **export**: Narrow series hole_at before reading kind
  ([#709](https://github.com/Teal-Insights/excel-grapher/pull/709),
  [`01d17da`](https://github.com/Teal-Insights/excel-grapher/commit/01d17da20454056e32d8bbe5db15c17bc75ab50d))


## v14.1.0 (2026-09-06)

### Features

- **grapher**: Add select_shortest_path_subgraph
  ([#706](https://github.com/Teal-Insights/excel-grapher/pull/706),
  [`7c542fe`](https://github.com/Teal-Insights/excel-grapher/commit/7c542feed053f23d1cab635ae1ec6b91426b5e67))

### Testing

- **grapher**: Use A1 keys that survive normalize_key
  ([#706](https://github.com/Teal-Insights/excel-grapher/pull/706),
  [`7c542fe`](https://github.com/Teal-Insights/excel-grapher/commit/7c542feed053f23d1cab635ae1ec6b91426b5e67))


## v14.0.1 (2026-09-06)

### Bug Fixes

- **export**: Honor blank_ranges CellRef in inverted-tree emit
  ([#704](https://github.com/Teal-Insights/excel-grapher/pull/704),
  [`f43823f`](https://github.com/Teal-Insights/excel-grapher/commit/f43823f8a10b6bad62c86f77bfd198337840dacb))

- **series_bindings**: Apply exclude_rows in validate (#594)
  ([#705](https://github.com/Teal-Insights/excel-grapher/pull/705),
  [`c2901d2`](https://github.com/Teal-Insights/excel-grapher/commit/c2901d2de6f34dc2ca1e8531daade16063b60b60))


## v14.0.0 (2026-09-05)

### Features

- **export**: Remove the ctx package exporter
  ([#702](https://github.com/Teal-Insights/excel-grapher/pull/702),
  [`24cb32e`](https://github.com/Teal-Insights/excel-grapher/commit/24cb32e86b53b38fed16f2fd636f4a8f4a0b8bbb))


## v13.7.1 (2026-09-05)

### Bug Fixes

- **export**: Honor blank_ranges in inverted-tree emit
  ([#701](https://github.com/Teal-Insights/excel-grapher/pull/701),
  [`94b3961`](https://github.com/Teal-Insights/excel-grapher/commit/94b39618e44659e839e7288414466774485efe14))


## v13.7.0 (2026-09-05)

### Bug Fixes

- **export**: Gather irregular inverted-tree year picks via literal index map
  ([#697](https://github.com/Teal-Insights/excel-grapher/pull/697),
  [`157e02f`](https://github.com/Teal-Insights/excel-grapher/commit/157e02f20bbfe13bf076238fbea9a09da464aa5a))

- **export**: Gather sparse inverted-tree year picks via literal index map
  ([#697](https://github.com/Teal-Insights/excel-grapher/pull/697),
  [`157e02f`](https://github.com/Teal-Insights/excel-grapher/commit/157e02f20bbfe13bf076238fbea9a09da464aa5a))

- **export**: Keep cached-value coerce typed as Any
  ([#698](https://github.com/Teal-Insights/excel-grapher/pull/698),
  [`9fac457`](https://github.com/Teal-Insights/excel-grapher/commit/9fac457877a340ec265ea70cfe02b0dda63bdcb9))

- **export**: Take full-catalog helpers at irregular gather call sites
  ([#697](https://github.com/Teal-Insights/excel-grapher/pull/697),
  [`157e02f`](https://github.com/Teal-Insights/excel-grapher/commit/157e02f20bbfe13bf076238fbea9a09da464aa5a))

### Features

- **export**: Retain holes in matrix formula series
  ([#698](https://github.com/Teal-Insights/excel-grapher/pull/698),
  [`9fac457`](https://github.com/Teal-Insights/excel-grapher/commit/9fac457877a340ec265ea70cfe02b0dda63bdcb9))


## v13.6.4 (2026-09-05)

### Bug Fixes

- **export**: Keep partial_graph_overlap series in inverted-tree emit
  ([#694](https://github.com/Teal-Insights/excel-grapher/pull/694),
  [`40398c6`](https://github.com/Teal-Insights/excel-grapher/commit/40398c6064a8180c0eac9c8ad43df6894fc88322))

- **export**: Narrow emit validation and keep scan working buffers
  ([#694](https://github.com/Teal-Insights/excel-grapher/pull/694),
  [`40398c6`](https://github.com/Teal-Insights/excel-grapher/commit/40398c6064a8180c0eac9c8ad43df6894fc88322))


## v13.6.3 (2026-09-04)

### Bug Fixes

- **export**: Narrow as_measure return type with dtype overloads
  ([#692](https://github.com/Teal-Insights/excel-grapher/pull/692),
  [`6658fa0`](https://github.com/Teal-Insights/excel-grapher/commit/6658fa09a884b2069ce5d4ad3a2abd97f9944d0c))


## v13.6.2 (2026-09-04)

### Bug Fixes

- **export**: Publish inverted-tree keyed meta via setattr
  ([#691](https://github.com/Teal-Insights/excel-grapher/pull/691),
  [`3b2c82c`](https://github.com/Teal-Insights/excel-grapher/commit/3b2c82c617896b430eb9cf51fd613e5240cab666))

- **export**: Widen numeric inverted-tree leaf annotations with | str
  ([#690](https://github.com/Teal-Insights/excel-grapher/pull/690),
  [`baa6bf7`](https://github.com/Teal-Insights/excel-grapher/commit/baa6bf7b32917f3e1bfe912208ce735da47fa140))


## v13.6.1 (2026-09-04)

### Bug Fixes

- **export**: Cut inverted-tree generate time on large catalogs
  ([#686](https://github.com/Teal-Insights/excel-grapher/pull/686),
  [`4efff1c`](https://github.com/Teal-Insights/excel-grapher/commit/4efff1c8023cbb686400ee6f2a515525371ec5d4))


## v13.6.0 (2026-09-04)

### Bug Fixes

- Lower absolute series refs as static catalog indexes
  ([#685](https://github.com/Teal-Insights/excel-grapher/pull/685),
  [`e10a17d`](https://github.com/Teal-Insights/excel-grapher/commit/e10a17d6571923cfa3bffde280b1b64a9226c96e))

### Features

- **cli**: Apply DynamicRefConfig constraints in bindings validate
  ([#684](https://github.com/Teal-Insights/excel-grapher/pull/684),
  [`8b1669e`](https://github.com/Teal-Insights/excel-grapher/commit/8b1669ea94da799bd2203da26726024a8c7a4c72))


## v13.5.0 (2026-09-04)

### Features

- **series_bindings**: Map scalar input needles via input.value_map
  ([#680](https://github.com/Teal-Insights/excel-grapher/pull/680),
  [`0212339`](https://github.com/Teal-Insights/excel-grapher/commit/02123398a3487f089c885071114fc540a9df03b9))


## v13.4.0 (2026-09-04)

### Features

- **export**: Publish inverted-tree key domains on compute_*
  ([#679](https://github.com/Teal-Insights/excel-grapher/pull/679),
  [`8c309d7`](https://github.com/Teal-Insights/excel-grapher/commit/8c309d72622e27504c58a8b6268abb3e0a03ed0b))


## v13.3.1 (2026-09-04)

### Bug Fixes

- **series_bindings**: Reject duplicate ids and A1 geometry in public names
  ([#678](https://github.com/Teal-Insights/excel-grapher/pull/678),
  [`865369f`](https://github.com/Teal-Insights/excel-grapher/commit/865369f028bc75406a658610c96c0a1e666841bc))

### Documentation

- Cite #667 and #668 from the inverted-tree fail-closed list
  ([#669](https://github.com/Teal-Insights/excel-grapher/pull/669),
  [`12abaeb`](https://github.com/Teal-Insights/excel-grapher/commit/12abaebe9a902e60fe57c9c0bbd4d44ce49777e8))

- Link #662 port follow-ups to #666 #667 #668
  ([#669](https://github.com/Teal-Insights/excel-grapher/pull/669),
  [`12abaeb`](https://github.com/Teal-Insights/excel-grapher/commit/12abaebe9a902e60fe57c9c0bbd4d44ce49777e8))

- Link #662 port follow-ups to #666–#668
  ([#669](https://github.com/Teal-Insights/excel-grapher/pull/669),
  [`12abaeb`](https://github.com/Teal-Insights/excel-grapher/commit/12abaebe9a902e60fe57c9c0bbd4d44ce49777e8))


## v13.3.0 (2026-09-04)

### Bug Fixes

- **export**: Drop zero row terms from INDIRECT catalog indexes
  ([#672](https://github.com/Teal-Insights/excel-grapher/pull/672),
  [`0a7b3c5`](https://github.com/Teal-Insights/excel-grapher/commit/0a7b3c584ffc308157bc397a39f68fcd8d5ad20a))

### Features

- **export**: Lower inverted-tree INDIRECT from graph edges
  ([#672](https://github.com/Teal-Insights/excel-grapher/pull/672),
  [`0a7b3c5`](https://github.com/Teal-Insights/excel-grapher/commit/0a7b3c584ffc308157bc397a39f68fcd8d5ad20a))


## v13.2.1 (2026-09-04)

### Bug Fixes

- **export**: Align inverted-tree data.py annotations with compute_* params
  ([#674](https://github.com/Teal-Insights/excel-grapher/pull/674),
  [`43448dc`](https://github.com/Teal-Insights/excel-grapher/commit/43448dccde859c121cd0776e45f27856ad71e36a))


## v13.2.0 (2026-09-04)

### Features

- **export**: Lower inverted-tree SUM/SUMPRODUCT and remaining ref shapes
  ([#671](https://github.com/Teal-Insights/excel-grapher/pull/671),
  [`14057d4`](https://github.com/Teal-Insights/excel-grapher/commit/14057d463649fec3d2860557729b1f4de43f8451))


## v13.1.0 (2026-09-04)

### Features

- **export**: Enforce input.domain on inverted-tree compute arguments
  ([#670](https://github.com/Teal-Insights/excel-grapher/pull/670),
  [`a746d4a`](https://github.com/Teal-Insights/excel-grapher/commit/a746d4a70397b531a1c960d191eb01357037f06d))


## v13.0.0 (2026-09-04)

### Features

- Import inverted-tree constants from data instead of compute kwargs
  ([#664](https://github.com/Teal-Insights/excel-grapher/pull/664),
  [`c2a745c`](https://github.com/Teal-Insights/excel-grapher/commit/c2a745cc5338fc9b8e9b213b501c9bfe132a2016))

- **export**: Import inverted-tree constants from data in compute bodies
  ([#664](https://github.com/Teal-Insights/excel-grapher/pull/664),
  [`c2a745c`](https://github.com/Teal-Insights/excel-grapher/commit/c2a745cc5338fc9b8e9b213b501c9bfe132a2016))

### Testing

- **export**: Type inverted-tree constant-count helper as ModuleType
  ([#664](https://github.com/Teal-Insights/excel-grapher/pull/664),
  [`c2a745c`](https://github.com/Teal-Insights/excel-grapher/commit/c2a745cc5338fc9b8e9b213b501c9bfe132a2016))


## v12.8.0 (2026-09-04)

### Documentation

- Inverted-tree as primary paradigm and ctx deprecation plan
  ([#665](https://github.com/Teal-Insights/excel-grapher/pull/665),
  [`32d4174`](https://github.com/Teal-Insights/excel-grapher/commit/32d4174d0c47abdb464ffb91aae142251a5ab5c3))

### Features

- **export**: Smoke inverted-tree computes and emit datetime leaves
  ([#665](https://github.com/Teal-Insights/excel-grapher/pull/665),
  [`32d4174`](https://github.com/Teal-Insights/excel-grapher/commit/32d4174d0c47abdb464ffb91aae142251a5ab5c3))

### Testing

- **export**: Probe ctx features against inverted-tree export
  ([#665](https://github.com/Teal-Insights/excel-grapher/pull/665),
  [`32d4174`](https://github.com/Teal-Insights/excel-grapher/commit/32d4174d0c47abdb464ffb91aae142251a5ab5c3))


## v12.7.2 (2026-09-04)

### Bug Fixes

- **export**: Index fused partition seeds with _area
  ([#661](https://github.com/Teal-Insights/excel-grapher/pull/661),
  [`959bd1a`](https://github.com/Teal-Insights/excel-grapher/commit/959bd1a1b1b21db91055cdfa6dac4e0f6b8527e2))

### Testing

- **export**: Replace inverted-tree generator wall-clock bounds with op counts
  ([#660](https://github.com/Teal-Insights/excel-grapher/pull/660),
  [`a83b75d`](https://github.com/Teal-Insights/excel-grapher/commit/a83b75d857c86bab16fe69353f1765a3505b1769))


## v12.7.1 (2026-09-04)

### Bug Fixes

- **export**: Classify block access per statement, not per series
  ([#659](https://github.com/Teal-Insights/excel-grapher/pull/659),
  [`ca0bd62`](https://github.com/Teal-Insights/excel-grapher/commit/ca0bd621adfb587474068f85b1945c19c159a32d))

- **export**: Lower INDEX into a block when the range overhangs it
  ([#659](https://github.com/Teal-Insights/excel-grapher/pull/659),
  [`ca0bd62`](https://github.com/Teal-Insights/excel-grapher/commit/ca0bd621adfb587474068f85b1945c19c159a32d))


## v12.7.0 (2026-09-04)

### Bug Fixes

- **core**: Compare with Excel type-rank, including the array fastpath
  ([#658](https://github.com/Teal-Insights/excel-grapher/pull/658),
  [`24a9196`](https://github.com/Teal-Insights/excel-grapher/commit/24a91964c4395d7a5c63c5e4aa4a47b0ed97edef))

- **core**: Share type-rank comparison with inverted-tree runtime
  ([#658](https://github.com/Teal-Insights/excel-grapher/pull/658),
  [`24a9196`](https://github.com/Teal-Insights/excel-grapher/commit/24a91964c4395d7a5c63c5e4aa4a47b0ed97edef))

- **export**: Classify INDEX/OFFSET access then emit via affine anchors
  ([#658](https://github.com/Teal-Insights/excel-grapher/pull/658),
  [`24a9196`](https://github.com/Teal-Insights/excel-grapher/commit/24a91964c4395d7a5c63c5e4aa4a47b0ed97edef))

- **export**: Fail closed only on several host±1 seeds
  ([#658](https://github.com/Teal-Insights/excel-grapher/pull/658),
  [`24a9196`](https://github.com/Teal-Insights/excel-grapher/commit/24a91964c4395d7a5c63c5e4aa4a47b0ed97edef))

- **export**: Type inverted-tree operator wrappers as FormulaValue
  ([#658](https://github.com/Teal-Insights/excel-grapher/pull/658),
  [`24a9196`](https://github.com/Teal-Insights/excel-grapher/commit/24a91964c4395d7a5c63c5e4aa4a47b0ed97edef))

### Features

- **export**: Add local inverted-tree workbook corpus harness
  ([#658](https://github.com/Teal-Insights/excel-grapher/pull/658),
  [`24a9196`](https://github.com/Teal-Insights/excel-grapher/commit/24a91964c4395d7a5c63c5e4aa4a47b0ed97edef))

- **export**: Derive inverted-tree INDEX/OFFSET access from graph edges
  ([#658](https://github.com/Teal-Insights/excel-grapher/pull/658),
  [`24a9196`](https://github.com/Teal-Insights/excel-grapher/commit/24a91964c4395d7a5c63c5e4aa4a47b0ed97edef))

### Testing

- **core**: Type-rank equality for numeric-string ndarray pairs
  ([#658](https://github.com/Teal-Insights/excel-grapher/pull/658),
  [`24a9196`](https://github.com/Teal-Insights/excel-grapher/commit/24a91964c4395d7a5c63c5e4aa4a47b0ed97edef))


## v12.6.2 (2026-09-04)

### Bug Fixes

- **export**: Lower INDEX/OFFSET into a 2-D block without dropping the column
  ([`88f1647`](https://github.com/Teal-Insights/excel-grapher/commit/88f1647ebf7bf17665db39a4531a2cc5c00d9d06))

- **export**: Tighten 2-D INDEX fixtures and emitted flat-index terms
  ([`53c61dd`](https://github.com/Teal-Insights/excel-grapher/commit/53c61dd8f078f218d2e87f4648799db4bffde4ac))

### Documentation

- **export**: Inverted-tree plan §11–§12 — graph-derived lowering and local corpus (#656)
  ([#657](https://github.com/Teal-Insights/excel-grapher/pull/657),
  [`6a186e6`](https://github.com/Teal-Insights/excel-grapher/commit/6a186e62b9312f31db069fb75edaff8bcb28e70b))

### Testing

- **export**: Cover INDEX/OFFSET into a 2-D bound block
  ([`9ad9b1e`](https://github.com/Teal-Insights/excel-grapher/commit/9ad9b1e8465e5510c44359b6671a2b17e5f78520))


## v12.6.1 (2026-09-04)

### Bug Fixes

- **export**: Do not treat matrix row copies as year-0 scan seeds
  ([`ffef793`](https://github.com/Teal-Insights/excel-grapher/commit/ffef79364b33474c4e4e3df122c062a0a175f3fe))

### Chores

- Retrigger CI after unrelated catalog timing flake
  ([#647](https://github.com/Teal-Insights/excel-grapher/pull/647),
  [`a42bfaa`](https://github.com/Teal-Insights/excel-grapher/commit/a42bfaa275273d39b8e11b3b7486c69ac2ac6fbd))

### Refactoring

- **export**: Inverted-tree rung symmetry and force_rung=2 fallthrough
  ([#647](https://github.com/Teal-Insights/excel-grapher/pull/647),
  [`a42bfaa`](https://github.com/Teal-Insights/excel-grapher/commit/a42bfaa275273d39b8e11b3b7486c69ac2ac6fbd))

- **export**: Symmetric inverted-tree rungs and force_rung=2 fallthrough
  ([#647](https://github.com/Teal-Insights/excel-grapher/pull/647),
  [`a42bfaa`](https://github.com/Teal-Insights/excel-grapher/commit/a42bfaa275273d39b8e11b3b7486c69ac2ac6fbd))

### Testing

- **export**: Absorb timer noise in catalog partition linearity check
  ([#647](https://github.com/Teal-Insights/excel-grapher/pull/647),
  [`a42bfaa`](https://github.com/Teal-Insights/excel-grapher/commit/a42bfaa275273d39b8e11b3b7486c69ac2ac6fbd))

- **export**: Retire inverted-tree _exec_scan and eval_instance greps
  ([#648](https://github.com/Teal-Insights/excel-grapher/pull/648),
  [`eb8b0de`](https://github.com/Teal-Insights/excel-grapher/commit/eb8b0de7ee530ec9c841c04c30a9b5c9b505c07f))

- **export**: Retire inverted-tree `_exec_scan` and `eval_instance` greps
  ([#648](https://github.com/Teal-Insights/excel-grapher/pull/648),
  [`eb8b0de`](https://github.com/Teal-Insights/excel-grapher/commit/eb8b0de7ee530ec9c841c04c30a9b5c9b505c07f))

- **export**: Stabilize partition_catalog linearity timing
  ([#648](https://github.com/Teal-Insights/excel-grapher/pull/648),
  [`eb8b0de`](https://github.com/Teal-Insights/excel-grapher/commit/eb8b0de7ee530ec9c841c04c30a9b5c9b505c07f))


## v12.6.0 (2026-09-04)

### Features

- **export**: Fuse matrix SCCs on per-partition TIME_PERIOD
  ([#646](https://github.com/Teal-Insights/excel-grapher/pull/646),
  [`3b71aac`](https://github.com/Teal-Insights/excel-grapher/commit/3b71aac5db08e1c12c7691f151ac83e01dde8de4))


## v12.5.4 (2026-09-04)

### Bug Fixes

- **export**: Coerce inverted-tree arithmetic and comparison operands
  ([#643](https://github.com/Teal-Insights/excel-grapher/pull/643),
  [`f8db842`](https://github.com/Teal-Insights/excel-grapher/commit/f8db842b1a45777ae96571cf633085becedd792a))

- **export**: Coerce inverted-tree operator operands (#635)
  ([#643](https://github.com/Teal-Insights/excel-grapher/pull/643),
  [`f8db842`](https://github.com/Teal-Insights/excel-grapher/commit/f8db842b1a45777ae96571cf633085becedd792a))

- **export**: Keep inverted-tree indexes integral after coercion
  ([#643](https://github.com/Teal-Insights/excel-grapher/pull/643),
  [`f8db842`](https://github.com/Teal-Insights/excel-grapher/commit/f8db842b1a45777ae96571cf633085becedd792a))

- **export**: Normalize inverted-tree catalog addresses once
  ([#645](https://github.com/Teal-Insights/excel-grapher/pull/645),
  [`effb353`](https://github.com/Teal-Insights/excel-grapher/commit/effb3537e73bf4c623553259fbcdc72b44175f20))


## v12.5.3 (2026-09-04)

### Bug Fixes

- **export**: Classify seed/terminal lags by schedule and relative refs
  ([#644](https://github.com/Teal-Insights/excel-grapher/pull/644),
  [`c747e8f`](https://github.com/Teal-Insights/excel-grapher/commit/c747e8f49bb75bea1f74dda1a300f548bee98441))


## v12.5.2 (2026-09-04)

### Bug Fixes

- **export**: O(1) statement lookup for fused-region planning
  ([#642](https://github.com/Teal-Insights/excel-grapher/pull/642),
  [`4285341`](https://github.com/Teal-Insights/excel-grapher/commit/4285341187a07f017e75caa49b2417ac0ff8c150))

- **export**: Skip statement-map rebuild when partition is unchanged
  ([#642](https://github.com/Teal-Insights/excel-grapher/pull/642),
  [`4285341`](https://github.com/Teal-Insights/excel-grapher/commit/4285341187a07f017e75caa49b2417ac0ff8c150))

### Testing

- **export**: Clarify N=5000 backward-chain docstring
  ([#642](https://github.com/Teal-Insights/excel-grapher/pull/642),
  [`4285341`](https://github.com/Teal-Insights/excel-grapher/commit/4285341187a07f017e75caa49b2417ac0ff8c150))


## v12.5.1 (2026-09-04)

### Bug Fixes

- **export**: Absolute IF selector is not a year-0 scan seed
  ([#632](https://github.com/Teal-Insights/excel-grapher/pull/632),
  [`61f354c`](https://github.com/Teal-Insights/excel-grapher/commit/61f354ce88753b728d9a4cf04349d30f27520442))

- **export**: Coerce string measures in inverted-tree xl_div
  ([#632](https://github.com/Teal-Insights/excel-grapher/pull/632),
  [`61f354c`](https://github.com/Teal-Insights/excel-grapher/commit/61f354ce88753b728d9a4cf04349d30f27520442))

- **export**: Do not treat absolute selectors as year-0 scan seeds
  ([#632](https://github.com/Teal-Insights/excel-grapher/pull/632),
  [`61f354c`](https://github.com/Teal-Insights/excel-grapher/commit/61f354ce88753b728d9a4cf04349d30f27520442))

- **export**: Index taken windows in rung-3 inverted-tree helpers
  ([#641](https://github.com/Teal-Insights/excel-grapher/pull/641),
  [`abadaec`](https://github.com/Teal-Insights/excel-grapher/commit/abadaeca1e511d12a5d92f9bda16723c1fe5e3a2))

### Testing

- **export**: Fix orientation rewrite and rung-3 corpus gaps
  ([#630](https://github.com/Teal-Insights/excel-grapher/pull/630),
  [`a07cfb3`](https://github.com/Teal-Insights/excel-grapher/commit/a07cfb3e9f2b8012ea6a829c5d95d700fd24cf41))

- **export**: Inverted-tree design-property suite (#621)
  ([#630](https://github.com/Teal-Insights/excel-grapher/pull/630),
  [`a07cfb3`](https://github.com/Teal-Insights/excel-grapher/commit/a07cfb3e9f2b8012ea6a829c5d95d700fd24cf41))

- **export**: Pin inverted-tree design properties
  ([#630](https://github.com/Teal-Insights/excel-grapher/pull/630),
  [`a07cfb3`](https://github.com/Teal-Insights/excel-grapher/commit/a07cfb3e9f2b8012ea6a829c5d95d700fd24cf41))


## v12.5.0 (2026-09-03)

### Bug Fixes

- **export**: Allow inverted-tree other-series t and t-1 reads
  ([#628](https://github.com/Teal-Insights/excel-grapher/pull/628),
  [`9d30182`](https://github.com/Teal-Insights/excel-grapher/commit/9d3018220085f390ce6f1627a843d3094e72dc24))

- **export**: Apply exclude_rows in inverted-tree catalog
  ([#628](https://github.com/Teal-Insights/excel-grapher/pull/628),
  [`9d30182`](https://github.com/Teal-Insights/excel-grapher/commit/9d3018220085f390ce6f1627a843d3094e72dc24))

- **export**: Cache inverted-tree schedule coordinates
  ([#628](https://github.com/Teal-Insights/excel-grapher/pull/628),
  [`9d30182`](https://github.com/Teal-Insights/excel-grapher/commit/9d3018220085f390ce6f1627a843d3094e72dc24))

- **export**: Define inverted-tree xl_exp for EXP
  ([#628](https://github.com/Teal-Insights/excel-grapher/pull/628),
  [`9d30182`](https://github.com/Teal-Insights/excel-grapher/commit/9d3018220085f390ce6f1627a843d3094e72dc24))

- **export**: Derive inverted-tree identity maps from the key-point join
  ([#628](https://github.com/Teal-Insights/excel-grapher/pull/628),
  [`9d30182`](https://github.com/Teal-Insights/excel-grapher/commit/9d3018220085f390ce6f1627a843d3094e72dc24))

- **export**: Drive rung-3 backward chains in reverse to avoid recursion
  ([#628](https://github.com/Teal-Insights/excel-grapher/pull/628),
  [`9d30182`](https://github.com/Teal-Insights/excel-grapher/commit/9d3018220085f390ce6f1627a843d3094e72dc24))

- **export**: Emit lag-zipper series as a co-scan
  ([#628](https://github.com/Teal-Insights/excel-grapher/pull/628),
  [`9d30182`](https://github.com/Teal-Insights/excel-grapher/commit/9d3018220085f390ce6f1627a843d3094e72dc24))

- **export**: Emit live_measure for fused refs off the SCC union (#623)
  ([#628](https://github.com/Teal-Insights/excel-grapher/pull/628),
  [`9d30182`](https://github.com/Teal-Insights/excel-grapher/commit/9d3018220085f390ce6f1627a843d3094e72dc24))

- **export**: Fail closed on partially resolved inverted-tree keys
  ([#628](https://github.com/Teal-Insights/excel-grapher/pull/628),
  [`9d30182`](https://github.com/Teal-Insights/excel-grapher/commit/9d3018220085f390ce6f1627a843d3094e72dc24))

- **export**: Gather inverted-tree series with take, not trim
  ([#628](https://github.com/Teal-Insights/excel-grapher/pull/628),
  [`9d30182`](https://github.com/Teal-Insights/excel-grapher/commit/9d3018220085f390ce6f1627a843d3094e72dc24))

- **export**: Index taken windows in 1-cell inverted-tree helpers
  ([#629](https://github.com/Teal-Insights/excel-grapher/pull/629),
  [`063f1b1`](https://github.com/Teal-Insights/excel-grapher/commit/063f1b11205d1db86023aada01c536eff0aa8e2a))

- **export**: Inverted-tree emit nits for errors, SCC call, and else
  ([#628](https://github.com/Teal-Insights/excel-grapher/pull/628),
  [`9d30182`](https://github.com/Teal-Insights/excel-grapher/commit/9d3018220085f390ce6f1627a843d3094e72dc24))

- **export**: Join inverted-tree matrix schedules on the full key tuple
  ([#628](https://github.com/Teal-Insights/excel-grapher/pull/628),
  [`9d30182`](https://github.com/Teal-Insights/excel-grapher/commit/9d3018220085f390ce6f1627a843d3094e72dc24))

- **export**: Join inverted-tree schedule on key-point domain
  ([#628](https://github.com/Teal-Insights/excel-grapher/pull/628),
  [`9d30182`](https://github.com/Teal-Insights/excel-grapher/commit/9d3018220085f390ce6f1627a843d3094e72dc24))

- **export**: Keep IMF sentinels in float inverted-tree constants
  ([#628](https://github.com/Teal-Insights/excel-grapher/pull/628),
  [`9d30182`](https://github.com/Teal-Insights/excel-grapher/commit/9d3018220085f390ce6f1627a843d3094e72dc24))

- **export**: Make inverted-tree rung 3 the demand floor
  ([#628](https://github.com/Teal-Insights/excel-grapher/pull/628),
  [`9d30182`](https://github.com/Teal-Insights/excel-grapher/commit/9d3018220085f390ce6f1627a843d3094e72dc24))

- **export**: Partition catalog statements incrementally in linear time
  ([#628](https://github.com/Teal-Insights/excel-grapher/pull/628),
  [`9d30182`](https://github.com/Teal-Insights/excel-grapher/commit/9d3018220085f390ce6f1627a843d3094e72dc24))

- **export**: Schedule zipper SCCs as demand-driven instances
  ([#628](https://github.com/Teal-Insights/excel-grapher/pull/628),
  [`9d30182`](https://github.com/Teal-Insights/excel-grapher/commit/9d3018220085f390ce6f1627a843d3094e72dc24))

- **export**: Splice mixed-source inverted-tree series by access
  ([#628](https://github.com/Teal-Insights/excel-grapher/pull/628),
  [`9d30182`](https://github.com/Teal-Insights/excel-grapher/commit/9d3018220085f390ce6f1627a843d3094e72dc24))

- **export**: Stream only catalog label cells, stop at last needed row
  ([#628](https://github.com/Teal-Insights/excel-grapher/pull/628),
  [`9d30182`](https://github.com/Teal-Insights/excel-grapher/commit/9d3018220085f390ce6f1627a843d3094e72dc24))

- **export**: Take overlapping series windows at inverted-tree call sites
  ([#628](https://github.com/Teal-Insights/excel-grapher/pull/628),
  [`9d30182`](https://github.com/Teal-Insights/excel-grapher/commit/9d3018220085f390ce6f1627a843d3094e72dc24))

- **export**: Treat residual cycles per schedule column
  ([#628](https://github.com/Teal-Insights/excel-grapher/pull/628),
  [`9d30182`](https://github.com/Teal-Insights/excel-grapher/commit/9d3018220085f390ce6f1627a843d3094e72dc24))

- **export**: Type the fused-scan exec helper for ty
  ([#628](https://github.com/Teal-Insights/excel-grapher/pull/628),
  [`9d30182`](https://github.com/Teal-Insights/excel-grapher/commit/9d3018220085f390ce6f1627a843d3094e72dc24))

### Code Style

- **export**: Format shared-orchestrator unpack join
  ([#628](https://github.com/Teal-Insights/excel-grapher/pull/628),
  [`9d30182`](https://github.com/Teal-Insights/excel-grapher/commit/9d3018220085f390ce6f1627a843d3094e72dc24))

### Documentation

- **export**: Inverted-tree scheduling research memo
  ([#628](https://github.com/Teal-Insights/excel-grapher/pull/628),
  [`9d30182`](https://github.com/Teal-Insights/excel-grapher/commit/9d3018220085f390ce6f1627a843d3094e72dc24))

### Features

- **export**: Accept inverted-tree layout matrix as 1-D sequence
  ([#628](https://github.com/Teal-Insights/excel-grapher/pull/628),
  [`9d30182`](https://github.com/Teal-Insights/excel-grapher/commit/9d3018220085f390ce6f1627a843d3094e72dc24))

- **export**: Add opt-in inverted-tree codegen paradigm
  ([#628](https://github.com/Teal-Insights/excel-grapher/pull/628),
  [`9d30182`](https://github.com/Teal-Insights/excel-grapher/commit/9d3018220085f390ce6f1627a843d3094e72dc24))

- **export**: Add symbolic IndexSet algebra for inverted-tree slices
  ([#628](https://github.com/Teal-Insights/excel-grapher/pull/628),
  [`9d30182`](https://github.com/Teal-Insights/excel-grapher/commit/9d3018220085f390ce6f1627a843d3094e72dc24))

- **export**: Classify inverted-tree affine access maps
  ([#628](https://github.com/Teal-Insights/excel-grapher/pull/628),
  [`9d30182`](https://github.com/Teal-Insights/excel-grapher/commit/9d3018220085f390ce6f1627a843d3094e72dc24))

- **export**: Demote guarded residual may-cycles to rung 3
  ([#628](https://github.com/Teal-Insights/excel-grapher/pull/628),
  [`9d30182`](https://github.com/Teal-Insights/excel-grapher/commit/9d3018220085f390ce6f1627a843d3094e72dc24))

- **export**: Emit fusible zipper SCCs as a union-domain loop
  ([#628](https://github.com/Teal-Insights/excel-grapher/pull/628),
  [`9d30182`](https://github.com/Teal-Insights/excel-grapher/commit/9d3018220085f390ce6f1627a843d3094e72dc24))

- **export**: Emit take() with range and slice forms
  ([#628](https://github.com/Teal-Insights/excel-grapher/pull/628),
  [`9d30182`](https://github.com/Teal-Insights/excel-grapher/commit/9d3018220085f390ce6f1627a843d3094e72dc24))

- **export**: Fuse inverted-tree SCCs by residual-order region
  ([#628](https://github.com/Teal-Insights/excel-grapher/pull/628),
  [`9d30182`](https://github.com/Teal-Insights/excel-grapher/commit/9d3018220085f390ce6f1627a843d3094e72dc24))

- **export**: Generalize rung-1 scan to shift-k self-lags
  ([#628](https://github.com/Teal-Insights/excel-grapher/pull/628),
  [`9d30182`](https://github.com/Teal-Insights/excel-grapher/commit/9d3018220085f390ce6f1627a843d3094e72dc24))

- **export**: Inverted-tree measures are number or error code
  ([#628](https://github.com/Teal-Insights/excel-grapher/pull/628),
  [`9d30182`](https://github.com/Teal-Insights/excel-grapher/commit/9d3018220085f390ce6f1627a843d3094e72dc24))

- **export**: Make DependenceEdge the inverted-tree dep source of truth
  ([#628](https://github.com/Teal-Insights/excel-grapher/pull/628),
  [`9d30182`](https://github.com/Teal-Insights/excel-grapher/commit/9d3018220085f390ce6f1627a843d3094e72dc24))

- **export**: Partition inverted-tree series into statements
  ([#628](https://github.com/Teal-Insights/excel-grapher/pull/628),
  [`9d30182`](https://github.com/Teal-Insights/excel-grapher/commit/9d3018220085f390ce6f1627a843d3094e72dc24))

- **export**: Select reversed loop direction for negative-distance SCCs
  ([#628](https://github.com/Teal-Insights/excel-grapher/pull/628),
  [`9d30182`](https://github.com/Teal-Insights/excel-grapher/commit/9d3018220085f390ce6f1627a843d3094e72dc24))

- **export**: Share inverted-tree orchestrator bodies
  ([#628](https://github.com/Teal-Insights/excel-grapher/pull/628),
  [`9d30182`](https://github.com/Teal-Insights/excel-grapher/commit/9d3018220085f390ce6f1627a843d3094e72dc24))

- **export**: Symbolic IndexSet and take() range/slice forms
  ([#628](https://github.com/Teal-Insights/excel-grapher/pull/628),
  [`9d30182`](https://github.com/Teal-Insights/excel-grapher/commit/9d3018220085f390ce6f1627a843d3094e72dc24))

### Performance Improvements

- **export**: Walk inverted-tree ASTs once and fuse regions in O(E)
  ([#628](https://github.com/Teal-Insights/excel-grapher/pull/628),
  [`9d30182`](https://github.com/Teal-Insights/excel-grapher/commit/9d3018220085f390ce6f1627a843d3094e72dc24))

### Refactoring

- **export**: Inverted-tree layering cleanup (#619)
  ([#628](https://github.com/Teal-Insights/excel-grapher/pull/628),
  [`9d30182`](https://github.com/Teal-Insights/excel-grapher/commit/9d3018220085f390ce6f1627a843d3094e72dc24))

- **export**: Move schedule index into catalog and drop column fallback
  ([#628](https://github.com/Teal-Insights/excel-grapher/pull/628),
  [`9d30182`](https://github.com/Teal-Insights/excel-grapher/commit/9d3018220085f390ce6f1627a843d3094e72dc24))


## v12.4.0 (2026-09-03)

### Bug Fixes

- **export**: Allow inverted-tree other-series t and t-1 reads
  ([#598](https://github.com/Teal-Insights/excel-grapher/pull/598),
  [`50d9ee0`](https://github.com/Teal-Insights/excel-grapher/commit/50d9ee0bed4566444415e29af8b70443c56b0e4d))

- **export**: Apply exclude_rows in inverted-tree catalog
  ([#598](https://github.com/Teal-Insights/excel-grapher/pull/598),
  [`50d9ee0`](https://github.com/Teal-Insights/excel-grapher/commit/50d9ee0bed4566444415e29af8b70443c56b0e4d))

- **export**: Cache inverted-tree schedule coordinates
  ([#598](https://github.com/Teal-Insights/excel-grapher/pull/598),
  [`50d9ee0`](https://github.com/Teal-Insights/excel-grapher/commit/50d9ee0bed4566444415e29af8b70443c56b0e4d))

- **export**: Define inverted-tree xl_exp for EXP
  ([#598](https://github.com/Teal-Insights/excel-grapher/pull/598),
  [`50d9ee0`](https://github.com/Teal-Insights/excel-grapher/commit/50d9ee0bed4566444415e29af8b70443c56b0e4d))

- **export**: Derive inverted-tree identity maps from the key-point join
  ([#598](https://github.com/Teal-Insights/excel-grapher/pull/598),
  [`50d9ee0`](https://github.com/Teal-Insights/excel-grapher/commit/50d9ee0bed4566444415e29af8b70443c56b0e4d))

- **export**: Drive rung-3 backward chains in reverse to avoid recursion
  ([#598](https://github.com/Teal-Insights/excel-grapher/pull/598),
  [`50d9ee0`](https://github.com/Teal-Insights/excel-grapher/commit/50d9ee0bed4566444415e29af8b70443c56b0e4d))

- **export**: Emit lag-zipper series as a co-scan
  ([#598](https://github.com/Teal-Insights/excel-grapher/pull/598),
  [`50d9ee0`](https://github.com/Teal-Insights/excel-grapher/commit/50d9ee0bed4566444415e29af8b70443c56b0e4d))

- **export**: Emit live_measure for fused refs off the SCC union (#623)
  ([#598](https://github.com/Teal-Insights/excel-grapher/pull/598),
  [`50d9ee0`](https://github.com/Teal-Insights/excel-grapher/commit/50d9ee0bed4566444415e29af8b70443c56b0e4d))

- **export**: Fail closed on partially resolved inverted-tree keys
  ([#598](https://github.com/Teal-Insights/excel-grapher/pull/598),
  [`50d9ee0`](https://github.com/Teal-Insights/excel-grapher/commit/50d9ee0bed4566444415e29af8b70443c56b0e4d))

- **export**: Gather inverted-tree series with take, not trim
  ([#598](https://github.com/Teal-Insights/excel-grapher/pull/598),
  [`50d9ee0`](https://github.com/Teal-Insights/excel-grapher/commit/50d9ee0bed4566444415e29af8b70443c56b0e4d))

- **export**: Inverted-tree emit nits for errors, SCC call, and else
  ([#598](https://github.com/Teal-Insights/excel-grapher/pull/598),
  [`50d9ee0`](https://github.com/Teal-Insights/excel-grapher/commit/50d9ee0bed4566444415e29af8b70443c56b0e4d))

- **export**: Join inverted-tree matrix schedules on the full key tuple
  ([#598](https://github.com/Teal-Insights/excel-grapher/pull/598),
  [`50d9ee0`](https://github.com/Teal-Insights/excel-grapher/commit/50d9ee0bed4566444415e29af8b70443c56b0e4d))

- **export**: Join inverted-tree schedule on key-point domain
  ([#598](https://github.com/Teal-Insights/excel-grapher/pull/598),
  [`50d9ee0`](https://github.com/Teal-Insights/excel-grapher/commit/50d9ee0bed4566444415e29af8b70443c56b0e4d))

- **export**: Keep IMF sentinels in float inverted-tree constants
  ([#598](https://github.com/Teal-Insights/excel-grapher/pull/598),
  [`50d9ee0`](https://github.com/Teal-Insights/excel-grapher/commit/50d9ee0bed4566444415e29af8b70443c56b0e4d))

- **export**: Make inverted-tree rung 3 the demand floor
  ([#598](https://github.com/Teal-Insights/excel-grapher/pull/598),
  [`50d9ee0`](https://github.com/Teal-Insights/excel-grapher/commit/50d9ee0bed4566444415e29af8b70443c56b0e4d))

- **export**: Partition catalog statements incrementally in linear time
  ([#598](https://github.com/Teal-Insights/excel-grapher/pull/598),
  [`50d9ee0`](https://github.com/Teal-Insights/excel-grapher/commit/50d9ee0bed4566444415e29af8b70443c56b0e4d))

- **export**: Schedule zipper SCCs as demand-driven instances
  ([#598](https://github.com/Teal-Insights/excel-grapher/pull/598),
  [`50d9ee0`](https://github.com/Teal-Insights/excel-grapher/commit/50d9ee0bed4566444415e29af8b70443c56b0e4d))

- **export**: Splice mixed-source inverted-tree series by access
  ([#598](https://github.com/Teal-Insights/excel-grapher/pull/598),
  [`50d9ee0`](https://github.com/Teal-Insights/excel-grapher/commit/50d9ee0bed4566444415e29af8b70443c56b0e4d))

- **export**: Stream only catalog label cells, stop at last needed row
  ([#598](https://github.com/Teal-Insights/excel-grapher/pull/598),
  [`50d9ee0`](https://github.com/Teal-Insights/excel-grapher/commit/50d9ee0bed4566444415e29af8b70443c56b0e4d))

- **export**: Take overlapping series windows at inverted-tree call sites
  ([#598](https://github.com/Teal-Insights/excel-grapher/pull/598),
  [`50d9ee0`](https://github.com/Teal-Insights/excel-grapher/commit/50d9ee0bed4566444415e29af8b70443c56b0e4d))

- **export**: Treat residual cycles per schedule column
  ([#598](https://github.com/Teal-Insights/excel-grapher/pull/598),
  [`50d9ee0`](https://github.com/Teal-Insights/excel-grapher/commit/50d9ee0bed4566444415e29af8b70443c56b0e4d))

- **export**: Type the fused-scan exec helper for ty
  ([#598](https://github.com/Teal-Insights/excel-grapher/pull/598),
  [`50d9ee0`](https://github.com/Teal-Insights/excel-grapher/commit/50d9ee0bed4566444415e29af8b70443c56b0e4d))

### Code Style

- **export**: Format shared-orchestrator unpack join
  ([#598](https://github.com/Teal-Insights/excel-grapher/pull/598),
  [`50d9ee0`](https://github.com/Teal-Insights/excel-grapher/commit/50d9ee0bed4566444415e29af8b70443c56b0e4d))

### Documentation

- **export**: Inverted-tree scheduling research memo
  ([#598](https://github.com/Teal-Insights/excel-grapher/pull/598),
  [`50d9ee0`](https://github.com/Teal-Insights/excel-grapher/commit/50d9ee0bed4566444415e29af8b70443c56b0e4d))

### Features

- **export**: Accept inverted-tree layout matrix as 1-D sequence
  ([#598](https://github.com/Teal-Insights/excel-grapher/pull/598),
  [`50d9ee0`](https://github.com/Teal-Insights/excel-grapher/commit/50d9ee0bed4566444415e29af8b70443c56b0e4d))

- **export**: Add opt-in inverted-tree codegen paradigm
  ([#598](https://github.com/Teal-Insights/excel-grapher/pull/598),
  [`50d9ee0`](https://github.com/Teal-Insights/excel-grapher/commit/50d9ee0bed4566444415e29af8b70443c56b0e4d))

- **export**: Add symbolic IndexSet algebra for inverted-tree slices
  ([#598](https://github.com/Teal-Insights/excel-grapher/pull/598),
  [`50d9ee0`](https://github.com/Teal-Insights/excel-grapher/commit/50d9ee0bed4566444415e29af8b70443c56b0e4d))

- **export**: Classify inverted-tree affine access maps
  ([#598](https://github.com/Teal-Insights/excel-grapher/pull/598),
  [`50d9ee0`](https://github.com/Teal-Insights/excel-grapher/commit/50d9ee0bed4566444415e29af8b70443c56b0e4d))

- **export**: Demote guarded residual may-cycles to rung 3
  ([#598](https://github.com/Teal-Insights/excel-grapher/pull/598),
  [`50d9ee0`](https://github.com/Teal-Insights/excel-grapher/commit/50d9ee0bed4566444415e29af8b70443c56b0e4d))

- **export**: Emit fusible zipper SCCs as a union-domain loop
  ([#598](https://github.com/Teal-Insights/excel-grapher/pull/598),
  [`50d9ee0`](https://github.com/Teal-Insights/excel-grapher/commit/50d9ee0bed4566444415e29af8b70443c56b0e4d))

- **export**: Emit take() with range and slice forms
  ([#598](https://github.com/Teal-Insights/excel-grapher/pull/598),
  [`50d9ee0`](https://github.com/Teal-Insights/excel-grapher/commit/50d9ee0bed4566444415e29af8b70443c56b0e4d))

- **export**: Fuse inverted-tree SCCs by residual-order region
  ([#598](https://github.com/Teal-Insights/excel-grapher/pull/598),
  [`50d9ee0`](https://github.com/Teal-Insights/excel-grapher/commit/50d9ee0bed4566444415e29af8b70443c56b0e4d))

- **export**: Generalize rung-1 scan to shift-k self-lags
  ([#598](https://github.com/Teal-Insights/excel-grapher/pull/598),
  [`50d9ee0`](https://github.com/Teal-Insights/excel-grapher/commit/50d9ee0bed4566444415e29af8b70443c56b0e4d))

- **export**: Inverted-tree codegen (opt-in prototype for #597)
  ([#598](https://github.com/Teal-Insights/excel-grapher/pull/598),
  [`50d9ee0`](https://github.com/Teal-Insights/excel-grapher/commit/50d9ee0bed4566444415e29af8b70443c56b0e4d))

- **export**: Inverted-tree measures are number or error code
  ([#598](https://github.com/Teal-Insights/excel-grapher/pull/598),
  [`50d9ee0`](https://github.com/Teal-Insights/excel-grapher/commit/50d9ee0bed4566444415e29af8b70443c56b0e4d))

- **export**: Make DependenceEdge the inverted-tree dep source of truth
  ([#598](https://github.com/Teal-Insights/excel-grapher/pull/598),
  [`50d9ee0`](https://github.com/Teal-Insights/excel-grapher/commit/50d9ee0bed4566444415e29af8b70443c56b0e4d))

- **export**: Partition inverted-tree series into statements
  ([#598](https://github.com/Teal-Insights/excel-grapher/pull/598),
  [`50d9ee0`](https://github.com/Teal-Insights/excel-grapher/commit/50d9ee0bed4566444415e29af8b70443c56b0e4d))

- **export**: Select reversed loop direction for negative-distance SCCs
  ([#598](https://github.com/Teal-Insights/excel-grapher/pull/598),
  [`50d9ee0`](https://github.com/Teal-Insights/excel-grapher/commit/50d9ee0bed4566444415e29af8b70443c56b0e4d))

- **export**: Share inverted-tree orchestrator bodies
  ([#598](https://github.com/Teal-Insights/excel-grapher/pull/598),
  [`50d9ee0`](https://github.com/Teal-Insights/excel-grapher/commit/50d9ee0bed4566444415e29af8b70443c56b0e4d))

- **export**: Symbolic IndexSet and take() range/slice forms
  ([#598](https://github.com/Teal-Insights/excel-grapher/pull/598),
  [`50d9ee0`](https://github.com/Teal-Insights/excel-grapher/commit/50d9ee0bed4566444415e29af8b70443c56b0e4d))

### Performance Improvements

- **export**: Walk inverted-tree ASTs once and fuse regions in O(E)
  ([#598](https://github.com/Teal-Insights/excel-grapher/pull/598),
  [`50d9ee0`](https://github.com/Teal-Insights/excel-grapher/commit/50d9ee0bed4566444415e29af8b70443c56b0e4d))


## v12.3.1 (2026-09-01)

### Bug Fixes

- **core**: Accept FormulaStyle member-name strings (#584)
  ([#592](https://github.com/Teal-Insights/excel-grapher/pull/592),
  [`f620088`](https://github.com/Teal-Insights/excel-grapher/commit/f62008827288f6d08791a538e95b090729bb4629))


## v12.3.0 (2026-08-31)

### Features

- **series-bindings**: Merge complementary shards across sheets
  ([#591](https://github.com/Teal-Insights/excel-grapher/pull/591),
  [`e8b5f7f`](https://github.com/Teal-Insights/excel-grapher/commit/e8b5f7fd2a84db5a47cdc1e7aaf2ee7689827210))


## v12.2.0 (2026-08-31)

### Features

- **grapher**: Raise default max_range_cells to 50_000
  ([#589](https://github.com/Teal-Insights/excel-grapher/pull/589),
  [`057b6d4`](https://github.com/Teal-Insights/excel-grapher/commit/057b6d4ae206c7a6a7a106ed66780ccb150db7c4))


## v12.1.0 (2026-08-31)

### Features

- **exporter**: Record range watches for export invalidation
  ([#585](https://github.com/Teal-Insights/excel-grapher/pull/585),
  [`842acce`](https://github.com/Teal-Insights/excel-grapher/commit/842accececbb111ffda3021fc09cdfe6e0090d0e))


## v12.0.0 (2026-08-31)

### Bug Fixes

- **grapher**: Raise when expand_range exceeds max_cells
  ([#587](https://github.com/Teal-Insights/excel-grapher/pull/587),
  [`34858b1`](https://github.com/Teal-Insights/excel-grapher/commit/34858b149bcadcc22078a4081f3f5796f7e0f2c1))


## v11.5.0 (2026-08-31)

### Features

- **exporter**: Emit CONSTANTS as MappingProxyType
  ([#582](https://github.com/Teal-Insights/excel-grapher/pull/582),
  [`0a376dd`](https://github.com/Teal-Insights/excel-grapher/commit/0a376dd25e0becae728263e12e4bf69a775bc2eb))

### Testing

- **runtime**: Avoid ty invalid-assignment on frozen CONSTANTS
  ([#582](https://github.com/Teal-Insights/excel-grapher/pull/582),
  [`0a376dd`](https://github.com/Teal-Insights/excel-grapher/commit/0a376dd25e0becae728263e12e4bf69a775bc2eb))


## v11.4.0 (2026-08-31)

### Bug Fixes

- **scripts**: Type the leaf-store scan as LeafStore for ty
  ([`7f814c0`](https://github.com/Teal-Insights/excel-grapher/commit/7f814c0bd23faf06e516b1b73ff010434ce2c6d6))

### Documentation

- Note projection rewarm for write_workbook
  ([#578](https://github.com/Teal-Insights/excel-grapher/pull/578),
  [`7b58489`](https://github.com/Teal-Insights/excel-grapher/commit/7b58489cb2320fd11cd15a6df58f530b5859264a))

- Note rewarming shapes before write_workbook(projection)
  ([#578](https://github.com/Teal-Insights/excel-grapher/pull/578),
  [`7b58489`](https://github.com/Teal-Insights/excel-grapher/commit/7b58489cb2320fd11cd15a6df58f530b5859264a))

- Remove duplicate changelog entries for v11.0.0 and v11.3.0
  ([#578](https://github.com/Teal-Insights/excel-grapher/pull/578),
  [`7b58489`](https://github.com/Teal-Insights/excel-grapher/commit/7b58489cb2320fd11cd15a6df58f530b5859264a))

### Features

- **exporter**: Emit leaf values as a sparse coordinate store
  ([`b8b27ad`](https://github.com/Teal-Insights/excel-grapher/commit/b8b27ad095fa582577ce53cd6d77142d88acbdaf))


## v11.3.0 (2026-08-31)

### Documentation

- Distinguish move-then-write from project-then-write
  ([#575](https://github.com/Teal-Insights/excel-grapher/pull/575),
  [`ca38ace`](https://github.com/Teal-Insights/excel-grapher/commit/ca38acee481a115b32f1a88104ad067539b51aad))

### Features

- **grapher**: Emit Excel shared formulas from interned shapes
  ([#576](https://github.com/Teal-Insights/excel-grapher/pull/576),
  [`c72ec83`](https://github.com/Teal-Insights/excel-grapher/commit/c72ec8331dbc981a0e23431c782c24e276bffa29))

### Testing

- **grapher**: Round-trip write_workbook after move_node
  ([#575](https://github.com/Teal-Insights/excel-grapher/pull/575),
  [`ca38ace`](https://github.com/Teal-Insights/excel-grapher/commit/ca38acee481a115b32f1a88104ad067539b51aad))


## v11.2.0 (2026-08-31)

### Features

- **grapher**: Write defined-name table from graph maps
  ([#574](https://github.com/Teal-Insights/excel-grapher/pull/574),
  [`cde45de`](https://github.com/Teal-Insights/excel-grapher/commit/cde45dee1629979b11653a88e4ca195fea69a60f))


## v11.1.0 (2026-08-31)

### Features

- **grapher**: Persist array-formula provenance on write-back
  ([#573](https://github.com/Teal-Insights/excel-grapher/pull/573),
  [`02ef429`](https://github.com/Teal-Insights/excel-grapher/commit/02ef429f31c5e02cda6e418357893decf98dc3d4))

- **grapher**: Write GraphReadView to a new xlsx
  ([#572](https://github.com/Teal-Insights/excel-grapher/pull/572),
  [`f1ef1bb`](https://github.com/Teal-Insights/excel-grapher/commit/f1ef1bba2836632ff3bb4d0097b8c0c0c03acd40))

### Testing

- **grapher**: Bypass address guard in write_workbook missing-anchor case
  ([#573](https://github.com/Teal-Insights/excel-grapher/pull/573),
  [`02ef429`](https://github.com/Teal-Insights/excel-grapher/commit/02ef429f31c5e02cda6e418357893decf98dc3d4))


## v11.0.0 (2026-08-31)

### Features

- **grapher**: Preserve axis intent on remaining string→AST entry points
  ([#563](https://github.com/Teal-Insights/excel-grapher/pull/563),
  [`79a08c6`](https://github.com/Teal-Insights/excel-grapher/commit/79a08c6c7c2c99bc83c3348e6d7de3f227795266))

### Testing

- **grapher**: Construct formula nodes from Excel-style axis text
  ([#563](https://github.com/Teal-Insights/excel-grapher/pull/563),
  [`79a08c6`](https://github.com/Teal-Insights/excel-grapher/commit/79a08c6c7c2c99bc83c3348e6d7de3f227795266))


## v10.3.0 (2026-08-31)

### Bug Fixes

- **grapher**: Drop occupancy edges and merge guards on move_node
  ([#571](https://github.com/Teal-Insights/excel-grapher/pull/571),
  [`881a829`](https://github.com/Teal-Insights/excel-grapher/commit/881a829a7755d15cb9ecdd7a7fae6ccda7e3c5b6))

### Features

- **grapher**: Move_node rewrites relative refs on geometry change
  ([#571](https://github.com/Teal-Insights/excel-grapher/pull/571),
  [`881a829`](https://github.com/Teal-Insights/excel-grapher/commit/881a829a7755d15cb9ecdd7a7fae6ccda7e3c5b6))

- **grapher**: Rewrite relative refs when NodeKey geometry changes
  ([#571](https://github.com/Teal-Insights/excel-grapher/pull/571),
  [`881a829`](https://github.com/Teal-Insights/excel-grapher/commit/881a829a7755d15cb9ecdd7a7fae6ccda7e3c5b6))

### Testing

- **grapher**: Fix move_node formula-shape test name
  ([#571](https://github.com/Teal-Insights/excel-grapher/pull/571),
  [`881a829`](https://github.com/Teal-Insights/excel-grapher/commit/881a829a7755d15cb9ecdd7a7fae6ccda7e3c5b6))


## v10.2.0 (2026-08-31)

### Bug Fixes

- **grapher**: Skip identity compression for range and whole-col/row sites
  ([#570](https://github.com/Teal-Insights/excel-grapher/pull/570),
  [`501b265`](https://github.com/Teal-Insights/excel-grapher/commit/501b2654ed3df3aa7bebfe14dbb41dd91a0ce3c9))

### Code Style

- **core**: Apply ruff format to address-leaf rewrite helpers
  ([#570](https://github.com/Teal-Insights/excel-grapher/pull/570),
  [`501b265`](https://github.com/Teal-Insights/excel-grapher/commit/501b2654ed3df3aa7bebfe14dbb41dd91a0ce3c9))

### Documentation

- Formula-shape overlay is opt-in; caller must rewarm
  ([`16fec80`](https://github.com/Teal-Insights/excel-grapher/commit/16fec80e62ad289bca1569d0cf7c17255ad8c927))

### Features

- **core**: Retarget range and whole-col/row leaves in AST rewrite
  ([#570](https://github.com/Teal-Insights/excel-grapher/pull/570),
  [`501b265`](https://github.com/Teal-Insights/excel-grapher/commit/501b2654ed3df3aa7bebfe14dbb41dd91a0ce3c9))

- **core**: Retarget range endpoints; fail-closed identity compression
  ([#570](https://github.com/Teal-Insights/excel-grapher/pull/570),
  [`501b265`](https://github.com/Teal-Insights/excel-grapher/commit/501b2654ed3df3aa7bebfe14dbb41dd91a0ce3c9))

### Testing

- **evaluator**: Lock formula-shape init snapshot after rewarm
  ([`f24a11c`](https://github.com/Teal-Insights/excel-grapher/commit/f24a11c00a75ca8d03f98a5b9c2f531fe0318697))


## v10.1.0 (2026-08-30)

### Bug Fixes

- **evaluator**: Bind relative formula ASTs before string-cache seed
  ([`b859657`](https://github.com/Teal-Insights/excel-grapher/commit/b8596570a8283e85ff05b57538b9dde45d57e2d5))

### Features

- **grapher**: Retire regex normalizer as peer of AST render dialect
  ([#559](https://github.com/Teal-Insights/excel-grapher/pull/559),
  [`34583d1`](https://github.com/Teal-Insights/excel-grapher/commit/34583d104f03ea5c15a6fc52c8556ed08ada5b85))


## v10.0.0 (2026-08-30)

### Features

- **grapher**: Render_formula API and drop stored normalized_formula
  ([#553](https://github.com/Teal-Insights/excel-grapher/pull/553),
  [`8e684c9`](https://github.com/Teal-Insights/excel-grapher/commit/8e684c963e28353c5a6f4a46b0f9e0ec6094641e))

- **grapher**: Render_formula API and drop stored normalized_formula (#543)
  ([#553](https://github.com/Teal-Insights/excel-grapher/pull/553),
  [`8e684c9`](https://github.com/Teal-Insights/excel-grapher/commit/8e684c963e28353c5a6f4a46b0f9e0ec6094641e))

### Testing

- **grapher**: Cover JSON cache load of legacy normalized_formula
  ([#553](https://github.com/Teal-Insights/excel-grapher/pull/553),
  [`8e684c9`](https://github.com/Teal-Insights/excel-grapher/commit/8e684c963e28353c5a6f4a46b0f9e0ec6094641e))


## v9.1.0 (2026-08-30)

### Bug Fixes

- **core**: Retarget identity cell refs without dropping relative axes
  ([#552](https://github.com/Teal-Insights/excel-grapher/pull/552),
  [`ef56099`](https://github.com/Teal-Insights/excel-grapher/commit/ef56099f6af046e75ab3701e812a95ff47a02cc8))

- **grapher**: Derive compressed A1 text from rewritten AST
  ([#552](https://github.com/Teal-Insights/excel-grapher/pull/552),
  [`ef56099`](https://github.com/Teal-Insights/excel-grapher/commit/ef56099f6af046e75ab3701e812a95ff47a02cc8))

- **grapher**: Key type analysis by stored formula_ast
  ([#552](https://github.com/Teal-Insights/excel-grapher/pull/552),
  [`ef56099`](https://github.com/Teal-Insights/excel-grapher/commit/ef56099f6af046e75ab3701e812a95ff47a02cc8))

- **grapher**: Preserve raw formula and relative axes on node setters
  ([#552](https://github.com/Teal-Insights/excel-grapher/pull/552),
  [`ef56099`](https://github.com/Teal-Insights/excel-grapher/commit/ef56099f6af046e75ab3701e812a95ff47a02cc8))

### Features

- **grapher**: Migrate formula consumers onto formula_ast
  ([#552](https://github.com/Teal-Insights/excel-grapher/pull/552),
  [`ef56099`](https://github.com/Teal-Insights/excel-grapher/commit/ef56099f6af046e75ab3701e812a95ff47a02cc8))

### Testing

- **grapher**: Satisfy ty on stale-A1 codegen exec helpers
  ([#552](https://github.com/Teal-Insights/excel-grapher/pull/552),
  [`ef56099`](https://github.com/Teal-Insights/excel-grapher/commit/ef56099f6af046e75ab3701e812a95ff47a02cc8))


## v9.0.0 (2026-08-29)

### Features

- **core**: Intern formula ASTs by hashable tree identity
  ([`43f1248`](https://github.com/Teal-Insights/excel-grapher/commit/43f1248ec48cd27bfcfa9f227045dc5b353edaa7))


## v8.0.0 (2026-08-29)

### Bug Fixes

- **core**: Satisfy ty on axis-aware parse and whole-row deps
  ([`69a2707`](https://github.com/Teal-Insights/excel-grapher/commit/69a27079d5fc21ae66d582a362dc2f9dadf3c9bf))

### Features

- **core**: Preserve per-axis relative/absolute refs in formula AST
  ([#545](https://github.com/Teal-Insights/excel-grapher/pull/545),
  [`9f62116`](https://github.com/Teal-Insights/excel-grapher/commit/9f621166dbfaf21b7696da6c45f60db18e142c08))


## v7.0.0 (2026-08-29)

### Bug Fixes

- **exporter**: Import datetime in internals for reader kwargs
  ([`3012fc4`](https://github.com/Teal-Insights/excel-grapher/commit/3012fc435a3b0f0191db8c4ea54a11d51233a9ff))

- **grapher**: Leave formula_ast unset when extraction parse fails
  ([`42e6bab`](https://github.com/Teal-Insights/excel-grapher/commit/42e6bab55d51f66a6f39e6c04eb5e3e18312259d))

### Features

- **grapher**: Intern cached formula ASTs and fail-soft on rewrite parse
  ([`f2fdf34`](https://github.com/Teal-Insights/excel-grapher/commit/f2fdf341a1623ed5c3880fb3d332e910b24839f2))

- **grapher**: Store per-node formula AST at extraction
  ([`82cc9a2`](https://github.com/Teal-Insights/excel-grapher/commit/82cc9a2e8551807253361e48aaaccc11d9010f47))

### Refactoring

- **grapher**: Drop duplicate intern test and getattr formula_ast
  ([`0bb5c10`](https://github.com/Teal-Insights/excel-grapher/commit/0bb5c10648d64404cce686796a893a86fc12b822))

### Breaking Changes

- **grapher**: FormulaShapeTable.lookup and intern_formula_shapes key bindings by NodeKey rather
  than normalized formula text.


## v6.3.3 (2026-08-19)

### Bug Fixes

- **grapher**: Share argument-subgraph ref walks with provenance
  ([#540](https://github.com/Teal-Insights/excel-grapher/pull/540),
  [`3decc35`](https://github.com/Teal-Insights/excel-grapher/commit/3decc352ece2a1a98ffd941846fc19cbf80c81bd))


## v6.3.2 (2026-08-19)

### Bug Fixes

- **grapher**: Retain is_target nodes in identity-transit compression
  ([#538](https://github.com/Teal-Insights/excel-grapher/pull/538),
  [`1544590`](https://github.com/Teal-Insights/excel-grapher/commit/15445909e355c18181aff62e335f6d94f3931346))


## v6.3.1 (2026-08-19)

### Bug Fixes

- **exporter**: Skip shared helpers for OFFSET, INDEX, and 1x1 ranges
  ([#536](https://github.com/Teal-Insights/excel-grapher/pull/536),
  [`1e8b9c9`](https://github.com/Teal-Insights/excel-grapher/commit/1e8b9c9b48728a2749185a010e103043ccb8e35d))


## v6.3.0 (2026-08-19)

### Features

- Intern parameterized formula AST shapes for parse, eval, and codegen (#517)
  ([#534](https://github.com/Teal-Insights/excel-grapher/pull/534),
  [`c452449`](https://github.com/Teal-Insights/excel-grapher/commit/c452449da01e0fbadc4637844f3af65a24b30494))

- **evaluator**: Compile interned formula shapes
  ([#534](https://github.com/Teal-Insights/excel-grapher/pull/534),
  [`c452449`](https://github.com/Teal-Insights/excel-grapher/commit/c452449da01e0fbadc4637844f3af65a24b30494))

- **grapher**: Intern parameterized formula AST shapes
  ([#534](https://github.com/Teal-Insights/excel-grapher/pull/534),
  [`c452449`](https://github.com/Teal-Insights/excel-grapher/commit/c452449da01e0fbadc4637844f3af65a24b30494))

- **scripts**: Measure shape eval and codegen wins
  ([#534](https://github.com/Teal-Insights/excel-grapher/pull/534),
  [`c452449`](https://github.com/Teal-Insights/excel-grapher/commit/c452449da01e0fbadc4637844f3af65a24b30494))


## v6.2.0 (2026-08-19)

### Features

- **grapher**: Record INDEX targets in dependency provenance
  ([#532](https://github.com/Teal-Insights/excel-grapher/pull/532),
  [`c93ed9f`](https://github.com/Teal-Insights/excel-grapher/commit/c93ed9f337dce2bab1614e1f7d4f181cf0650209))


## v6.1.2 (2026-08-18)

### Bug Fixes

- **grapher**: Isolate argument-subgraph memo and cover OFFSET reuse
  ([#530](https://github.com/Teal-Insights/excel-grapher/pull/530),
  [`948cc56`](https://github.com/Teal-Insights/excel-grapher/commit/948cc56a96b3107f26a1472b9fc05295b699c004))

### Performance Improvements

- **grapher**: Reuse argument-env expansion across formula variants
  ([#530](https://github.com/Teal-Insights/excel-grapher/pull/530),
  [`948cc56`](https://github.com/Teal-Insights/excel-grapher/commit/948cc56a96b3107f26a1472b9fc05295b699c004))


## v6.1.1 (2026-08-17)

### Performance Improvements

- **grapher**: Cache _qualify_fragment defined-name regexes
  ([#529](https://github.com/Teal-Insights/excel-grapher/pull/529),
  [`8a2396f`](https://github.com/Teal-Insights/excel-grapher/commit/8a2396f729501eca2cb8455f1b34b7faed0f4897))


## v6.1.0 (2026-08-12)

### Bug Fixes

- **grapher**: Narrow RangeRef.element return after intern
  ([#522](https://github.com/Teal-Insights/excel-grapher/pull/522),
  [`29a62a6`](https://github.com/Teal-Insights/excel-grapher/commit/29a62a60ec92efa26382232e9286d480c509b0fc))

- **grapher**: Weak intern pool and cheap uncond adjacency check
  ([#522](https://github.com/Teal-Insights/excel-grapher/pull/522),
  [`29a62a6`](https://github.com/Teal-Insights/excel-grapher/commit/29a62a60ec92efa26382232e9286d480c509b0fc))

### Features

- **grapher**: Intern GuardExpr trees and add is_guarded
  ([#522](https://github.com/Teal-Insights/excel-grapher/pull/522),
  [`29a62a6`](https://github.com/Teal-Insights/excel-grapher/commit/29a62a60ec92efa26382232e9286d480c509b0fc))

- **grapher**: Intern GuardExpr trees and add is_guarded (#491)
  ([#522](https://github.com/Teal-Insights/excel-grapher/pull/522),
  [`29a62a6`](https://github.com/Teal-Insights/excel-grapher/commit/29a62a60ec92efa26382232e9286d480c509b0fc))


## v6.0.0 (2026-08-12)

### Bug Fixes

- **grapher**: Require CellKey when loading graph cache nodes
  ([#524](https://github.com/Teal-Insights/excel-grapher/pull/524),
  [`da8fbdb`](https://github.com/Teal-Insights/excel-grapher/commit/da8fbdbef5ca9879cbc599750e4b50fc9ab87796))

### Refactoring

- Remove multi-member occupancy from DependencyGraph
  ([#524](https://github.com/Teal-Insights/excel-grapher/pull/524),
  [`da8fbdb`](https://github.com/Teal-Insights/excel-grapher/commit/da8fbdbef5ca9879cbc599750e4b50fc9ab87796))

### Testing

- Relax unpickle peak/current bound after occupancy removal
  ([#524](https://github.com/Teal-Insights/excel-grapher/pull/524),
  [`da8fbdb`](https://github.com/Teal-Insights/excel-grapher/commit/da8fbdbef5ca9879cbc599750e4b50fc9ab87796))


## v5.2.0 (2026-08-12)

### Bug Fixes

- **core**: Avoid grapher import in formula_shape summary
  ([#518](https://github.com/Teal-Insights/excel-grapher/pull/518),
  [`884bb28`](https://github.com/Teal-Insights/excel-grapher/commit/884bb2819631d31f16ed9c21479af4d5317246e4))

- **core**: Correct mean_instances_per_shape and share summarize helper
  ([#518](https://github.com/Teal-Insights/excel-grapher/pull/518),
  [`884bb28`](https://github.com/Teal-Insights/excel-grapher/commit/884bb2819631d31f16ed9c21479af4d5317246e4))

- **grapher**: Document keep_formula_cache load path
  ([#519](https://github.com/Teal-Insights/excel-grapher/pull/519),
  [`fdaa645`](https://github.com/Teal-Insights/excel-grapher/commit/fdaa6452a2dc25303231916d904509a2802c8cfd))

### Features

- **core**: Fingerprint parameterized formula AST shapes
  ([#518](https://github.com/Teal-Insights/excel-grapher/pull/518),
  [`884bb28`](https://github.com/Teal-Insights/excel-grapher/commit/884bb2819631d31f16ed9c21479af4d5317246e4))

- **core**: Validate parameterized formula AST shape interning (#517)
  ([#518](https://github.com/Teal-Insights/excel-grapher/pull/518),
  [`884bb28`](https://github.com/Teal-Insights/excel-grapher/commit/884bb2819631d31f16ed9c21479af4d5317246e4))

### Performance Improvements

- **grapher**: Load formulas and caches in one fastpyxl pass
  ([#519](https://github.com/Teal-Insights/excel-grapher/pull/519),
  [`fdaa645`](https://github.com/Teal-Insights/excel-grapher/commit/fdaa6452a2dc25303231916d904509a2802c8cfd))


## v5.1.6 (2026-08-11)

### Bug Fixes

- **grapher**: Exclude ROW/COLUMN address-only refs from dependency edges
  ([#516](https://github.com/Teal-Insights/excel-grapher/pull/516),
  [`144afc3`](https://github.com/Teal-Insights/excel-grapher/commit/144afc37e260a29b716f8a06679e14d869c4e41d))

### Code Style

- Fix import formatting after ROW/COLUMN dep exclusion
  ([#516](https://github.com/Teal-Insights/excel-grapher/pull/516),
  [`144afc3`](https://github.com/Teal-Insights/excel-grapher/commit/144afc37e260a29b716f8a06679e14d869c4e41d))


## v5.1.5 (2026-08-09)

### Bug Fixes

- **grapher**: Cut DependencyGraph unpickle peak memory
  ([#514](https://github.com/Teal-Insights/excel-grapher/pull/514),
  [`95f865a`](https://github.com/Teal-Insights/excel-grapher/commit/95f865a57873e5bef8a2a9fe1e3721cbc3a8b737))

- **grapher**: Cut DependencyGraph unpickle peak memory (#513)
  ([#514](https://github.com/Teal-Insights/excel-grapher/pull/514),
  [`95f865a`](https://github.com/Teal-Insights/excel-grapher/commit/95f865a57873e5bef8a2a9fe1e3721cbc3a8b737))

### Testing

- **grapher**: Avoid monkeypatching __reduce_ex__ in legacy load test
  ([#514](https://github.com/Teal-Insights/excel-grapher/pull/514),
  [`95f865a`](https://github.com/Teal-Insights/excel-grapher/commit/95f865a57873e5bef8a2a9fe1e3721cbc3a8b737))


## v5.1.4 (2026-08-08)

### Bug Fixes

- **series-bindings**: Allow empty shards and union concept schemes
  ([#512](https://github.com/Teal-Insights/excel-grapher/pull/512),
  [`cf50a1d`](https://github.com/Teal-Insights/excel-grapher/commit/cf50a1d160261452f0f821e1e5e07e6740460cb9))


## v5.1.3 (2026-08-08)

### Bug Fixes

- **core**: INDEX row/col 0 returns whole axis
  ([#509](https://github.com/Teal-Insights/excel-grapher/pull/509),
  [`1f16849`](https://github.com/Teal-Insights/excel-grapher/commit/1f16849fdc8a1f1b317f42671648c67a9e90edf5))

- **core**: INDEX(array, 0) returns whole column/row (#502)
  ([#509](https://github.com/Teal-Insights/excel-grapher/pull/509),
  [`1f16849`](https://github.com/Teal-Insights/excel-grapher/commit/1f16849fdc8a1f1b317f42671648c67a9e90edf5))

- **evaluator**: Satisfy ty on value-mode INDEX args
  ([#509](https://github.com/Teal-Insights/excel-grapher/pull/509),
  [`1f16849`](https://github.com/Teal-Insights/excel-grapher/commit/1f16849fdc8a1f1b317f42671648c67a9e90edf5))

### Code Style

- **evaluator**: Inline INDEX reference predicate
  ([#509](https://github.com/Teal-Insights/excel-grapher/pull/509),
  [`1f16849`](https://github.com/Teal-Insights/excel-grapher/commit/1f16849fdc8a1f1b317f42671648c67a9e90edf5))


## v5.1.2 (2026-08-08)

### Bug Fixes

- **core**: Align INDEX zero selectors and computed-array path (#503)
  ([#510](https://github.com/Teal-Insights/excel-grapher/pull/510),
  [`7ef28f1`](https://github.com/Teal-Insights/excel-grapher/commit/7ef28f1c2e1802b94d407fffa39d4d42bc7b21ff))

- **core**: LOOKUP skips lookup-vector errors (#504)
  ([#508](https://github.com/Teal-Insights/excel-grapher/pull/508),
  [`c3e4a4f`](https://github.com/Teal-Insights/excel-grapher/commit/c3e4a4f709a0cbb063fbc2fa714f5ab7e6e3a52f))

- **core**: LOOKUP skips lookup-vector errors; array arithmetic preserves them
  ([#508](https://github.com/Teal-Insights/excel-grapher/pull/508),
  [`c3e4a4f`](https://github.com/Teal-Insights/excel-grapher/commit/c3e4a4f709a0cbb063fbc2fa714f5ab7e6e3a52f))

### Testing

- **evaluator**: Expect array arithmetic to preserve embedded errors
  ([#508](https://github.com/Teal-Insights/excel-grapher/pull/508),
  [`c3e4a4f`](https://github.com/Teal-Insights/excel-grapher/commit/c3e4a4f709a0cbb063fbc2fa714f5ab7e6e3a52f))


## v5.1.1 (2026-08-08)

### Bug Fixes

- **grapher**: Infer MATCH extent for INDEX((range<>0),0)
  ([#507](https://github.com/Teal-Insights/excel-grapher/pull/507),
  [`d18e838`](https://github.com/Teal-Insights/excel-grapher/commit/d18e8389cd7956f9824261dc08aa6830db24d1a3))

- **grapher**: MATCH extent for INDEX((range<>0),0) first-nonzero pattern (#506)
  ([#507](https://github.com/Teal-Insights/excel-grapher/pull/507),
  [`d18e838`](https://github.com/Teal-Insights/excel-grapher/commit/d18e8389cd7956f9824261dc08aa6830db24d1a3))

- **tests**: Use constant Literal forms in #506 MCVE
  ([#507](https://github.com/Teal-Insights/excel-grapher/pull/507),
  [`d18e838`](https://github.com/Teal-Insights/excel-grapher/commit/d18e8389cd7956f9824261dc08aa6830db24d1a3))


## v5.1.0 (2026-08-07)

### Features

- **grapher**: Element-aware guards for array-context IF (#483)
  ([#495](https://github.com/Teal-Insights/excel-grapher/pull/495),
  [`7b0eea6`](https://github.com/Teal-Insights/excel-grapher/commit/7b0eea644321f6446d91e1157c39fcd9770b66b8))

### Performance Improvements

- **core**: Hand ndarray operands straight to the array operator paths
  ([#505](https://github.com/Teal-Insights/excel-grapher/pull/505),
  [`6f689eb`](https://github.com/Teal-Insights/excel-grapher/commit/6f689eb7cecfe87a0f3c947a0dfd5e21887b5eca))


## v5.0.0 (2026-08-07)

### Features

- **grapher**: Drop direct_sites_formula, make raw formula storage opt-in
  ([#500](https://github.com/Teal-Insights/excel-grapher/pull/500),
  [`8a03de3`](https://github.com/Teal-Insights/excel-grapher/commit/8a03de31e4343ee1687c8439ec138c6c925b2bd4))

- **scripts**: Add a repeatable graph memory measurement harness (#490)
  ([#501](https://github.com/Teal-Insights/excel-grapher/pull/501),
  [`35ebe7a`](https://github.com/Teal-Insights/excel-grapher/commit/35ebe7a71a259b93c80477e34cdf999cd675b733))

### Performance Improvements

- **grapher**: Share one empty metadata mapping across nodes (fixes #493)
  ([#499](https://github.com/Teal-Insights/excel-grapher/pull/499),
  [`27287d0`](https://github.com/Teal-Insights/excel-grapher/commit/27287d0f6cd88cbb877e05dbe73ec6a27bd74a1b))

### Breaking Changes

- **grapher**: `EdgeProvenance.direct_sites_formula` is removed, and `FormulaRewrite` no longer
  carries `before_formula` / `after_formula` -- the compression audit trail reports normalized
  formulas only. `Node.formula` is now `None` unless the graph is built with
  `create_dependency_graph(..., store_raw_formula=True)`, which TACO range compression requires.
  Graph caches written by schema 3 are rejected and rebuilt.


## v4.1.2 (2026-08-07)

### Bug Fixes

- **grapher**: Resolve INDEX empty-arg vectors for MATCH (fixes #497)
  ([#498](https://github.com/Teal-Insights/excel-grapher/pull/498),
  [`e892363`](https://github.com/Teal-Insights/excel-grapher/commit/e8923639ffd9c861372c96009474169eac9c6aed))

### Documentation

- **tests**: Drop stale xfail note from extraction_basics docstring
  ([#494](https://github.com/Teal-Insights/excel-grapher/pull/494),
  [`9073405`](https://github.com/Teal-Insights/excel-grapher/commit/9073405189a04d0a8f80b95fabe71116c65a16ec))

### Refactoring

- **grapher**: Share collapse/densify skeleton across numeric domains
  ([#489](https://github.com/Teal-Insights/excel-grapher/pull/489),
  [`b016196`](https://github.com/Teal-Insights/excel-grapher/commit/b016196b185909b27c0ab8a887bd8f8cfd03cd1a))

### Testing

- Make xfail markers strict by default
  ([#494](https://github.com/Teal-Insights/excel-grapher/pull/494),
  [`9073405`](https://github.com/Teal-Insights/excel-grapher/commit/9073405189a04d0a8f80b95fabe71116c65a16ec))


## v4.1.1 (2026-08-07)

### Bug Fixes

- **grapher**: Avoid ZeroDivisionError in _div_numeric_domains
  ([#488](https://github.com/Teal-Insights/excel-grapher/pull/488),
  [`194ab3d`](https://github.com/Teal-Insights/excel-grapher/commit/194ab3d3eade40170af88cf8011372dd16db3d0f))

- **grapher**: Isolate memoized candidate static-ref sets
  ([#485](https://github.com/Teal-Insights/excel-grapher/pull/485),
  [`3f68255`](https://github.com/Teal-Insights/excel-grapher/commit/3f6825525be6e9d29668df935db5f7d9b2d6e27f))

### Performance Improvements

- **grapher**: Memoize candidate static-ref walks (fixes #484)
  ([#485](https://github.com/Teal-Insights/excel-grapher/pull/485),
  [`3f68255`](https://github.com/Teal-Insights/excel-grapher/commit/3f6825525be6e9d29668df935db5f7d9b2d6e27f))


## v4.1.0 (2026-08-07)

### Bug Fixes

- **grapher**: Allow optimal inline of bodies with guarded outs
  ([#481](https://github.com/Teal-Insights/excel-grapher/pull/481),
  [`40361bb`](https://github.com/Teal-Insights/excel-grapher/commit/40361bb9667e99e74e392f1e7a8579e005f3f42d))

### Features

- **grapher**: Extract guards from conditionals embedded in expressions
  ([#481](https://github.com/Teal-Insights/excel-grapher/pull/481),
  [`40361bb`](https://github.com/Teal-Insights/excel-grapher/commit/40361bb9667e99e74e392f1e7a8579e005f3f42d))

### Refactoring

- **viz**: Share SCC condensation ranking between grapher and exporter
  ([#482](https://github.com/Teal-Insights/excel-grapher/pull/482),
  [`75cc154`](https://github.com/Teal-Insights/excel-grapher/commit/75cc1544489f689eeab484d67b0a41d01557fe7e))


## v4.0.2 (2026-08-07)

### Bug Fixes

- **grapher**: Use dict LRU so AddressKey and str share cache entries
  ([#479](https://github.com/Teal-Insights/excel-grapher/pull/479),
  [`12f2ff4`](https://github.com/Teal-Insights/excel-grapher/commit/12f2ff46f155a7d850e1e079372d86649cfe26f5))

### Code Style

- Format edge-provenance changes and silence ty on rejection test
  ([#480](https://github.com/Teal-Insights/excel-grapher/pull/480),
  [`6145a05`](https://github.com/Teal-Insights/excel-grapher/commit/6145a05122844c7e8f76ad023a7b20654ae4d267))

- Reformat after rebase onto IntFlag main
  ([#480](https://github.com/Teal-Insights/excel-grapher/pull/480),
  [`6145a05`](https://github.com/Teal-Insights/excel-grapher/commit/6145a05122844c7e8f76ad023a7b20654ae4d267))

### Performance Improvements

- **grapher**: Flatten _edge_extra into typed _edge_provenance
  ([#480](https://github.com/Teal-Insights/excel-grapher/pull/480),
  [`6145a05`](https://github.com/Teal-Insights/excel-grapher/commit/6145a05122844c7e8f76ad023a7b20654ae4d267))

- **grapher**: Flatten _edge_extra into typed _edge_provenance (#474)
  ([#480](https://github.com/Teal-Insights/excel-grapher/pull/480),
  [`6145a05`](https://github.com/Teal-Insights/excel-grapher/commit/6145a05122844c7e8f76ad023a7b20654ae4d267))

- **grapher**: Slots=True on Node with address-keyed derived-field LRU (#476)
  ([#479](https://github.com/Teal-Insights/excel-grapher/pull/479),
  [`12f2ff4`](https://github.com/Teal-Insights/excel-grapher/commit/12f2ff46f155a7d850e1e079372d86649cfe26f5))

- **grapher**: Use slots=True on Node with address-keyed LRU
  ([#479](https://github.com/Teal-Insights/excel-grapher/pull/479),
  [`12f2ff4`](https://github.com/Teal-Insights/excel-grapher/commit/12f2ff46f155a7d850e1e079372d86649cfe26f5))

### Testing

- **grapher**: Avoid ruff B010 in Node slots attribute check
  ([#479](https://github.com/Teal-Insights/excel-grapher/pull/479),
  [`12f2ff4`](https://github.com/Teal-Insights/excel-grapher/commit/12f2ff46f155a7d850e1e079372d86649cfe26f5))


## v4.0.1 (2026-08-07)

### Performance Improvements

- **core**: Slot CellKey, RangeKey, and UnionKey
  ([#478](https://github.com/Teal-Insights/excel-grapher/pull/478),
  [`47e073f`](https://github.com/Teal-Insights/excel-grapher/commit/47e073f1fbb6986dd7e11a1a9f862de3fa3912b1))


## v4.0.0 (2026-08-07)

### Features

- Store DependencyCause as IntFlag bitmask
  ([#477](https://github.com/Teal-Insights/excel-grapher/pull/477),
  [`9672528`](https://github.com/Teal-Insights/excel-grapher/commit/96725282e3c75b3996c946fa877bce9bfeb78567))


## v3.21.2 (2026-08-07)

### Performance Improvements

- **dynamic-refs**: Serve cached refs in bulk during env expansion (#465)
  ([#471](https://github.com/Teal-Insights/excel-grapher/pull/471),
  [`50fb962`](https://github.com/Teal-Insights/excel-grapher/commit/50fb962a8c5f91cebd630a2c7e9b3796d9c59c1b))


## v3.21.1 (2026-08-07)

### Bug Fixes

- **dynamic-refs**: Gate consumed-leaf bookkeeping on persistent cache (#463)
  ([#464](https://github.com/Teal-Insights/excel-grapher/pull/464),
  [`fa5722c`](https://github.com/Teal-Insights/excel-grapher/commit/fa5722c944dcbf282da15be2ba3f03eb77f4e43d))

- **grapher**: Conjoin nested conditional guards into AND edge guards (#115)
  ([#470](https://github.com/Teal-Insights/excel-grapher/pull/470),
  [`e6e7607`](https://github.com/Teal-Insights/excel-grapher/commit/e6e760769f6843b6d224e8edc05d403e94f4bbc5))

### Chores

- Clean up stale LIC DSF example files
  ([#468](https://github.com/Teal-Insights/excel-grapher/pull/468),
  [`0937bd8`](https://github.com/Teal-Insights/excel-grapher/commit/0937bd8df8b04782b5b6fe0a400b826802953283))

- **deps**: Upgrade ruff to 0.16 and format Markdown code blocks
  ([#466](https://github.com/Teal-Insights/excel-grapher/pull/466),
  [`ce029f3`](https://github.com/Teal-Insights/excel-grapher/commit/ce029f3cb37ecaffff3605d1f04d496557bb1563))


## v3.21.0 (2026-07-27)

### Features

- **exporter**: Omit compute_all when output bindings cover targets
  ([#461](https://github.com/Teal-Insights/excel-grapher/pull/461),
  [`3b78ee8`](https://github.com/Teal-Insights/excel-grapher/commit/3b78ee83f11b7cce3645f3d38125d4c1efe5bc82))

- **series-bindings**: Enforce input.domain on generated setters
  ([#462](https://github.com/Teal-Insights/excel-grapher/pull/462),
  [`c7c6d0a`](https://github.com/Teal-Insights/excel-grapher/commit/c7c6d0ac1e978015a2d53a7dc31a4c1498b11de9))

- **series-bindings**: Signal when read_*_range is omitted (#459)
  ([#460](https://github.com/Teal-Insights/excel-grapher/pull/460),
  [`8939626`](https://github.com/Teal-Insights/excel-grapher/commit/8939626bcc3057635a18fd71c230de90aae810d4))

- **series-bindings**: Warn when read_*_range omitted for non-contiguous selection
  ([#460](https://github.com/Teal-Insights/excel-grapher/pull/460),
  [`8939626`](https://github.com/Teal-Insights/excel-grapher/commit/8939626bcc3057635a18fd71c230de90aae810d4))

### Testing

- **series-bindings**: Expect omitted range warning for grouped matrix
  ([#460](https://github.com/Teal-Insights/excel-grapher/pull/460),
  [`8939626`](https://github.com/Teal-Insights/excel-grapher/commit/8939626bcc3057635a18fd71c230de90aae810d4))


## v3.20.2 (2026-07-27)

### Bug Fixes

- **series-bindings**: Honour exclude_rows/columns in read_*_range
  ([#456](https://github.com/Teal-Insights/excel-grapher/pull/456),
  [`98e3e09`](https://github.com/Teal-Insights/excel-grapher/commit/98e3e0904b9cbb3bd0d9a6ab86549377521d8406))

- **series-bindings**: Honour exclude_rows/columns in read_*_range (#453)
  ([#456](https://github.com/Teal-Insights/excel-grapher/pull/456),
  [`98e3e09`](https://github.com/Teal-Insights/excel-grapher/commit/98e3e0904b9cbb3bd0d9a6ab86549377521d8406))

### Code Style

- **series-bindings**: Format exclude-aware range reader tests
  ([#456](https://github.com/Teal-Insights/excel-grapher/pull/456),
  [`98e3e09`](https://github.com/Teal-Insights/excel-grapher/commit/98e3e0904b9cbb3bd0d9a6ab86549377521d8406))


## v3.20.1 (2026-07-27)

### Bug Fixes

- **exporter**: Re-export public read_* helpers from generated api.py (#454)
  ([#455](https://github.com/Teal-Insights/excel-grapher/pull/455),
  [`7a1af13`](https://github.com/Teal-Insights/excel-grapher/commit/7a1af13cb38dca375925194bd105291e20ed1788))

- **exporter**: Re-export read_* helpers from generated api.py
  ([#455](https://github.com/Teal-Insights/excel-grapher/pull/455),
  [`7a1af13`](https://github.com/Teal-Insights/excel-grapher/commit/7a1af13cb38dca375925194bd105291e20ed1788))

### Code Style

- **exporter**: Format api readers import wrapping
  ([#455](https://github.com/Teal-Insights/excel-grapher/pull/455),
  [`7a1af13`](https://github.com/Teal-Insights/excel-grapher/commit/7a1af13cb38dca375925194bd105291e20ed1788))


## v3.20.0 (2026-07-22)

### Features

- **series-bindings**: Add series-level exclude_columns
  ([#450](https://github.com/Teal-Insights/excel-grapher/pull/450),
  [`b495238`](https://github.com/Teal-Insights/excel-grapher/commit/b495238b1a1c657c8e62f266588320a3991d0c8e))


## v3.19.1 (2026-07-22)

### Bug Fixes

- **exporter**: Keep public cells out of OptimalCompression
  ([#449](https://github.com/Teal-Insights/excel-grapher/pull/449),
  [`09738ec`](https://github.com/Teal-Insights/excel-grapher/commit/09738ece9ef8ffa203a74f9d7e2dd4be425ad177))

- **exporter**: Keep public/series-bound cells out of OptimalCompression
  ([#449](https://github.com/Teal-Insights/excel-grapher/pull/449),
  [`09738ec`](https://github.com/Teal-Insights/excel-grapher/commit/09738ece9ef8ffa203a74f9d7e2dd4be425ad177))

### Documentation

- Clarify OptimalCompression preserve contract
  ([#449](https://github.com/Teal-Insights/excel-grapher/pull/449),
  [`09738ec`](https://github.com/Teal-Insights/excel-grapher/commit/09738ece9ef8ffa203a74f9d7e2dd4be425ad177))


## v3.19.0 (2026-07-21)

### Features

- **series-bindings**: Add constant direction for reader-only graph leaves
  ([#442](https://github.com/Teal-Insights/excel-grapher/pull/442),
  [`f38fc10`](https://github.com/Teal-Insights/excel-grapher/commit/f38fc104bc78dbdbdacd102fdee43e843362b1ff))


## v3.18.0 (2026-07-20)

### Code Style

- Format test_versions schema version set
  ([#441](https://github.com/Teal-Insights/excel-grapher/pull/441),
  [`7780a20`](https://github.com/Teal-Insights/excel-grapher/commit/7780a20b78a6946c2b1fa234328fb3065c415287))

### Features

- **series-bindings**: Call compute helpers by dims; move output leaves
  ([#441](https://github.com/Teal-Insights/excel-grapher/pull/441),
  [`7780a20`](https://github.com/Teal-Insights/excel-grapher/commit/7780a20b78a6946c2b1fa234328fb3065c415287))

- **series-bindings**: Call compute helpers by dims; move output leaves (#435)
  ([#441](https://github.com/Teal-Insights/excel-grapher/pull/441),
  [`7780a20`](https://github.com/Teal-Insights/excel-grapher/commit/7780a20b78a6946c2b1fa234328fb3065c415287))


## v3.17.1 (2026-07-20)

### Bug Fixes

- **series-bindings**: Align reader Google/Numpy docstrings with signatures
  ([#440](https://github.com/Teal-Insights/excel-grapher/pull/440),
  [`768cda6`](https://github.com/Teal-Insights/excel-grapher/commit/768cda6920cc87361f0e8343ada9b3303def6774))


## v3.17.0 (2026-07-20)

### Features

- **series-bindings**: Soft-capture XlErrorException in compute_* measures
  ([#438](https://github.com/Teal-Insights/excel-grapher/pull/438),
  [`7e713e2`](https://github.com/Teal-Insights/excel-grapher/commit/7e713e2a62914d9cac148784b6deaaf2799f5b09))

### Testing

- Lock Excel trailing-space compare semantics (#434)
  ([#437](https://github.com/Teal-Insights/excel-grapher/pull/437),
  [`5a769ca`](https://github.com/Teal-Insights/excel-grapher/commit/5a769ca1f137bcd8fbc3ae642c4716ecf0c6fb48))


## v3.16.0 (2026-07-19)

### Bug Fixes

- **runtime**: Emit HelperCacheKey with EvalContextBase
  ([#430](https://github.com/Teal-Insights/excel-grapher/pull/430),
  [`a9746f7`](https://github.com/Teal-Insights/excel-grapher/commit/a9746f71b939eb905146acb5559225c86f5ca833))

### Features

- **runtime**: Memoize parameterized helpers via xl_helper
  ([#430](https://github.com/Teal-Insights/excel-grapher/pull/430),
  [`a9746f7`](https://github.com/Teal-Insights/excel-grapher/commit/a9746f71b939eb905146acb5559225c86f5ca833))

- **runtime**: Memoize parameterized helpers via xl_helper / xl_memoize
  ([#430](https://github.com/Teal-Insights/excel-grapher/pull/430),
  [`a9746f7`](https://github.com/Teal-Insights/excel-grapher/commit/a9746f71b939eb905146acb5559225c86f5ca833))


## v3.15.3 (2026-07-19)

### Bug Fixes

- **evaluator**: Lower empty IF branches to 0 instead of None
  ([#433](https://github.com/Teal-Insights/excel-grapher/pull/433),
  [`94d0fff`](https://github.com/Teal-Insights/excel-grapher/commit/94d0fff720b60e61a6f533e99c00b7c192e13a97))


## v3.15.2 (2026-07-18)

### Bug Fixes

- **ci**: Avoid orphan release tags when main push races
  ([#429](https://github.com/Teal-Insights/excel-grapher/pull/429),
  [`34b4d0b`](https://github.com/Teal-Insights/excel-grapher/commit/34b4d0b9f994ae2c92d040bc25e46cbc7111dee5))

- **coercions**: Reject empty text in numeric coercion
  ([#426](https://github.com/Teal-Insights/excel-grapher/pull/426),
  [`e009de4`](https://github.com/Teal-Insights/excel-grapher/commit/e009de42c2c5f621eb36b7f0386d9113177ef899))

- **coercions**: Reject empty text in numeric coercion (#420)
  ([#426](https://github.com/Teal-Insights/excel-grapher/pull/426),
  [`e009de4`](https://github.com/Teal-Insights/excel-grapher/commit/e009de42c2c5f621eb36b7f0386d9113177ef899))

- **evaluator**: Ignore blanks/text/bools in range aggregates
  ([#425](https://github.com/Teal-Insights/excel-grapher/pull/425),
  [`199bd60`](https://github.com/Teal-Insights/excel-grapher/commit/199bd608dfb804edf1a944f3a404bb980bee967f))

- **evaluator**: Ignore blanks/text/bools in range aggregates (#419)
  ([#425](https://github.com/Teal-Insights/excel-grapher/pull/425),
  [`199bd60`](https://github.com/Teal-Insights/excel-grapher/commit/199bd608dfb804edf1a944f3a404bb980bee967f))

- **evaluator**: Scalarize 1x1 binary operands (INDEX gates)
  ([#424](https://github.com/Teal-Insights/excel-grapher/pull/424),
  [`83d1cb2`](https://github.com/Teal-Insights/excel-grapher/commit/83d1cb22b3bb70aaf6a7fb29b5b1c771df08a3ba))

- **evaluator**: Scalarize 1x1 ranges in binary operands
  ([#424](https://github.com/Teal-Insights/excel-grapher/pull/424),
  [`83d1cb2`](https://github.com/Teal-Insights/excel-grapher/commit/83d1cb22b3bb70aaf6a7fb29b5b1c771df08a3ba))

- **evaluator**: Stop propagating errors through ISNUMBER/ISTEXT
  ([#423](https://github.com/Teal-Insights/excel-grapher/pull/423),
  [`d5ce6c4`](https://github.com/Teal-Insights/excel-grapher/commit/d5ce6c4819f7f61145db2f24483c4df77f8c6ef6))

- **exporter**: Emit 1x1 ranges as scalar cell reads
  ([#424](https://github.com/Teal-Insights/excel-grapher/pull/424),
  [`83d1cb2`](https://github.com/Teal-Insights/excel-grapher/commit/83d1cb22b3bb70aaf6a7fb29b5b1c771df08a3ba))

- **math**: Skip non-numeric text in aggregates
  ([#426](https://github.com/Teal-Insights/excel-grapher/pull/426),
  [`e009de4`](https://github.com/Teal-Insights/excel-grapher/commit/e009de42c2c5f621eb36b7f0386d9113177ef899))

- **series-bindings**: Emit full-arity positional docstring examples
  ([#427](https://github.com/Teal-Insights/excel-grapher/pull/427),
  [`ef8fdc7`](https://github.com/Teal-Insights/excel-grapher/commit/ef8fdc722d4d2088c6845507df467452eed89b01))

### Code Style

- Keep coercion docstring one-line for export baseline
  ([#426](https://github.com/Teal-Insights/excel-grapher/pull/426),
  [`e009de4`](https://github.com/Teal-Insights/excel-grapher/commit/e009de42c2c5f621eb36b7f0386d9113177ef899))

- **tests**: Clarify range-aggregate MCVE helper docstring
  ([#425](https://github.com/Teal-Insights/excel-grapher/pull/425),
  [`199bd60`](https://github.com/Teal-Insights/excel-grapher/commit/199bd608dfb804edf1a944f3a404bb980bee967f))

- **tests**: Drop escaped quotes in #421 docstrings
  ([#424](https://github.com/Teal-Insights/excel-grapher/pull/424),
  [`83d1cb2`](https://github.com/Teal-Insights/excel-grapher/commit/83d1cb22b3bb70aaf6a7fb29b5b1c771df08a3ba))

- **tests**: Ruff-format golden parity SUM comment line
  ([#425](https://github.com/Teal-Insights/excel-grapher/pull/425),
  [`199bd60`](https://github.com/Teal-Insights/excel-grapher/commit/199bd608dfb804edf1a944f3a404bb980bee967f))

### Testing

- **exporter**: Align golden SUM range expectation with Excel
  ([#425](https://github.com/Teal-Insights/excel-grapher/pull/425),
  [`199bd60`](https://github.com/Teal-Insights/excel-grapher/commit/199bd608dfb804edf1a944f3a404bb980bee967f))


## v3.15.1 (2026-07-16)

### Bug Fixes

- **series-bindings**: Share workbook reader across validate/resolve
  ([#417](https://github.com/Teal-Insights/excel-grapher/pull/417),
  [`6ea8062`](https://github.com/Teal-Insights/excel-grapher/commit/6ea806238e788425b21062c36692ad6490c8a0ca))

- **series-bindings**: Share workbook reader across validate/resolve (#416)
  ([#417](https://github.com/Teal-Insights/excel-grapher/pull/417),
  [`6ea8062`](https://github.com/Teal-Insights/excel-grapher/commit/6ea806238e788425b21062c36692ad6490c8a0ca))

### Testing

- **series-bindings**: Include schema_version in load-count fixtures
  ([#417](https://github.com/Teal-Insights/excel-grapher/pull/417),
  [`6ea8062`](https://github.com/Teal-Insights/excel-grapher/commit/6ea806238e788425b21062c36692ad6490c8a0ca))


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
