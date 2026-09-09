# Readable public named-axis computation under 10 MB

Date: 2026-09-09. Branch: `migration/named-axis-codegen`.
Status: proposed work; implementation and feasibility gate remain open.

## Objective and decision

Make generated public named-axis code the executable, inspectable model for the
LIC DSF workbook in `sandbox/lic-dsf/`. Readability and interpretability are the
primary objective. Typed interfaces, storage efficiency, bundle size, and runtime
performance support that objective.

Acceptance requires a complete standalone distribution **under 10,000,000 bytes**
of uncompressed source and required model assets, with behavioral parity intact.
The threshold is not permission to hide calculations in private kernels, encoded
execution plans, compressed source, or a generic formula interpreter. Whether the
target can be achieved under these constraints is an open engineering question.

This plan supersedes the performance acceptance decision and permission for
private scheduled calculation kernels in [the migration report](named-axis-migration.md).
That report's historical measurements and numerical acceptance remain evidence
for its specific artifact; they do not establish acceptance of this architecture.

## Baseline and diagnosis

The [recorded benchmark](named-axis-performance.json) compares the same graph
projection: 7,206,684 bytes for flat output and 32,494,289 bytes for named output.
It reports fresh-import times of 1.878 s and 5.244 s, and peak process working sets
of 379.0 MB and 853.8 MB. These are historical single-process samples, not a
breakdown of retained tensor memory or a whole-model performance comparison.

Inspection of the saved `benchmark_named_fx_20260909` package found:

| Component | Source bytes | Finding |
| --- | ---: | --- |
| `data.py` | 11,829,252 | Coordinate-to-cell maps account for 8,155,376 bytes across 130,423 entries and 1,664 maps. |
| `internals.py` | 8,974,556 | Named helpers are not imported by the generated public entry points. |
| `_kernel.py`, `_kernels.py`, `_data.py` | 9,034,079 | Public wrappers still execute the private flat implementation. |
| `api.py` | 2,605,777 | Lines containing `CoordinateBuffer` occupy about 1.24 MB; schema-validation statements occupy about 0.65 MB. |
| Other support | 50,625 | Small relative to generated model code and metadata. |

These categories describe the existing artifact, not additive promises of future
savings. In particular, deleting unused named helpers would optimize the wrong
execution architecture. They must become the real calculation path, with gaps
implemented and the superseded flat path removed.

Relevant implementation points:

- [Named emission](../excel_grapher/exporter/inverted_tree/named_emit.py) emits both
  named helpers and wrappers around retained flat modules. `_semantic_body`
  rejects some dynamic references, and fused helpers can fall back as a unit.
- [Expression emission](../excel_grapher/exporter/inverted_tree/ast_emit.py), in
  `_positional_table_source`, bypasses compact slice/zip lowering for named
  coordinates and emits per-cell lazy callbacks.
- [Tensor storage](../excel_grapher/exporter/export_runtime/tensor.py) already has
  ordered axes and tuple-backed values, but also builds a coordinate-to-value
  dictionary per tensor. Record construction adds temporary allocations.

Named dimensions do not inherently remove adjacency: declared axis order supplies
it. Worksheet order, semantic domain order, and dependency evaluation order are
distinct and must remain explicit. Regularity should be recognized at generation
time and expressed as meaningful loops and selections.

## Architecture and readability contract

- Public entry points call inspectable named model functions containing the actual
  formulas, branches, loops, and recurrence relationships. Public orchestration
  may call other named functions; each output need not duplicate its closure.
- Formula families use semantic variables such as scenario, year, and vintage.
  Seeds, horizon boundaries, and exceptional formulas remain visible.
- Shared runtime primitives may implement Excel operators, lazy selection, tensor
  access, and dependency bookkeeping. Workbook-specific formulas and scheduling
  relationships must remain visible in generated model source.
- Small helpers should name recognizable operations. Moving formula bodies into
  underscore-prefixed modules, opaque dispatch tables, bytecode, or a generic
  execution engine does not satisfy public computation merely by renaming files.
- Preserve descriptive identifiers, type annotations, provenance access, and
  useful docstrings. Minification and removal of explanatory structure are not
  size strategies.
- Named access may use contiguous physical storage and shared indexes. Storage
  layout is an implementation detail; public formula code must not reconstruct
  dimensions through anonymous slots, breakpoints, or instance partitions.
- Irregular domains and genuinely different formulas may require explicit entries
  and branches. Do not force compression through an incorrect regularity claim.

## Work sequence

Use test-driven development for every implementation change: add the smallest
meaningful regression, observe it fail for the intended reason, implement the
change, and refactor while green. Record RED/GREEN evidence with each work item.
The following phases are ordered; measurements determine priorities within them.

### 1. Establish reproducible measurements and execution coverage

- [ ] Add a reproducible export/measurement command using the sandbox workbook,
  bindings, constraints, targets, and graph projection. Record the exact command,
  repository revision, input hashes, environment, and generated artifact hashes.
  Do not depend on the historical temporary downstream checkout surviving.
- [ ] Report bytes per module and category, formula/helper counts, explicit
  coordinate entries, range callbacks, and duplicated formula bodies. Include all
  required assets even if metadata moves out of Python. Exclude build caches and
  bytecode from the source gate, but report their sizes separately if relevant.
- [ ] Measure generation, fresh source-only import, warm import, retained imported
  state, and representative/full evaluation separately. Use repeated fresh
  processes and report sample counts and variation. Profile Python allocations
  separately from uninstrumented timing and process-memory measurements.
- [ ] Trace the actual call/import graph from public outputs. Inventory every
  fallback and classify the missing named lowering, especially OFFSET/INDIRECT,
  fused recurrences, partial windows, and lookup ranges.
- [ ] Add a failing public-computation contract test: an exported package must run
  without the private flat modules and must execute its generated named formula
  bodies. Include an execution witness, not just a forbidden-filename check.
- [ ] Establish readable source exemplars for a regular year series, sparse
  scenario/year domain, vintage recurrence, lookup, and fused dependency group.
  Review their actual calculation path as well as their entry-point signatures.

Deliverable: reproducible baseline report, fallback inventory, and architecture
tests that expose the present wrapper/kernel dependency.

### 2. Make named public computation complete

- [ ] Route public functions through named calculation functions and explicit
  named orchestration. Preserve sharing of intermediate results where needed.
- [ ] Implement each missing lowering from the inventory with focused parity
  tests. Keep dynamic selection, dependencies, seeds, and boundary behavior
  inspectable in the generated source.
- [ ] Preserve lazy branches and deferred range access: unselected errors and
  unevaluated recurrence members must not become eagerly evaluated.
- [ ] Handle fused dependencies with readable named computations and explicit
  dependency relationships. General bookkeeping may be shared; the model's
  formulas must not move into a hidden evaluator.
- [ ] Remove flat kernel emission, adapters on the computation path, and duplicated
  formula implementations once the named path passes the corresponding tests.
- [ ] Run the full LIC DSF differential on this architecture before attributing
  later gains to optimization of a correct replacement.

Deliverable: complete public named execution, with no fallback to private flat
calculation. Record its size even if it exceeds the target substantially.

### 3. Compact formulas through readable iteration and views

- [ ] Recognize contiguous axis runs, fixed-coordinate selections, affine shifts,
  rectangular selections, and repeated formula families during generation.
- [ ] Emit semantic loops and explicit seed/boundary branches. Preserve declared
  categorical order, reversed scans, irregular years, and unequal horizons; do
  not equate an axis step with integer addition unless proven valid.
- [ ] Replace per-cell range lambdas with readable lazy range/view primitives that
  preserve Excel range ordering, blank geometry, and deferred errors.
- [ ] Share common named subcomputations without eagerly evaluating unused values
  or materializing full domains when only a window is required.
- [ ] Add scaling tests: lengthening a regular horizon should grow necessary data
  but not linearly replicate formula bodies or callbacks. Test sparse exceptions
  separately instead of imposing a constant-size claim on arbitrary workbooks.

Deliverable: before/after source exemplars, size attribution, and passing range,
ordering, recurrence, and error-semantics tests.

### 4. Compact supporting metadata and public plumbing

- [ ] Represent regular coordinate-to-cell provenance with worksheet range and
  axis mapping descriptors. Retain exact explicit exceptions for irregular cells.
  Keep provenance lookup and iteration available without eager full expansion.
- [ ] Share identical axes, domains, and membership indexes; represent product
  domains and regular required subsets structurally. Preserve authored versus
  graph-required coordinates and authored versus canonical ordering.
- [ ] Remove repeated defaults and redundant coordinate literals. Preserve blanks,
  missing coordinates, zeros, and error observations as distinct cases.
- [ ] Consolidate validation where immutable tensors and schema identity make it
  sound. Maintain public-boundary checks, constant override validation, and useful
  coordinate-specific errors. Do not replace visible formulas with execution plans
  to reduce repeated call text.
- [ ] Verify exact provenance equivalence over the LIC DSF corpus and round trips
  for serialization, schemas, defaults, and partial graph projections.

Deliverable: smaller importable metadata with unchanged observable meaning, and
an updated package-size report counting every required file.

### 5. Reduce tensor and construction memory

- [ ] Use axis indexes and strides for product-domain lookup; share a
  coordinate-to-position index for sparse domains instead of duplicating a
  coordinate-to-value dictionary for every tensor.
- [ ] Reduce record-construction copies and repeated validation with trusted
  internal construction paths that preserve public validation guarantees.
- [ ] Use views for appropriate selections and avoid needless materialization.
  Check ownership and lifetime behavior so a small result does not inadvertently
  retain a large calculation graph or unrelated buffers.
- [ ] Remeasure retained memory, import peak, and evaluation peak separately.
  Source-compilation savings and runtime-storage savings are different outcomes.

Deliverable: attributed memory changes and timing results without a regression in
immutability, indexing, iteration order, errors, or serialization.

### 6. Final acceptance and documentation

- [ ] Generate a fresh complete package for the pinned LIC DSF workload. Require
  total uncompressed source plus required assets to be strictly below 10,000,000
  bytes. Do not reduce targets, scenarios, defaults, or provenance to pass.
- [ ] Verify public execution without private flat kernels in a standalone process
  without an installed `excel_grapher` dependency. Check that the formulas being
  reviewed are the formulas actually executed.
- [ ] Repeat the established differential: 64 scenarios, 1,785 targets per
  scenario, 114,240 comparisons, zero mismatches and zero unexpected matched
  errors. Preserve `atol=1e-6`, `rtol=1e-12` and the existing narrow
  [chart error contract](named-axis-chart-error-contract.json). Matching errors
  outside that contract must not silently count as parity.
- [ ] Run live-Excel comparisons where automation is available; otherwise skip
  those tests with a clear reason. Distinguish live-Excel evidence from evaluator
  and cache-based comparisons in the final report.
- [ ] Run repository lint, format, type checks, and tests through `uv run`, plus
  relevant slow tests, Tiny-DSA CLI smoke validation, generated-package checks,
  and downstream integration gates. Use `fastpyxl` for workbook fixtures.
- [ ] Review the source exemplars and representative LIC DSF calculation paths
  for readability. Reject size reductions that obscure formulas or dependencies.
- [ ] Update export documentation, package materialization, module manifests, and
  representation/cache fingerprints as required. For changed `.qmd` documentation,
  format its source with `uv run python scripts/format_qmd.py` and re-render before
  checking generated Markdown.
- [ ] Save final size, memory, timing, provenance, coverage, parity, and source
  review evidence in `plans/`, with reproducible commands and artifact hashes.

## Feasibility checkpoints

After phase 2, quantify the size of a correct public-computation implementation.
After phases 3 and 4, measure the remaining gap against 10 MB and attribute it to
specific formulas, irregular data, or remaining duplication. Reprioritize based
on measured costs rather than the historical split between named and flat code.

The size limit is a hard acceptance gate; memory and timing have no newly agreed
numeric thresholds. Report their measured changes and material regressions rather
than inventing additional gates or treating smaller source as proof of lower peak
memory. Correctness and readability remain mandatory throughout.

If readable public computation still exceeds 10 MB after the concrete structural
opportunities are exhausted, document the remaining cost and the source examples
that explain it. Report the target as unmet. Do not declare the replacement
accepted, hide formulas, relax parity, or silently return to private kernels.
