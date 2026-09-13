# Named-axis inverted-tree replacement

Branch: `migration/named-axis-codegen`. Status: **replacement implemented; downstream acceptance passed** (2026-09-09).
All changes remain uncommitted and unmerged.

The generated public contract is now `named-axis-v1`: immutable tensors with
named coordinates replace flat sequence arguments and results. Scalars remain
scalars. There is no parallel legacy public mode. Explicit catalog-order adapters
are available at integration boundaries; private scheduled buffers remain compiler
implementation details.

## Implementation

- Ordered typed axes, arbitrary-rank product and explicit sparse domains, exact
  construction, label indexing, `sel`, `isel`, schemas, and versioned serialization.
- Generated per-series facades plus year/scenario conveniences, including unequal
  horizons. Missing coordinates, blanks, zero, and Excel errors remain distinct.
- Authored domains and workbook provenance remain separate from graph-required
  coordinates. Structural blank ranges neither add graph nodes nor create binding
  obligations. Existing semantic binding keys supply dimensions without geometry
  inference or redundant shape configuration.
- Generated semantic coordinate expressions and loops, reusable domains and lookup
  tables, frozen named fused results, and demand-driven coordinate readers.
  Dynamic references retain private scheduled lowering. Public inputs and outputs
  do not expose partitions or slot arithmetic.
- Shared Excel operators preserve coercion, error strings, lazy branches, and
  recurrence/cycle behavior. Regression fixes cover categorical schedule order,
  workbook ROW/COLUMN geometry, sparse references, INDEX vectors, scalar OFFSET,
  shared producer windows, lazy CHOOSE, and structural-blank formula boundaries.
- Representation/domain fingerprints invalidate generated caches. Downstream
  package materialization includes all ten generated modules. Documentation,
  public contract tests, and executable DataFrame examples use immutable tensors.

Dedicated hierarchical-axis views remain deferred under specification section
12.4. Separate named grouping dimensions already represent sparse valid pairs and
arbitrary-rank grouped rows. Arbitrary dynamic formulas may require private
lowering rather than a compact symbolic coordinate expression.

## Validation

Tests were changed or added first and observed failing before their corresponding
implementation fixes.

| Check | Result |
| --- | --- |
| Full repository suite | 3,742 passed; 1 skipped; 54 deselected; 1 expected failure; exit 0 |
| Downstream graph/internal coverage and integration gates | 114 passed; exit 0 |
| Downstream binding contracts after restoring explicit year readers | 37 passed |
| Ruff lint and format | Passed |
| Type checks over source, tests, examples, and scripts | No errors; three existing redundant-cast warnings |
| Tiny-DSA CLI bindings validation and compute smoke tests | Passed |
| Immutable DataFrame examples | Both execute successfully |

The fresh downstream differential **passed all 114,240 comparisons** across
64 authored scenarios and all 1,785 target cells, including constant targets.
There were **zero mismatches, zero unexpected matched errors, and process exit 0**.
An independent CSV audit verified unique scenario/address pairs, exact per-scenario
coverage, and identical target sets. [Acceptance evidence](named-axis-acceptance.json)
records report hashes, final generated module hashes, implementation hashes,
workbook/binding/projection fingerprints, and the successful cache metadata.
The [parity report](named-axis-parity-report.txt) records the original numerical
tolerances: `atol=1e-6`, `rtol=1e-12`.

The run preserves 11,760 matching `#N/A` values from active authored chart-suppression
branches. These are accepted only at the specified cells when separately evaluated
control values select that explicit branch. Errors in the alternative INDEX/MATCH
branch, errors at other coordinates, and other error types still fail. The
[workbook formula audit](named-axis-chart-error-contract.json) preserves all 273
eligible formulas across 13 rows. Four RED cases drove this narrow contract;
all ten branch/coordinate/value cases and 57 related harness/cache tests passed.
No scenario-wide or global matched-error override was enabled.

This acceptance compares a fresh FormulaEvaluator graph evaluation against the
final standalone package. It is not a fresh live-Excel automation comparison.

The downstream workspace is an isolated copy at
`C:/Users/chris/AppData/Local/Temp/named-axis-downstream-20260908`.
The real sibling checkout is unchanged. The 20-file
[downstream patch](downstream-named-axis.patch) passes ordinary `git apply --check`
against it. It updates tensor adapters, cache/materialization contracts, batched
oracle reads, tests, and numeric binding declarations that previously truncated
fractions or converted debt values to text. Mixed-value declarations retain text
sentinels; authored dimension metadata remains unchanged.

The final canonical export completed in 822.7 seconds, including 746.1 seconds
of code generation. A final metadata refinement restored explicit integer year
readers. The successful differential loaded this final canonical export directly.
[Generated-source equivalence evidence](named-axis-generated-equivalence.json)
compares all ten modules: calculation code, domains, schemas, and annotations
are identical after normalizing only the provenance fingerprint assignment and
unordered literal `frozenset` member ordering. This establishes that the performance
measurements apply to the final metadata refinement; the fresh acceptance run
uses the final package fingerprint directly.

## Performance decision

The following measurements use the same workbook graph projection, fresh
source-only package directories, and disabled bytecode writing. Import and process
peak working set are single fresh-process samples. Evaluation medians use 20 warm
calls with identical defaults. [Raw measurements](named-axis-performance.json)
record the source sizes, values, and method.

| Metric | Flat baseline | Named replacement | Ratio |
| --- | ---: | ---: | ---: |
| Generated Python source | 7,206,684 bytes | 32,494,289 bytes | 4.51x |
| Fresh import | 1.878 s | 5.244 s | 2.79x |
| Peak process working set | 379.0 MB | 853.8 MB | 2.25x |
| Commodity applicability, warm median | 0.562 ms | 2.959 ms | 5.26x |
| Natural-disaster applicability, warm median | 0.040 ms | 0.105 ms | 2.64x |

The baseline market-financing function fails with a `TypeError`; its runtime is
not a valid speed comparison. The replacement returns 1.0 with a 109.5 ms warm
median. These samples do not establish whole-model evaluation speed parity.

Engineering decision: accept these measured regressions for this unmerged
replacement branch following complete numerical acceptance. The named
contract and correctness fixes justify proceeding with the replacement, but
startup cost and approximately 475 MB additional peak process memory are material
limitations. This is not performance parity or a production deployment decision.
Shared lookup tables already reduced the earlier named candidate from about
66 MB to 32.5 MB of source. Further size and startup optimization can preserve the
new contract rather than reintroduce flat public APIs.

## Reproduction

Use `uv run` for Python. In this sandbox, use a writable temporary cache and a
fresh pytest base directory:

```powershell
$env:UV_CACHE_DIR = "$env:TEMP/excel-grapher-uv-cache"
uv run --no-sync pytest -p no:cacheprovider --basetemp "$env:TEMP/unique-test-run"
uv run --no-sync ruff check .
uv run --no-sync ruff format --check .
uv run --no-sync ty check excel_grapher tests examples scripts
```

After applying the downstream patch in an isolated checkout and installing this
branch, run extraction/export and the downstream full differential. Binding
changes require refreshed extraction manifests; stale fingerprints must not be
bypassed. The run here uses `src.extraction_pipeline --stop-after-stage export`
and `src.differential_validation.run_post_refactor_differential` with
`no_cache=True`. Progress instrumentation only logs mismatches and unexpected
matched errors and writes per-scenario CSV files; it does not modify values or
comparison policy.
