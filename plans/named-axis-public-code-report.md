# Readable public named-axis computation under 10 MB: evidence

Date: 2026-09-10. Branch: `migration/named-axis-codegen`.
Plan: [named-axis-public-code-optimization.md](named-axis-public-code-optimization.md).
Machine-readable evidence: [named-axis-public-code-evidence.json](named-axis-public-code-evidence.json).

## Result

The complete standalone LIC DSF package generated at revision `4d1ed70` is
**9,710,332 bytes** of uncompressed source across eight modules, under the
10,000,000-byte gate. Every public output is computed by generated named formula
functions; there is no private flat kernel, encoded plan, or generic interpreter.
The package imports and evaluates without `excel_grapher` installed.

| Module | Bytes | Content |
| --- | ---: | --- |
| `__init__.py` | 12,169 | Package exports |
| `api.py` | 2,716,723 | 103 `compute_*` functions, `Model`, input checks, constant sets |
| `internals.py` | 4,502,278 | 1,548 named formula functions and 27 recurrence groups |
| `data.py` | 2,338,303 | Axes, domains, schemas, series classes, defaults, provenance |
| `tensor.py` | 21,723 | Immutable axes, domains, tensors, schemas |
| `provenance.py` | 8,142 | Rectangle and grid cell descriptors |
| `runtime.py` | 34,217 | Views, spans, readers, publication |
| `excel.py` | 76,777 | Embedded Excel value semantics |
| total | 9,710,332 | |

Size history of the same graph projection under this architecture (all bytes):

| Stage | Total | api | internals | data |
| --- | ---: | ---: | ---: | ---: |
| First complete named computation (phase 2) | 27,628,401 | 7,044,878 | 10,742,446 | 9,696,244 |
| Memoized `Model` orchestration | 23,647,565 | 3,128,859 | 10,674,801 | 9,696,244 |
| Grid provenance, product views, folded families, coordinate runs, table views | 10,555,027 | 2,854,361 | 5,181,706 | 2,366,506 |
| Formula bodies as coordinate functions | 10,225,307 | 2,854,361 | 4,851,460 | 2,366,506 |
| Shared constant sets and value types, trusted results, union conditions | 9,676,463 | 2,713,615 | 4,485,125 | 2,324,695 |
| Keyed scalar layouts and sandbox binding retyping (final) | 9,710,332 | 2,716,723 | 4,502,278 | 2,338,303 |

## Reproduction

Inputs live in the gitignored `sandbox/lic-dsf/` directory (see its README).

```bash
uv run python scripts/measure_named_export.py sandbox/lic-dsf \
  --graph sandbox/lic-dsf/graph/6130166b3d5cc4158e80fafaf79b689ea124de4891bd9551c4df6bc7adb1118b.pkl.gz \
  --out-dir build/lic_dsf_named --report build/lic_dsf_named.json \
  --reconcile-bindings --standalone --tracemalloc
uv run python scripts/named_differential.py build/lic_dsf_named \
  --graph sandbox/lic-dsf/graph/6130166b3d5cc4158e80fafaf79b689ea124de4891bd9551c4df6bc7adb1118b.pkl.gz \
  --blank-ranges sandbox/lic-dsf/lic_dsf_blank_ranges.py \
  --contract plans/named-axis-chart-error-contract.json --report build/lic_dsf_differential.json
```

`--reconcile-bindings` applies the sandbox binding adjustments recorded in the
report (239 entries): output bindings without graph formulas are
dropped, overlapping formula bindings resolve to the first binding, scalar
constant bindings are added for unbound closure leaves, text-typed input
measures over numeric cells and integer measures over fractional cells are read
as floats. The adjustments change which cells are bound and how their values are
typed, not how any bound formula is computed.

Input fingerprints (SHA-256):

| Input | Hash |
| --- | --- |
| `workbook.xlsm` | `3a0a0b80c7cbc95ac953f25ecae0b437129d669ceb8aeefb54ab86dc8727ea86` |
| graph projection `6130166b…8b.pkl.gz` | `2e61d0ffc5871beea74df71ab420e022fa9f00b77a803972d0c49dad01872702` |
| 17 binding files, concatenated | `6d0e223d8bfb6d34d1a8087aab80eb5145269b8015028810fd920c3998250160` |
| `lic_dsf_blank_ranges.py` | `b53e036758077ebecd9e3125f3242a11e4ae9085fc674802dc50b5816fc7dbee` |
| `targets.json` | `f715b997c922fec24e060bd15ebde080e91ad820326627d357ea43cf414d51c9` |

Generated module hashes are in the evidence JSON. Environment: Python 3.13.12,
Linux x86_64 container, `uv run`.

## Architecture as generated

- `api.py`: 103 keyword-only `compute_*` functions, each binding exactly the
  input leaves of its output and reading one attribute of `Model`. `Model`
  defines every formula series once as a memoized attribute (1,965 attributes
  including 27 recurrence groups); inputs are validated once per model.
- `internals.py`: 1,548 named functions. Each defines `formula(<coordinates>)`
  over the series' semantic axes and publishes it with `evaluate()`; recurrence
  groups define one coordinate function per member and read them through
  demand-driven `CoordinateReader`s. Ranges are `view`/`span` selections over
  named axes (1,715 views, 1,207 spans); positional lookup tables mix views with
  per-cell callbacks (165 tables).
- `data.py`: shared axes, domains (ragged domains as coordinate runs), required
  subdomains, schemas, concrete `Series` classes, defaults, and provenance
  descriptors (1,433 rectangles, 405 grids, 670 explicit dictionaries).
- `tensor.py`, `provenance.py`, `runtime.py`, `excel.py`: shared primitives and
  the embedded Excel value semantics.

Byte attribution of the final package is recorded in the evidence JSON
(`attribution`). The largest remaining costs are inherent to the workbook:
formula bodies (2.4 MB), public signatures naming every input leaf (1.5 MB), and
the `Model` attribute bodies naming every dependency (0.6 MB).

## Source exemplars

Regular year series with a seed and a growing window (`internals.py`):

```python
def pv_base_output_cumulative(*, pv_base_output_new_forex_borrowing_gross_usd: data.PvBaseOutputNewForexBorrowingGrossUsd) -> data.PvBaseOutputCumulative:
    """Compute `pv_base_output_cumulative` using authored coordinate identities."""
    def formula(instrument: str, time_period: int) -> float | str | None:
        if time_period == 2024:
            return as_measure(pv_base_output_new_forex_borrowing_gross_usd[instrument, time_period])
        return as_measure(xl_sum(view(pv_base_output_new_forex_borrowing_gross_usd, rows=(instrument,), cols=span(data.TIME_PERIOD_AXIS_43, 2024, time_period))))

    return data.PvBaseOutputCumulative.collect(evaluate(formula, data.PV_BASE_OUTPUT_CUMULATIVE_REQUIRED))
```

Label keys derived from the host key and per-instrument families (`pv_base_discount_amortization`):

```python
    def formula(instrument: str, time_period: int) -> float | str | None:
        if instrument == 'IDA - regular' and time_period == 2024:
            return as_measure((0 if xl_lt(..., pv_base_block_terms[instrument, f'Grace {instrument}', 'external']) else ...))
        elif instrument == 'IDA - SML' and 2075 <= time_period <= 2076:
            ...
```

Nested vintage block read as a product view (`input5_new_debt_debt_stock_on_new_debt_denominated_in_local_currency`):

```python
        return as_measure(xl_sum(view(input5_vintage_stock, rows={'HOLDER': (holder,), 'INSTRUMENT': span(data.INSTRUMENT_AXIS_20, 'Central bank financing', 'T-bills (denominated in local currency)'), 'ISSUANCE_YEAR': span(data.ISSUANCE_YEAR_AXIS, time_period - 1, time_period)}, cols={'TIME_PERIOD': (time_period,)}), ...))
```

Public entry point (`api.py`):

```python
@publish(data.PV_BASE_OUTPUT_CUMULATIVE_SCHEMA, constants=_CONSTANTS_4, cells=data.PV_BASE_OUTPUT_CUMULATIVE_CELLS)
def compute_...(*, input4_interest_rate: data.Input4InterestRate, ...) -> data....:
    """Compute `...` using authored coordinate identities."""
    return Model(**locals())....
```

## Parity

`scripts/named_differential.py` evaluates every public output on the workbook
defaults and compares each published cell with `FormulaEvaluator` on the same
graph projection and blank ranges (`atol=1e-6`, `rtol=1e-12`; text and booleans
exact; an Excel error matches only the same code).

| Comparison | Result |
| --- | ---: |
| Public outputs evaluated | 103 of 103 |
| Output cells compared | 1,783 |
| Mismatches | 0 |
| Matched errors, all `#N/A` inside the chart error contract | 105 |
| Matched errors outside the contract | 0 |

`--internals` extends the comparison to every named series and input default
(1,938 series including recurrence-group members): 0 mismatched series. Three
classes of sandbox binding defects surfaced on the way and are corrected by the
recorded reconciliation: scalar layouts with one cell per scenario sheet (now
keyed series), text-typed inputs over numeric cells, and integer-typed measures
over fractional cells (both read as floats).

The 64-scenario downstream differential (114,240 comparisons) requires the
downstream extraction pipeline repository, which is not available in this
environment; it was not repeated here. Live-Excel comparisons were skipped: no
Excel automation is available on this Linux container.

## Timing and memory

Measured with `scripts/measure_named_export.py --package-dir … --standalone
--tracemalloc` (fresh subprocess per probe, `excel_grapher` imports blocked) and
one direct probe, on the final package:

| Measurement | Value |
| --- | ---: |
| Bytecode of the package (`__pycache__`, excluded from the source gate) | 8,665,255 bytes |
| Import with bytecode present, median of 3 fresh processes | 0.27 s |
| Import while tracing allocations (compiles from source) | 28.0 s |
| Peak RSS after import | 370.0 MB |
| Per-output first call (each `compute_*` builds its own `Model`), median | 8.4 s |
| All 103 outputs, one call each | 811 s |
| One `Model` evaluating all 1,938 series once | 10.3 s |
| Peak RSS with every series retained | 434.4 MB |
| Traced Python allocations, import plus every series retained | 256.1 MB |

The historical flat export recorded a 1.878 s fresh import and a 379.0 MB
working set on a different machine; the named package imports in a comparable
footprint. Each public call recomputes its own closure because `Model` memoizes
per instance; callers evaluating many outputs should read them from one `Model`.
Tensor memory is bounded by the shared axis and domain indexes (`Domain.position`)
and tuple-backed values; no coordinate-to-value dictionaries are built per tensor.

## Checks

All run through `uv run` on this branch at the final revision:

| Check | Result |
| --- | --- |
| `pytest` (full suite, default deselection of `slow`) | 3,818 passed, 1 skipped, 54 deselected |
| `ruff check .` | passed |
| `ruff format --check .` | 616 files already formatted |
| `ty check excel_grapher tests examples scripts` | no errors; three pre-existing redundant-cast warnings in `tests/unit/core` |
| Tiny-DSA CLI `bindings validate --smoke-test` | all compute functions passed smoke checks |
| Public-computation contract tests (`tests/integration/exporter/inverted_tree/test_public_computation_contract.py`) | pass, including the standalone subprocess with `excel_grapher` blocked |
| Downstream integration gates | not available in this environment |

Every emission change on this branch was made test-first: the new tests under
`tests/unit/exporter/inverted_tree/` (`test_model_orchestration`,
`test_nested_layouts`, `test_family_folding`, `test_lookup_tables_compact`,
`test_formula_functions`, `test_series_facade_types`, `test_keyed_scalars`,
`test_public_computation_contract`) and `tests/unit/scripts/` were observed
failing before the corresponding emitter change and passing after it.
