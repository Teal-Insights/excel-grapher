# Readable public named-axis computation under 10 MB: evidence

Date: 2026-09-10. Branch: `migration/named-axis-codegen`.
Plan: [named-axis-public-code-optimization.md](named-axis-public-code-optimization.md).
Machine-readable evidence: [named-axis-public-code-evidence.json](named-axis-public-code-evidence.json).

## Result

The complete standalone LIC DSF package generated at revision `REVISION` is
**TOTAL_BYTES bytes** of uncompressed source across eight modules, under the
10,000,000-byte gate. Every public output is computed by generated named formula
functions; there is no private flat kernel, encoded plan, or generic interpreter.
The package imports and evaluates without `excel_grapher` installed.

| Module | Bytes | Content |
| --- | ---: | --- |
MODULE_ROWS
| total | TOTAL_BYTES | |

Size history of the same graph projection under this architecture (all bytes):

| Stage | Total | api | internals | data |
| --- | ---: | ---: | ---: | ---: |
| First complete named computation (phase 2) | 27,628,401 | 7,044,878 | 10,742,446 | 9,696,244 |
| Memoized `Model` orchestration | 23,647,565 | 3,128,859 | 10,674,801 | 9,696,244 |
| Grid provenance, product views, folded families, coordinate runs, table views | 10,555,027 | 2,854,361 | 5,181,706 | 2,366,506 |
| Formula bodies as coordinate functions | 10,225,307 | 2,854,361 | 4,851,460 | 2,366,506 |
| Shared constant sets and value types, trusted results, union conditions | TOTAL_BYTES | API_BYTES | INTERNALS_BYTES | DATA_BYTES |

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
report (214 entries: output bindings without graph formulas dropped, overlapping
formula bindings resolved to the first binding, scalar constant bindings added
for unbound closure leaves). The adjustments change which cells are bound, not
how any bound formula is computed.

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
  named axes (1,710 views, 1,203 spans); positional lookup tables mix views with
  per-cell callbacks (165 tables).
- `data.py`: shared axes, domains (ragged domains as coordinate runs), required
  subdomains, schemas, concrete `Series` classes, defaults, and provenance
  descriptors (1,432 rectangles, 393 grids, 683 explicit dictionaries).
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

PARITY_SECTION

The 64-scenario downstream differential (114,240 comparisons) requires the
downstream extraction pipeline repository, which is not available in this
environment; it was not repeated here. Live-Excel comparisons were skipped: no
Excel automation is available on this Linux container.

## Timing and memory

TIMING_SECTION

## Checks

CHECKS_SECTION
