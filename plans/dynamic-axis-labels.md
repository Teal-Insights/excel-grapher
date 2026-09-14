# Dynamic axis labels (formula-derived row/column headers)

Status: **proposal** (2026-09-14). Responds to
[#841](https://github.com/Teal-Insights/excel-grapher/issues/841) and
[lic-dsf-extraction-pipeline#53](https://github.com/Teal-Insights/lic-dsf-extraction-pipeline/issues/53).

## Verdict

The direction is worth doing, but not as #841 describes it. Most of the work
#841 budgets for (reading formulas with `data_only=False`, parsing header ASTs
in the binder, a runtime "formula evaluation layer", rewriting the 273
`time_period == 2024` comparisons to `first_projection_year + 0`) is either
already handled by the existing codegen or is aimed at the wrong thing.

The actual gap is one missing concept: **an axis key is used both as a
position name and as a label value, and the two only coincide at snapshot
time.** Fixing that is a bounded change at the public boundary of the
generated package, not a rewrite of emission.

## What #841 gets wrong

1. **The literal comparisons are positional, not semantic.**
   `if time_period == 2024:` in `internals.py` comes from
   `_family_condition` (`named_emit.py`) and says "the formula shape at the
   first projection column". `D30 = D29` versus `E30 = SUM(D29:E29)` is a
   fact about where formulas sit in the sheet. Changing
   `'Input 1 - Basics'!C18` in Excel does not move formulas, so those
   branches must **not** change. Rewriting them to
   `first_projection_year + k` would be wrong for any header formula that is
   not affine in one input, and pointless when it is. The same holds for
   `time_period - 1` lags (`_integer_driver`), `xl_offset`, range tables, and
   diagonal conditions: all positional.
2. **Header formulas already compile.** A header row bound as its own series
   (tiny-DSA's `engine_year_labels`; downstream `macro_debt_data` =
   `'Macro-Debt_Data'!U4` and its five consumers) is an ordinary formula
   series. Its value is computed from inputs by the existing pipeline. No
   new formula reader, AST extraction, or dependency tracker is needed: the
   dependency graph already has the formula and the edges.
3. **Inputs do not need to move to a relative axis.** Making
   `sel(TIME_PERIOD=k)` ordinal pushes label arithmetic onto every consumer
   and every chart. Labels should stay the public keys; positions stay
   internal.

Where #841 is right: the binder resolves `column_header` / `row_label`
through `_read_cell_value` (graph cached value, else `data_only=True`
reader), so axis keys are snapshot values, `Axis` is frozen, and public
tensors are keyed by those snapshot values. That is exactly the boundary this
plan changes.

## Design

Keep the snapshot label as the **internal, stable position identifier** of
every coordinate (nothing in `internals.py` changes). Add an authored link
from an axis to the **series whose values are its labels**, and relabel at
the public boundary:

- **Outputs**: `compute_*` (and `Model` output attributes) evaluate the label
  series and return the tensor with the axis keys replaced by label values.
- **Inputs keyed on a labelled axis**: callers supply tensors keyed by
  labels. The package evaluates the label series (from the other inputs),
  maps label → snapshot key, and hands the snapshot-keyed tensor to
  `internals`.
- Everything between stays snapshot-keyed, so dispatch, lags, OFFSET, range
  tables, literal tables, provenance, and fingerprints are untouched.

Excel parity of the values is preserved because the label series is the
header cell's own formula, evaluated by the same runtime as every other
cell.

### Binding schema (1.17.0)

Series-level field on the series that owns the header cells:

```yaml
- id: engine_year_labels
  sheet: Baseline - external
  data_range: "'Baseline - external'!M8:AG8"
  internal: {}                  # header cells hold formulas (=N8-1, =U4, =N8+1)
  axis_labels: TIME_PERIOD      # this series' values are the runtime labels of its TIME_PERIOD key
  structure:
    measure: {concept: OBS_VALUE, dtype: int, bind: {kind: data_cell, read: int}}
    dimensions:
      - id: TIME_PERIOD
        concept: TIME_PERIOD
        role: key
        scope: cell
        bind: {kind: column_header, header_row: 8, read: int}
  key: [TIME_PERIOD]
```

Rules (all enforced in `series_bindings/validate.py`, error issues):

| Rule | Why |
| --- | --- |
| `axis_labels` names one of the series' own key fields, and it is the only key field. | Labels are a function of position on that axis alone. |
| Measure dtype matches the axis key type (`int` or `string`). | `Axis` requires typed keys. |
| At most one labeller per axis identity (dimension id + snapshot key set); a second series with the same `axis_labels` and identical keys is an error. Axes whose snapshot keys are a subset of the labeller's keys are relabelled through it; keys outside it are an export error naming the missing keys. | Named axes are shared by `(name, keys, type)`; history columns typed as literals plus projection columns as formulas is one row and one labeller. |
| Export-time sanity check: the labeller's cached values equal its snapshot keys. | Catches a mis-declared row; costs nothing (graph `node.value`). |
| Export-time cycle check: `leaf_closure(labeller)` must not contain any input keyed on the labelled axis. | Inputs on that axis are mapped through the labeller before the model can read them. |
| Every labeller cell must be in the extracted graph (`_retained`). | Off-graph cells cannot be compiled; see downstream note. |

Direction can be `internal` (formula headers) or `constant`/`input` (literal
headers, which then relabel to themselves; harmless and lets one schema cover
both).

### Runtime (`export_runtime/tensor.py`, `inverted_tree/runtime.py`)

- `Tensor.relabel(**labels: Tensor) -> Tensor`: replace the keys of the named
  axis by `labels[key]` for each key, keeping positions. Raises
  `AxisError` on duplicate or wrongly typed label values (Excel tolerates
  duplicate headers; a tensor cannot), naming the axis and the colliding
  labels.
- `relabel_input(tensor, labels: Tensor, axis: str, *, series_id) -> Tensor`:
  the inverse, mapping caller keys (labels) to snapshot keys; unknown
  labels raise `SchemaError` listing the labels the model accepts, so a
  caller passing 2024 to a model whose first year is 2026 gets a real error
  instead of silence.
- Both ~30 lines; no formula evaluation.

### Codegen (`inverted_tree/`)

- `catalog.py`: `BoundSeries.axis_labels: str | None`; `SeriesCatalog.labeller_for(axis_name, keys) -> BoundSeries | None`.
- `emit.py::plan_inverted_tree`: the two export-time checks above.
- `named_emit.py`:
  - `emit_named_data`: `LABELLED_AXES = {'TIME_PERIOD': 'engine_year_labels'}`.
  - `emit_named_api` / `Model`: inputs keyed on a labelled axis are stored
    raw and exposed as `cached_property` that calls `relabel_input` with the
    labeller attribute; `param_ids` of the labeller are already available on
    the model. Model attributes for internals stay snapshot-keyed (they are
    the engine). Output attributes and `compute_*` return
    `tensor.relabel(TIME_PERIOD=self.engine_year_labels)`.
  - `_public_function`: add the labeller's leaf closure to the output's
    parameters (the label series' inputs are now inputs of every output on
    that axis; that is the point).
  - `emit_named_validation`: schema checks for labelled-axis inputs validate
    axis names and key types only; coordinate membership is checked after
    relabelling.
  - Docstrings: outputs on a labelled axis say which input(s) determine the
    keys.
- `series_bindings/`: schema field, `versions.py` (1.17.0), `normalize.py`
  pass-through, `validate.py` rules, docstring renderers.

### Parity harness

`tests/integration/utils/parity_harness.py` maps evaluator cells to tensor
coordinates via `coordinate_cells` (snapshot keys). Add a step that
relabels expected values through the labeller's evaluator values when the
output axis is labelled, so live-Excel and evaluator comparisons keep working
when the test overrides the header input. Cache-based comparisons are
unaffected at the snapshot value.

## Test plan (RED first, in this order)

1. `tests/unit/exporter/export_runtime/test_tensor.py`: `relabel` keeps
   positions and values; duplicate labels raise; `relabel_input` maps labels
   and rejects unknown ones.
2. `tests/unit/series_bindings/test_validate.py`: each rule in the table.
3. Fixture: extend `tests/fixtures/inverted_tree/tiny_dsa/tiny-dsa.xlsx` so
   `Engine!C5:G5` become `=Inputs!<first_year>`, `=C5+1`, …, and bind
   `engine_year_labels` with `internal: {}` + `axis_labels: TIME_PERIOD`
   (keep a copy of the literal-header workbook so existing tests stay
   green). Add `first_projection_year` as an input.
4. `tests/integration/exporter/inverted_tree/test_tiny_dsa_canary.py`:
   `compute_*(first_projection_year=2030)` returns keys 2030…, values equal
   the 2024 run shifted by position; an input tensor keyed 2030… is
   accepted; one keyed 2024… raises with the accepted labels in the message;
   a formula that compares a header value to an input (add one to the
   fixture) changes with the input.
5. `test_public_computation_contract.py`: labelled outputs still satisfy the
   named-axis-v1 contract (immutable, `sel`, `items`, JSON round trip).
6. `test_design_properties.py`: `internals.py` is byte-identical with and
   without `axis_labels` on the fixture (proves the change is boundary-only).
7. Parity: evaluator run with the header input overridden versus package
   output, mapped through the labeller.

## Downstream (lic-dsf) steps, after release

1. Graph extraction is target-driven (`workbook_config.TARGETS`). The
   engine-sheet header rows (`'Baseline - external'!M8:AG8` and siblings) are
   not in today's graph because no output formula reads them, which is why
   nothing binds them. Add one header row per distinct label formula to the
   targets and bind it with `axis_labels: TIME_PERIOD`.
2. `first_projection_year` then becomes a real parameter of every output on
   `TIME_PERIOD` through the labeller's leaf closure, and the dead
   `1990–2100` domain gets a real consequence.
3. Language: change `start_working_language` from `constant: {}` to
   `input: {}`. The 146 translation references already compile; this alone
   makes them live. Bind the scenario label columns (`'PV Stress'!A27:…`)
   with `axis_labels: SCENARIO` so scenario keys follow the language.
4. Rerun the differential at the snapshot values (must be identical) and at
   one shifted year and one other language against live Excel.

## Out of scope (say so in the docs)

- **Translated input enums.** `input.domain` enum literals and
  `input.value_map` are snapshot strings. Making them follow the language
  is a separate small feature (`domain: {values_from: <series_id>}` resolved
  at runtime). Until then the contract is "supply the value the workbook
  cell holds in the snapshot language". No canonicalisation layer.
- Rewriting internal keys to ordinals. Cosmetic, high churn, no behavioural
  gain.
- Non-injective or non-scalar label formulas (duplicate headers). Excel
  allows them; the tensor contract cannot. Fail loudly.
- The translation row-451 column order. Workbook bug.

## Effort

excel-grapher: about one engineer-week (schema/validate ~150 LOC, catalog and
plan checks ~100, runtime ~80, named_emit ~200, harness ~60, tests ~400,
docs). Downstream: binding authoring plus one canonical export (the export
itself is ~14 minutes of codegen) and a differential run.
