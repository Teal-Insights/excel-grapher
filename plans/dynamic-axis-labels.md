# Dynamic axis labels (formula-derived row and column headers)

**Status:** design proposal, not implemented. Motivating report:
[lic-dsf-extraction-pipeline#53](https://github.com/Teal-Insights/lic-dsf-extraction-pipeline/issues/53).

## The problem

A series binding names each dimension of a series and says where its keys come
from. For a table laid out across columns, `bind.kind: column_header` reads the
key from a header cell; `row_label` does the same down a column. The binder
resolves those cells against a `data_only=True` workbook load
(`series_bindings/resolve.py`), so a key is whatever the header cell displayed
when the workbook was last saved. Call that value the **snapshot key**.

Snapshot keys become the keys of an `Axis`
(`exporter/export_runtime/tensor.py`, a frozen dataclass), and `Axis` keys are
the public coordinates of every tensor the generated package accepts and
returns.

That is correct while a header cell holds a literal. When it holds a formula,
the package freezes a value that Excel recomputes. The LIC-DSF workbook does
exactly that:

```
'Input 1 - Basics'!C18   = 2024                 <- the first_projection_year input
Macro-Debt_Data!U4       = 'Input 1 - Basics'!C18
'Baseline - external'!N8 = Macro-Debt_Data!U4
'Baseline - external'!M8 = N8-1, O8 = N8+1      <- repeated on every engine sheet
```

The downstream report finds that the generated package accepts
`first_projection_year`, validates it against a 1990-2100 domain, and then
ignores it: 49 of 87 public inputs and all 15 chart outputs stay keyed
2024-2044 whatever the caller passes. The workbook's language selector has the
same shape. It is bound as a `constant`, which freezes every translated row
label and scenario name to English, even though the package still carries 146
translation references that would otherwise be live.

Two facts follow, and a fix has to respect both.

1. **The keys are wrong.** A package that reports `TIME_PERIOD=2024` when the
   caller asked for 2026 is misdescribing its own output.
2. **The values are already right.** Header cells are ordinary formula cells.
   Where they sit in the extracted dependency graph they compile today like any
   other cell, so computing a header needs no new reader, no formula-AST
   extraction in the binder, and no runtime evaluation layer.

## Why the offset rewrite is the wrong fix

The obvious move is to make generated code year-relative: emit
`time_period - first_projection_year == 0` in place of `time_period == 2024`,
and key axes by offset. The downstream report counts 273 occurrences of that
one comparison, which makes the rewrite look like the whole job.

Those comparisons are not what they look like. They come from
`_family_condition` in `exporter/inverted_tree/named_emit.py`, which groups the
cells of a series by the formula each carries and then describes the group as a
condition over axis keys. In the workbook, `'PV Stress'!D30 = D29` while
`E30 = SUM(D29:E29)`: one column seeds a running total and the rest continue
it. The emitted branch means "the cell in the first projection column", which
is a fact about worksheet layout. Changing `'Input 1 - Basics'!C18` in Excel
relabels headers; it does not move formulas. So those branches have to keep
their snapshot keys. Rewriting them to input-relative arithmetic would be wrong
for any header formula that is not affine in a single input, and a no-op for
the ones that are.

The rest of the emitter is positional in the same way: relative lags
(`time_period - 1`, from `_integer_driver` in `ast_emit.py`), `OFFSET`
lowering, range tables, literal tables, diagonal conditions, and coordinate
provenance. Re-keying axes to ordinals would also push label arithmetic onto
every caller and every chart.

The real gap is narrower. **A snapshot key plays two roles at once: the
internal name of a position, and the label the caller sees. They coincide only
when the header is a literal.** Separating them is a change at the public
boundary of the generated package, not to emission.

## Design

Keep snapshot keys as internal position identifiers everywhere. Add an authored
link from an axis to the series whose values are that axis's labels, called the
axis's **labeller**, and translate at the boundary.

| Layer | Generated module | Keyed by |
| --- | --- | --- |
| Public inputs and outputs (`Model`, `compute_*`) | `api.py` | runtime labels |
| Calculation functions | `internals.py` | snapshot keys |
| Axes, domains, schemas, provenance, defaults | `data.py` | snapshot keys |

- **Outputs.** `compute_*` and the output attributes of `Model` evaluate the
  labeller like any other series and return the tensor with that axis's keys
  replaced by the labeller's values, positions unchanged.
- **Inputs keyed on a labelled axis.** The caller supplies a tensor keyed by
  labels. The package evaluates the labeller from the other inputs, maps label
  back to snapshot key, and hands a snapshot-keyed tensor to `internals`.
- Everything in between is untouched, so dispatch, lags, `OFFSET`, range
  tables, literal tables, provenance, and codegen fingerprints do not move.

Values keep Excel parity because a labeller is the header cell's own formula,
evaluated by the same runtime as every other cell.

### Binding schema (1.17.0)

One new series-level field, on the series that owns the header cells. The
example is `tests/fixtures/inverted_tree/tiny_dsa/`, as it would look after the
fixture change in step 3 of the test plan below (today those header cells hold
the literals 1 to 5 and the series is bound `constant: {}`):

```yaml
- id: engine_year_labels
  sheet: Engine
  data_range: Engine!C5:G5
  internal: {}              # header cells hold formulas: =Inputs!$B$4, =C5+1, ...
  axis_labels: TIME_PERIOD  # this series' values are the runtime labels of its own TIME_PERIOD key
  structure:
    measure:
      concept: OBS_VALUE
      dtype: int
      bind: {kind: data_cell, read: int}
    dimensions:
      - id: TIME_PERIOD
        concept: TIME_PERIOD
        role: key
        scope: cell
        bind: {kind: column_header, header_row: 5, read: int}
  key: [TIME_PERIOD]
```

The direction may be `internal` (formula headers) or `constant` / `input`
(literal headers, which relabel to themselves; harmless, and it lets one rule
set cover both). Series that do not declare `axis_labels` are unaffected, so a
second `Year` row over the same keys, such as `Engine!C13:G13` in the fixture,
stays an ordinary series.

Rules, enforced as error issues in `series_bindings/validate.py` unless marked
otherwise:

| Rule | Why |
| --- | --- |
| `axis_labels` names one of the series' own key fields, and that field is the series' only key. | A label is a function of position along one axis. |
| The measure dtype matches the axis key type, so `int` or `string`. | `Axis` accepts only `str` or `int` keys, so a float or date header has to be rejected at authoring time. |
| At most one labeller per axis identity, where identity is the dimension id plus the snapshot key set. A second series declaring the same `axis_labels` over identical keys is an error. | Axes are shared across series by `(name, keys, key_type)` in `inverted_tree/named_axes.py`. |
| An axis whose snapshot keys are a subset of its labeller's keys is relabelled through it. Keys outside the labeller are an export error naming the missing ones. | A history block of literal columns and a projection block of formula columns are one axis with one labeller. |
| Export time: the labeller's cached values equal its snapshot keys. | Catches a mis-declared row for free, from `node.value` on the graph. |
| Export time: `leaf_closure(labeller)` contains no input keyed on the labelled axis. | Such an input can only be relabelled after the labeller is evaluated, so this would be a cycle. |
| Export time: every labeller cell is in the extracted graph (that is, `_retained`). | Off-graph cells cannot be compiled. See the downstream section: this is the one that bites first. |

### Runtime

Two helpers, roughly 30 lines each, no formula evaluation. They go in
`exporter/export_runtime/tensor.py`, which is copied verbatim into the
generated package as `tensor.py`.

- `Tensor.relabel(**labels: Tensor) -> Tensor` replaces the keys of the named
  axis with the labeller's value at each key, keeping positions and values.
  Duplicate or wrongly typed label values raise `AxisError` naming the axis and
  the collision. Excel tolerates two columns headed 2024; a tensor axis cannot.
- `relabel_input(tensor, labels, axis, *, series_id) -> Tensor` is the inverse,
  mapping caller keys to snapshot keys. An unknown label raises `SchemaError`
  listing the labels the model does accept, so a caller passing 2024 to a model
  whose first year is 2026 gets a real error rather than silence.

### Codegen

- `inverted_tree/catalog.py`: `BoundSeries.axis_labels: str | None`, and
  `SeriesCatalog.labeller_for(axis_name, keys) -> BoundSeries | None`.
- `inverted_tree/emit.py`, in `plan_inverted_tree`: the three export-time
  checks in the rules table.
- `inverted_tree/named_emit.py`:
  - `emit_named_data` emits the axis-to-labeller map, for example
    `LABELLED_AXES = {'TIME_PERIOD': 'engine_year_labels'}`, as the single
    place `api.py` and the tests read the relationship from.
  - `emit_named_api` stores a labelled-axis input raw and exposes it as a
    `cached_property` that calls `relabel_input` against the labeller
    attribute, which `Model` can already evaluate. Internal attributes stay
    snapshot-keyed. Output attributes and `compute_*` return
    `tensor.relabel(TIME_PERIOD=self.engine_year_labels)`.
  - `_public_function` adds the labeller's leaf closure to the output's
    parameters. The labeller's inputs become inputs of every output on that
    axis, which is the whole point: that is how `first_projection_year` starts
    mattering.
  - `emit_named_validation` checks axis names and key types for a
    labelled-axis input, and defers coordinate membership until after
    relabelling.
  - Output docstrings name the inputs that determine the keys.
- `series_bindings/`: the schema field, `versions.py` (1.17.0),
  `normalize.py` pass-through, `validate.py` rules, docstring renderers.

### Parity harness

`tests/integration/utils/parity_harness.py` maps evaluator cells to tensor
coordinates through `coordinate_cells`, which is snapshot-keyed. Add a step
that relabels expected values through the labeller's evaluator values when the
output axis is labelled, so evaluator and live-Excel comparisons keep working
when a test overrides the header input. Comparisons at the snapshot values are
unaffected.

## Test plan

Test first, observe the failure, then implement.

1. `tests/unit/exporter/export_runtime/test_tensor.py`: `relabel` keeps
   positions and values; duplicate and wrongly typed labels raise;
   `relabel_input` maps labels and rejects unknown ones with the accepted
   labels in the message.
2. `tests/unit/series_bindings/test_validate.py`: one case per rule above.
3. Fixture. In `tests/fixtures/inverted_tree/tiny_dsa/tiny-dsa.xlsx`, add a
   `first_projection_year` input on the Inputs sheet and make the `Year` header
   rows over the projection columns formulas over it, keeping them consistent
   with each other. Bind `engine_year_labels` (`Engine!C5:G5`) as
   `internal: {}` with `axis_labels: TIME_PERIOD`; the other `Year` rows stay
   ordinary series. Keep the literal-header workbook alongside it so existing
   tests stay green. The fixture already carries the semantic case this feature
   exists for: `Engine!C10 = IF(C5>=Inputs!$B$21,1,0)` compares a header value
   against the `shock_year` input.
4. `tests/integration/exporter/inverted_tree/test_tiny_dsa_canary.py`: with the
   header input shifted, `compute_*` returns shifted keys and values equal to
   the snapshot run at the same positions; an input tensor keyed by the shifted
   labels is accepted; one keyed by the snapshot labels raises; the shock
   activation flag moves when `shock_year` moves with the labels.
5. `test_public_computation_contract.py`: a relabelled output still satisfies
   the named-axis contract, meaning immutability, `sel`, `items`, and JSON
   round trip.
6. `test_design_properties.py`: `internals.py` is byte-identical with and
   without `axis_labels` on the fixture. This is the property that makes the
   change safe to land, so it is worth asserting directly.
7. Parity: an evaluator run with the header input overridden against package
   output, mapped through the labeller.

## Downstream adoption (lic-dsf-extraction-pipeline), after release

1. Graph extraction there is target-driven through `workbook_config.TARGETS`.
   The engine header rows (`'Baseline - external'!M8:AG8` and its siblings) are
   not in today's graph, because no output formula reads them, which is why
   nothing binds them. Add one header row per distinct label formula to the
   targets and bind it with `axis_labels: TIME_PERIOD`.
2. `first_projection_year` then reaches every output on `TIME_PERIOD` through
   the labeller's leaf closure, and its 1990-2100 domain acquires a
   consequence.
3. Language: change `start_working_language` from `constant: {}` to
   `input: {}`. The 146 translation references already compile, so that switch
   alone makes them live. Bind the scenario label column (`'PV Stress'!A27`
   and its block) with `axis_labels: SCENARIO` so scenario keys follow the
   language.
4. Re-run the differential at the snapshot values, where it must be identical,
   and then at one shifted year and one other language against live Excel.

## Out of scope

State these limits in the user guide rather than leaving them implied.

- **Translated input enums.** `input.domain` enum literals and
  `input.value_map` needles are snapshot strings, so a French caller still
  passes the English token. Making a domain follow a labeller is a separate
  small feature, something like `domain: {values_from: <series_id>}` resolved
  at runtime. No canonicalisation layer in this change.
- **Ordinal internal keys.** Cosmetic, high churn, no behavioural gain.
- **Non-injective or non-scalar header formulas.** Excel permits duplicate
  headers; the tensor contract cannot. Fail loudly, as above.
- **The LIC-DSF translation sheet's row 451 column order**, which yields
  `['Spanish']` under `language == 'French'`. That is a faithful reproduction
  of a workbook bug, not a codegen defect.

## Effort

excel-grapher, roughly one engineer-week: schema and validation about 150 LOC,
catalog and export-time checks about 100, runtime about 80, `named_emit` about
200, parity harness about 60, tests about 400, plus docs. Downstream: binding
authoring, one canonical export (codegen itself runs about 14 minutes), and a
differential run.
