---
status: proposal
tracking: https://github.com/Teal-Insights/excel-grapher/issues/841
supersedes: relabel-at-the-boundary design in #841 body, implemented on PR #847 (codex/implement-issue-841)
pin: 2d93d23 (v19.0.2); findings in §1 reproduced against the tiny_dsa fixture at that commit
updated: 2026-09-15
---

# Dynamic axis labels, keyed by label all the way down

A snapshot key plays two roles today: the internal name of a position along
an axis, and the label a caller sees. #841 proposed keeping snapshot keys
inside `internals.py` and translating at the `api.py` boundary; PR #847 built
that. The follow-up on #841 rejects it: internals then lie about labels
(`internals.out.sel(TIME_PERIOD=2024)` selects the wrong year after a shift),
outputs leave `Series` for bare `Tensor`, input validation is weakened to
`validate_structure`, and the internals surface is useless for drilldown and
differential testing.

This plan makes **runtime labels the keys everywhere**. Positions stay the
codegen's internal truth (they are what worksheet layout means), but they are
expressed as positions, never as snapshot literals, wherever the axis is
labelled. The labeller's evaluated values become the axis; every tensor on that
axis, internal or public, is keyed by those values.

## 1. What the generated package does today (pinned findings)

Generated from `tests/fixtures/inverted_tree/tiny_dsa` at the pin. Every
snapshot-key site a labelled axis would have to lose is one of these six
shapes; nothing else in the emitter mentions axis keys.

| # | Site | Emitted today | Positional meaning |
| --- | --- | --- | --- |
| 1 | Axis constant, `data.py:12` | `TIME_PERIOD_AXIS = Axis('TIME_PERIOD', (1, 2, 3, 4, 5), int)` | 5 positions |
| 2 | Family condition, `internals.py:91` (`_family_condition`) | `if time_period == 1:` seed of the recurrence | position 0 |
| 3 | Lag, `internals.py:93` (`_named_keys` via `_integer_driver`) | `baseline_path_internal[time_period - 1]` | previous position |
| 4 | Domain, schema, required, provenance, `data.py:27-29,91-93` | `Domain.product(TIME_PERIOD_AXIS)`, `row_cells('Engine', 10, 'C', TIME_PERIOD_AXIS)` | built from the axis at import |
| 5 | Defaults, `data.py:55` | `GrowthBaseline(GROWTH_BASELINE_DOMAIN, (3.5, ...))` | workbook state at snapshot |
| 6 | Range corners / steps / literal tables (`_named_range_view`, `_emit_named_offset`, `_LITERALS`) | `span(data.AXIS, 1, 3)`, `axis_step(data.AXIS, 'Growth', k)`, `_X_LITERALS[time_period,]` | positions 0..2, position of `'Growth'`, position |

Observation that makes the whole change tractable: `span` and `axis_step` in
`inverted_tree/runtime.py:109,219` are already position-based (`axis.keys.index`);
they only receive snapshot literals as arguments. `Tensor.isel` already exists.
`CoordinateReader` and `evaluate` iterate whatever `Domain` they are handed.
The emitter is positional in substance and literal only in syntax.

Second observation: the labeller is already a `deps` param wherever a formula
reads a header cell (`internals.shock_active(*, engine_year_labels, shock_year)`
at the pin, because `Engine!C10 = IF(C5>=Inputs!$B$21,1,0)` reads `C5`). The
mechanism that carries the labeller's leaf closure into public signatures
exists; it is just not applied to series that do not read the header.

## 2. Design

### 2.1 Vocabulary

- **Labeller**: a series with `axis_labels: <FIELD>` whose values are the
  labels of its own single key axis. Schema field, `validate.py` rules, catalog
  `BoundSeries.axis_labels` and `SeriesCatalog.labeller_for`, and the
  export-time checks from #847 all carry over unchanged (§4 lists what to keep).
- **Static axis**: unlabelled, or labelled by a `constant` series. Keys known
  at export. Nothing in this plan touches static axes; packages without a
  runtime-labelled axis must be byte-identical before and after (§6, P1).
- **Runtime axis**: labelled by an `internal`, `output`, or `input` labeller.
  Keys are known only once the labeller is evaluated.

### 2.2 The invariant

> For a runtime axis `A` with labeller `L`, every tensor keyed on `A` inside
> the package — inputs after validation, internals, outputs — has
> `axis.keys == tuple(L.values)`, and `L` itself is the identity map
> `label -> label` over that axis. Positional facts from the worksheet are
> emitted as positions on `A`, never as snapshot literals.

Consequences: `internals.x.sel(TIME_PERIOD=2026)` means 2026; `compute_*`
returns a real `Series`; inputs validate against the bound schema with full
membership; no `relabel`, no `_internal_*` aliases, no `validate_structure`.

### 2.3 The labeller becomes a dependency of its axis

`collect_all_deps` (`deps.py:2261`) adds one implicit edge per runtime axis:
every series with a runtime axis `A` depends on `A`'s labeller, whether or not
its formulas read the header cells. This is the single change that does the
plumbing:

- `param_ids` of every internals function on `A` gain the labeller, so the
  function can bind its own domain from `labeller.domain.axes[0]`.
- `leaf_closure` (`deps.py:2283`) then carries the labeller's inputs into every
  `compute_*` on `A` for free. `_public_function` needs no special expansion
  (#847's version is dropped).
- `Model` attributes get the labeller argument from `_model_attribute` with no
  new code path.
- `scc_external_params` treats the labeller as external to any recurrence
  group on `A`; export-time rule "no input keyed on `A` in
  `leaf_closure(L)`" (kept from #847) guarantees that is acyclic.

Static axes add no edge, so unlabelled packages do not change (P1).

### 2.4 The labeller compiles positionally, then publishes over itself

A labeller cannot iterate its own axis before it has computed the labels. Its
internals function is the one place a **positional domain** appears:

```python
@publish(data.ENGINE_YEAR_LABELS_SCHEMA, cells=data.ENGINE_YEAR_LABELS_CELLS)
def engine_year_labels(*, first_projection_year: int | str) -> data.EngineYearLabels:
    def formula(_p: int) -> int | str | None:
        if _p == 0:
            return as_measure(first_projection_year, 'int')
        return as_measure(xl_add(engine_year_labels[_p - 1], 1), 'int')

    engine_year_labels = CoordinateReader('engine_year_labels', data.ENGINE_YEAR_LABELS_POSITIONS, formula)
    return data.EngineYearLabels.from_labels(
        label_axis('TIME_PERIOD', tuple(engine_year_labels[p] for p in data.ENGINE_YEAR_LABELS_POSITIONS), int)
    )
```

- `data.X_POSITIONS = Domain.product(Axis('TIME_PERIOD', (0, 1, 2, 3, 4), int))`
  is emitted for labellers only.
- `label_axis(name, values, key_type)` (new, `export_runtime/tensor.py`)
  builds the `Axis` and raises `AxisError` naming the axis and the offending
  value on: a duplicate (Excel tolerates two columns headed 2024, a tensor
  axis cannot); a wrong type (a float header, an Excel error code string on an
  `int` axis). Fail loudly, as #841 already decided.
- `Series.from_labels(axis)` publishes the identity tensor `label -> label`.
  The labeller's own `Series` facade validates like any other (§2.6).

Labellers are emitted by the existing `_semantic_body` with the host's
coordinate variable bound to the position; every rule in §2.5 applies to the
labeller's body with "position variable" in place of "key variable".

### 2.5 Emitting positions instead of snapshot literals

All in `named_emit.py` and `ast_emit.py`, gated on "this axis is a runtime
axis" (`catalog.labeller_for(axis.name, axis.keys)` is not `None` and not
`constant`). The gate is what keeps static packages byte-identical.

| Site (§1) | Today | Runtime axis |
| --- | --- | --- |
| 2 `_family_condition` `==`, `>=`, `<=`, `in`, diagonal | `time_period == 1` | `_p_time_period == 0`; `>= 3` → `_p >= 2`; `in (1, 3)` → `_p in (0, 2)`; diagonal `right - left == k` → `_p_right - _p_left == k` when either axis is runtime (that is what a worksheet diagonal means) |
| 3 `_named_keys` integer lag | `x[time_period - 1]` | `x[axis_step(_ax_time_period, time_period, -1)]` |
| 3 `_named_keys` literal key on a runtime axis | `x[2]` | `x[_ax_time_period.keys[1]]` |
| 3 `_named_keys` cross-field reuse (`current == target`, `_string_driver`) | host variable | host variable only when both fields share one labeller identity; otherwise `keys[p]` of the producer's axis. Cross-field equality between a runtime axis and a static axis on *snapshot* values is an export error (it is a coincidence at the snapshot, not a layout fact) |
| 3 `_key_template` (f-string over a host label) | template | export error when the host axis is runtime (out of scope, see §7) |
| 6 `_named_range_view` corners | `span(AXIS, 1, 3)` | `span(_ax, _ax.keys[0], _ax.keys[2])`; whole-axis `data.AXIS.keys` → `_ax.keys` |
| 6 `_emit_named_offset` anchor literal | `axis_step(AXIS, 'Growth', k)` | `axis_step(_ax, _ax.keys[p], k)` |
| 6 `_LITERALS` tables | keyed by snapshot coordinate | keyed by position on runtime axes: `data._X_LITERALS[_p_time_period,]` |

Prologue emitted once per formula body that needs it:

```python
_ax_time_period = engine_year_labels.domain.axes[0]
def formula(time_period: int) -> float | str | None:
    _p_time_period = _ax_time_period.position(time_period)
```

`Axis.position(key) -> int` is a two-line addition to `tensor.py`
(`Domain` already keeps the index; expose it on `Axis`). The loop variable
stays the label so ordinary reads `growth_baseline[time_period]` are untouched
and read naturally in tracebacks.

`_integer_driver` ranking and `_key_field_axis` geometry do not change: they
decide *which* host variable a key moves with; only the arithmetic emitted
afterwards changes.

### 2.6 `data.py`: templates for runtime axes, constants for static ones

For a series with at least one runtime axis, the six static objects become
templates bound at call time. Static-axis series keep today's emission
verbatim.

```python
TIME_PERIOD_AXIS = AxisTemplate('TIME_PERIOD', int, size=5, labeller='engine_year_labels', snapshot=(1, 2, 3, 4, 5))
SHOCK_ACTIVE_DOMAIN = DomainTemplate.product(TIME_PERIOD_AXIS)
SHOCK_ACTIVE_REQUIRED = SHOCK_ACTIVE_DOMAIN            # or DomainTemplate.explicit over positions
SHOCK_ACTIVE_CELLS = row_cells('Engine', 10, 'C', TIME_PERIOD_AXIS)   # provenance over a template
SHOCK_ACTIVE_SCHEMA = SchemaTemplate('shock_active', SHOCK_ACTIVE_REQUIRED, INT_VALUES)
class ShockActive(Series[int | str | None]):
    schema = SHOCK_ACTIVE_SCHEMA
LABELLED_AXES = {'TIME_PERIOD': 'engine_year_labels'}
```

- `AxisTemplate.bind(axis)` checks `name`, `key_type`, `size`; returns `axis`.
  `snapshot` is kept for diagnostics and for defaults (below).
- `DomainTemplate.bind(**labellers)` → `Domain`. Explicit (ragged) domains are
  stored as positional coordinates and mapped through the bound axes.
- `SchemaTemplate.bind(**labellers)` → `TensorSchema`; `.validate(tensor)`
  without binding resolves each runtime axis from the tensor's own axis after
  checking size and type — so `Series.__post_init__` keeps validating on
  construction with no model in scope. The *identity* check (keys equal the
  labeller's values) happens where a tensor enters: `Model` for inputs, and by
  construction for internals, whose domain was bound from the labeller.
- Provenance: `row_cells` / `column_cells` / `block_cells` / `grid_cells`
  accept an `AxisTemplate` and return a `ProvenanceTemplate`; `__cells__` on
  the published function is that template, and `Model.cells(series_id)` binds
  it. `grid_cells` over an explicit runtime domain maps positions.
- Defaults for inputs on a runtime axis are published over the **snapshot
  axis** (`AxisTemplate.snapshot`): the workbook's saved values are only
  consistent with the workbook's saved labels. A caller who shifts
  `first_projection_year` and passes `GROWTH_BASELINE_DEFAULT` gets the
  "unknown labels ..., accepted labels ..." `SchemaError` from #847's message,
  which is the correct, loud outcome (#841 asked for exactly this rather than
  silence).
- `publish(schema=SchemaTemplate)` sets `__domain__` to the template;
  `as_records` accepts a template and reads axes from the result tensor.

### 2.7 `api.py` and `validation.py`

`Model.__init__` today checks every input in one pass. With runtime axes the
order is:

1. Check inputs whose axes are all static (unchanged `CHECKS`).
2. Evaluate labellers. Safe because of the export-time rule that no input on
   a runtime axis is in a labeller's leaf closure; labellers are ordinary
   `cached_property` attributes, so this is just touching them.
3. Check inputs on runtime axes against the bound schema:
   `validation._check_growth_baseline(value, TIME_PERIOD=self.engine_year_labels)`
   — full `TensorSchema.validate`, no `validate_structure`. Unknown labels
   raise `SchemaError` listing the accepted labels.

`emit_named_validation` emits the extra keyword parameters from
`LABELLED_AXES`; `_input_check` binds before validating. Nothing is stored
raw; no `_raw_`/`_internal_` split.

`compute_*` functions are unchanged in shape: their signature already grows
by the labeller's leaves through `leaf_closure` (§2.3), and they return
`Model(...).<output>`, which is a `Series` over the runtime axis. Docstrings
name the inputs that determine the keys (keep #847's `key_note`).

### 2.8 Runtime (`export_runtime/tensor.py`, copied verbatim into packages)

New, about 120 lines total: `Axis.position`, `label_axis`, `Series.from_labels`,
`AxisTemplate`, `DomainTemplate`, `SchemaTemplate`, `ProvenanceTemplate`
(provenance one lives in `export_runtime/provenance.py`). Removed from #847:
`Tensor.relabel`, `_replace_axes`, `relabel_input`,
`TensorSchema.validate_structure`, the `as_records` loosening.

## 3. Why not the alternatives

- **Relabel at the boundary (#847).** Rejected on #841 for the reasons in the
  preamble. Its authoring surface, validation rules, catalog fields, and
  export-time checks are right and are kept.
- **Ordinal internal keys (0..n-1) everywhere.** Internals would stop lying,
  but `sel(TIME_PERIOD=2026)` still would not work on internals, defaults and
  provenance still need a translation, and every static-axis package churns.
  Positions are the right *emission*, not the right *key*.
- **Input-relative arithmetic (`time_period - first_projection_year == 0`).**
  Rejected in #841: wrong for any header formula that is not affine in one
  input, and a no-op for the ones that are.

## 4. Relationship to PR #847

Rebase #847 onto `main` (it is currently `mergeable_state: dirty`) and then:

Keep as is: `series_binding.schema.json` `axis_labels` + version 1.17.0;
`validate.py::_validate_axis_labels` and its tests; `catalog.py`
`BoundSeries.axis_labels`, `SeriesCatalog.labeller_for`, `build_catalog`
pass-through; `emit.py::plan_inverted_tree` export-time checks (duplicate
labeller identity, off-graph cells, cached values equal snapshot keys, no
input on the axis in the labeller's leaf closure, labeller covers the axis);
`named_codegen_fingerprint` including `axis_labels`; `LABELLED_AXES` in
`data.py`; the `key_note` docstring; the user-guide section (rewrite the
"how it works" paragraph).

Remove: `Tensor.relabel`, `_replace_axes`, `relabel_input`,
`validate_structure`, `as_records` loosening; `LABELLED_INPUTS`,
`_needs_internal_alias`, `_runtime_labellers`, `_internal_*` and `_raw_*`
emission in `_model_attribute`, `_model_recurrence_group`, `emit_named_api`;
the `_public_function` closure expansion (subsumed by §2.3); the
`_input_check` `validate_structure` branch; `test_tensors.py` relabel tests;
the `test_design_properties.py` relabel-through-labeller oracle (replaced by
§6 P3).

## 5. Implementation order (each step RED → GREEN → refactor)

1. **Runtime primitives** (`tensor.py`, `provenance.py`; unit
   `tests/unit/exporter/export_runtime/test_tensor.py`). `Axis.position`;
   `label_axis` raising on duplicate / wrong type; `Series.from_labels`;
   `AxisTemplate`/`DomainTemplate`/`SchemaTemplate` bind and validate-by-
   own-axes; `ProvenanceTemplate.bind`; `publish`/`as_records` with a
   template. No emitter change yet.
2. **Fixture** (`tests/fixtures/inverted_tree/tiny_dsa/`, kept beside a copy
   of the literal-header workbook so existing tests stay green). Add
   `first_projection_year` on `Inputs`; make `Engine!C5 = Inputs!$B$4`,
   `D5:G5 = C5+1`; make `Engine!C13:G13` and `Outputs!B11:F11` `=Engine!C5`
   etc. so headers stay consistent; bind `engine_year_labels` as
   `internal: {}` with `axis_labels: TIME_PERIOD`; bind the new input.
   `Engine!C10 = IF(C5>=Inputs!$B$21,1,0)` already gives the semantic case.
   Write the §6 P2/P4 canary tests now; watch them fail.
3. **Implicit labeller edge** (`deps.py::collect_all_deps`, `schedule.py`).
   Unit test on the fixture: `inspect.signature(internals.growth_baseline
   consumer)` and every `compute_*` on `TIME_PERIOD` include the labeller /
   `first_projection_year`. Export-time cycle rule already covers the
   pathological case; add the cross-field-equality rule from §2.5.
4. **Labeller body** (`_semantic_body` positional mode, `X_POSITIONS`
   emission, `from_labels` publish). Test: the labeller evaluates to the
   shifted years; a duplicate-header workbook raises `AxisError` from
   `compute_*`.
5. **Positional emission** (`_family_condition`, `_diagonal_condition`,
   `_union_condition`, `_named_keys`, `_named_range_view`,
   `_emit_named_offset`, literal tables). One unit test per row of the §2.5
   table using the `tests/unit/exporter/inverted_tree/test_shape_*` style
   (`a10_other_series_lag`, `a22_shift_k`, an OFFSET shape, a diagonal
   shape) with the header row made a formula. Assert generated source
   contains no snapshot literal of the labelled axis (P3) and the shifted
   oracle (P2).
6. **`data.py` templates** (`emit_named_data`, `_domain_source`,
   `_coordinates_source`, `_provenance_source`, `_read_defaults`). Gate on
   runtime axes so P1 holds; run the full corpus differential
   (`test_design_properties.py`) with P1 asserted as a byte-compare against
   packages generated at the pin.
7. **`api.py` / `validation.py`** (§2.7). Canary: shifted input tensor
   accepted, snapshot-keyed tensor rejected with the accepted-labels message,
   `Model.cells('shock_active')` bound.
8. **Parity harness.** `named_input_kwargs` and the evaluator comparison in
   `test_design_properties.py::_package_matches_evaluator` map through the
   labeller's evaluator values when an output axis is runtime; add one
   evaluator run with `first_projection_year` overridden (P5). Live-Excel
   variant `pytest.skip`s without automation, per `.cursor/rules/parity.mdc`.
9. **Docs.** User guide 05 (authoring, the defaults caveat, the loud
   failures), `AGENTS.md` note that `internals` are label-keyed, changelog.
   Then §8 downstream.

Steps 1–2 and 3–5 can be worked in parallel by two people; 6 depends on 4–5;
7 on 6; 8 on 7.

## 6. Properties to assert (replace the #847 tests)

- **P1 Static packages are byte-identical.** Every corpus package with no
  runtime-labelled axis generates the same `internals.py`, `data.py`,
  `api.py`, `validation.py` before and after. (Replaces #847's "internals
  identical with and without `axis_labels`", which is false by design here.)
- **P2 Shift oracle.** With the labeller's input shifted by *k*, every
  internal and output series on the axis has keys shifted by *k* and values
  equal to the snapshot run at the same positions; `shock_active` moves when
  `shock_year` moves with the labels.
- **P3 No snapshot literals.** For a runtime axis, the generated
  `internals.py` contains no comparison, subscript, or `span`/`axis_step`
  argument that is a snapshot key of that axis. Check via the emitter's own
  AST, not a regex.
- **P4 Contract.** A `compute_*` result on a runtime axis is a `Series`
  instance of its facade class, immutable, `sel`/`items`/JSON round trip; an
  input keyed by snapshot labels raises `SchemaError` naming the accepted
  labels; a duplicate header raises `AxisError`.
- **P5 Parity.** `FormulaEvaluator` with the header input overridden equals
  package output coordinate-by-coordinate after mapping evaluator cells
  through the labeller's evaluated values.
- **P6 Internals are honest.** `Model(...).baseline_path_internal.sel(TIME_PERIOD=<shifted year>)`
  equals the evaluator's value for that year's cell.

## 7. Out of scope (state in the user guide)

Unchanged from #841: translated input enums / `value_map` needles;
non-injective header formulas (fail loudly); the LIC-DSF row-451 column-order
artefact. Added: `_key_template` over a runtime string axis (export error
with a clear message); cross-field key equality between a runtime axis and a
static axis (export error).

## 8. Downstream (lic-dsf-extraction-pipeline #53), after release

As in #841 §"Downstream adoption": add one header row per distinct label
formula to `workbook_config.TARGETS`, bind with `axis_labels: TIME_PERIOD`;
`first_projection_year` reaches every output on that axis through the
implicit edge; flip `start_working_language` to `input: {}` and bind the
scenario label column with `axis_labels: SCENARIO`. Then the differential at
snapshot values (must be identical) and at one shifted year and one other
language against live Excel. The `internals` surface is now usable for that
differential directly, which was the point.

## 9. Effort

excel-grapher, roughly two engineer-weeks: runtime ~150 LOC, deps edge ~40,
positional emission ~250 (spread over eight functions), `data.py` templates
~200, api/validation ~80, fixture + tests ~600, docs. The largest risk is
step 5: the emitter's key-arithmetic sites are scattered, and each needs its
own shape test. Step 6 is the second: templates must not leak into static
packages (P1 catches it).
