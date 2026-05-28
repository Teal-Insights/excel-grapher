# Series Bindings Micro-Workbook


The workbook [series_bindings.xlsx](series_bindings.xlsx) offers a
sandbox for demonstrating **declarative series bindings**:
analyst-authored YAML that describes SDMX-shaped structure and
spreadsheet binds to inform the input API during code generation.
Tooling validates bindings against dependency-graph leaves.

``` python
from pathlib import Path
from pprint import pformat
from textwrap import dedent

import yaml
from excel_grapher.grapher import create_dependency_graph, DependencyGraph
from excel_grapher.series_bindings import (
    derive_input_series,
    derive_output_series,
    expand_data_range,
    resolve_series_binding,
    validate_bindings_document,
    validate_series_bindings,
)

workbook_path = Path("series_bindings.xlsx")
```

## 01. Borvelia row series (offset column headers)

The first example demonstrates a time series with a few idiosyncratic
(and therefore challenging) features: the year headers are offset
integers rather than calendar years, the values are in a “wide row”
rather than the more standard columnar layout, and the rows are grouped
by country, such that the same indicator is repeated for each country,
and you need to look in a different row to find the country name for a
given series.

| Year                             | 1   | 2    | 3   | 4   | 5   |
|----------------------------------|-----|------|-----|-----|-----|
| Borvelia                         |     |      |     |     |     |
| Real GDP growth (% per annum)    | 2.1 | 2.2  | 2.3 | 2.4 | 2.5 |
| Real interest rate (% per annum) | 1   | 1.1  | 1.2 | 1.3 | 1.4 |
| Primary balance (% of GDP)       | -1  | -0.5 | 0   | 0.5 | 1   |

We will target Borvelia’s primary balance series and extract the
dependency graph. The cells are constants (no formulas), so each leaf is
both an input and, trivially, its own output. Normally we would load the
workbook bindings from a sidecar file like `bindings.yaml` with the
`load_series_bindings` function, but for educational purposes below we
inline the bindings we care about:

``` python
# Build the dependency graph
graph: DependencyGraph = create_dependency_graph(
    workbook_path, ["Sheet1!F5:J5"], load_values=True
)
```

The manifest has a top-level `series` array with one entry per series.
Each entry has an `id`, a `data_range`, a `layout`, optional `input` and
`output` blocks, a `structure`, and a `key`. The structure uses SDMX
concepts for the dimensions and measure. Each dimension has a `bind`
object that describes how to read the value from the spreadsheet. Here
is the entry for the primary balance series:

``` python
bindings_yaml = dedent(
    """
    schema_version: "1.2.0"
    workbook: series_bindings.xlsx
    series:
      - id: borvelia_primary_balance
        data_range: Sheet1!F5:J5
        layout: row_series
        editable: true
        input:
          setter:
            name: set_borvelia_primary_balance
        output:
          compute:
            name: compute_borvelia_primary_balance
        structure:
          measure:
            concept: OBS_VALUE
            dtype: float
            bind:
              kind: data_cell
              read: float
          dimensions:
            - concept: REF_AREA
              role: key
              scope: series
              bind:
                kind: cell
                address: Sheet1!A2
                read: string
              include_in_record: false
            - concept: INDICATOR
              role: key
              scope: series
              bind:
                kind: row_label
                label_column: A
                read: string
                normalize: strip_trailing_unit
              include_in_record: false
            - concept: TIME_PERIOD
              role: key
              scope: cell
              bind:
                kind: column_header
                header_row: 1
                read: int
          attributes:
            - concept: UNIT_MEASURE
              role: attribute
              value: PC_GDP
              include_in_record: true
        key:
          - TIME_PERIOD
        series_context:
          REF_AREA: Borvelia
          INDICATOR: Primary balance (% of GDP)
        sdmx_notes: Wide row series with offset-style column headers 1-5.
    """
)
bindings = validate_bindings_document(yaml.safe_load(bindings_yaml))
series_by_id = {s["id"]: s for s in bindings["series"]}
```

Projection years **1–5** sit in row 1 (not calendar years). Values for
primary balance are in `F5:J5`. Country (`REF_AREA`) and indicator text
come from series-scoped binds plus `series_context`; only `TIME_PERIOD`
is required on each incoming record for matching.

``` python
series = series_by_id["borvelia_primary_balance"]
print(f"```text\n{pformat(series, indent=2, width=100)}\n```\n")
print(
    f"```text\n{str([graph.get_node(a).value for a in expand_data_range(series['data_range'], workbook=workbook_path)])}\n```\n"
)
```

``` text
{ 'data_range': 'Sheet1!F5:J5',
  'editable': True,
  'id': 'borvelia_primary_balance',
  'input': {'setter': {'name': 'set_borvelia_primary_balance'}},
  'key': ['TIME_PERIOD'],
  'layout': 'row_series',
  'output': {'compute': {'name': 'compute_borvelia_primary_balance'}},
  'sdmx_notes': 'Wide row series with offset-style column headers 1-5.',
  'series_context': {'INDICATOR': 'Primary balance (% of GDP)', 'REF_AREA': 'Borvelia'},
  'sheet': 'Sheet1',
  'structure': { 'attributes': [ { 'concept': 'UNIT_MEASURE',
                                   'include_in_record': True,
                                   'role': 'attribute',
                                   'value': 'PC_GDP'}],
                 'dimensions': [ { 'bind': { 'address': 'Sheet1!A2',
                                             'kind': 'cell',
                                             'read': 'string'},
                                   'concept': 'REF_AREA',
                                   'include_in_record': False,
                                   'role': 'key',
                                   'scope': 'series'},
                                 { 'bind': { 'kind': 'row_label',
                                             'label_column': 'A',
                                             'normalize': 'strip_trailing_unit',
                                             'read': 'string'},
                                   'concept': 'INDICATOR',
                                   'include_in_record': False,
                                   'role': 'key',
                                   'scope': 'series'},
                                 { 'bind': { 'header_row': 1,
                                             'kind': 'column_header',
                                             'read': 'int'},
                                   'concept': 'TIME_PERIOD',
                                   'role': 'key',
                                   'scope': 'cell'}],
                 'measure': { 'bind': {'kind': 'data_cell', 'read': 'float'},
                              'concept': 'OBS_VALUE',
                              'dtype': 'float'}}}
```

``` text
[-1, -0.5, 0, 0.5, 1]
```

### Validation

Before codegen, tooling intersects `data_range` with the extracted
graph’s **leaf** cells (inputs, not formulas depending on other cells).
A binding can describe a wider workbook series than the current graph
needs; setter generation uses the overlapping leaves and skips a series
with no overlap.

``` python
report = validate_series_bindings(graph, bindings, workbook=workbook_path)
print(f"```text\n{pformat(report, indent=2, width=100)}\n```\n")
```

``` text
{'issues': [], 'ok': True}
```

With `require_unique_key` (the default), the generated setter can use
key-based matching only if no two participating graph leaves have the
same `TIME_PERIOD`. That holds here: offsets 1–5 across columns F–J are
distinct.

### Coordinate resolution

**Resolution** walks each graph leaf that overlaps `data_range` and
evaluates the bind executors (`column_header` for `TIME_PERIOD`,
`data_cell` for `OBS_VALUE`, and so on). The result is a **key** tuple
used to match records and a **record** dict shaped for the generated
setter (including `OBS_VALUE` and optional attributes such as
`UNIT_MEASURE`).

``` python
resolved = resolve_series_binding(graph, workbook_path, series)
print(
    f"```text\n{pformat({'requires_address': resolved['requires_address'], 'leaf_count': len(resolved['leaves'])}, indent=2)}\n```\n"
)
leaf = next(leaf for leaf in resolved["leaves"] if leaf["key"]["TIME_PERIOD"] == 3)
print(
    f"```text\n{pformat({'address': leaf['address'], 'key': leaf['key'], 'record': leaf['record']}, indent=2)}\n```\n"
)
```

``` text
{'leaf_count': 5, 'requires_address': False}
```

``` text
{ 'address': 'Sheet1!H5',
  'key': {'TIME_PERIOD': 3},
  'record': { 'INDICATOR': 'Primary balance (% of GDP)',
              'OBS_VALUE': 0.0,
              'REF_AREA': 'Borvelia',
              'TIME_PERIOD': 3,
              'UNIT_MEASURE': 'PC_GDP'}}
```

For period 3 the matched cell is `Sheet1!H5` with value `0.0` from the
workbook. The record carries `TIME_PERIOD` for matching plus
documentation fields from `series_context`.

### Input series

For inspection and generated input APIs, bindings can be projected into
**input series**: one item per binding series with graph-leaf overlap.
This is the series-binding-derived replacement for earlier input-group
discovery ideas.

``` python
input_series = derive_input_series(graph, bindings, workbook=workbook_path)
print(
    f"```text\n{pformat({'id': input_series[0]['id'], 'key_fields': input_series[0]['key_fields'], 'cell_count': len(input_series[0]['cells'])}, indent=2)}\n```\n"
)
```

``` text
{ 'cell_count': 5,
  'id': 'borvelia_primary_balance',
  'key_fields': ['TIME_PERIOD']}
```

### Generated setter (Records API)

Passing the same binding into `CodeGenerator` appends a
`set_borvelia_primary_balance` function. Callers pass `list[dict]`
**records** with at least `OBS_VALUE` and each key field (`TIME_PERIOD`
here); the setter writes into the graph’s input map via
`EvalContext.set_inputs`.

``` python
from excel_grapher.exporter import CodeGenerator

targets = expand_data_range(series["data_range"], workbook=workbook_path)
with CodeGenerator(graph) as gen:
    code = gen.generate(
        targets,
        series_bindings=bindings,
        bindings_workbook=workbook_path,
    )

signature_start = code.index("def set_borvelia_primary_balance(")
signature_end = code.index(") -> None:", signature_start) + len(") -> None:")
print(f"```text\n{code[signature_start:signature_end]}\n```\n")
```

``` text
def set_borvelia_primary_balance(
    ctx: EvalContext,
    records: Records,
    *,
    strict: bool = True,
) -> None:
```

### Calling the setter

In an exported model, the generated setter is imported and called like
any other package function. Each record must include an `OBS_VALUE` key
and each key field (`TIME_PERIOD` here):

``` python
from exported_model import make_context, set_borvelia_primary_balance

ctx = make_context()
records = [
    {"TIME_PERIOD": 4, "OBS_VALUE": 7.5},
    {"TIME_PERIOD": 5, "OBS_VALUE": 8.0},
]

set_borvelia_primary_balance(ctx, records)
```

### Structured docstring callback

If you want richer generated docstrings for series APIs, register a
named callback and pass `series_docstring_callback` to `generate`. The
callback receives deterministic binding metadata through `ctx.contract`
and returns prose fields via `SeriesFunctionDoc`.

``` python
from excel_grapher.exporter import (
    CodeGenerator,
    FieldDoc,
    SeriesFunctionDoc,
    register_series_docstring_callback,
)

callback_name = "micro_workbook_series_docs"

register_series_docstring_callback(
    callback_name,
    lambda ctx: SeriesFunctionDoc(
        summary=f"Set {ctx.contract.series_id}.",
        purpose="Updates workbook inputs from Records.",
        record_matching="Records match cells by key fields.",
        field_descriptions={
            field: FieldDoc(description=f"Value for {field}.")
            for field in ctx.contract.required_fields
        },
    ),
    replace=True,
)

with CodeGenerator(graph) as gen:
    code_with_docs = gen.generate(
        targets,
        series_bindings=bindings,
        bindings_workbook=workbook_path,
        series_docstring_callback=callback_name,
    )

namespace_with_docs: dict = {}
exec(code_with_docs, namespace_with_docs)
doc = namespace_with_docs["set_borvelia_primary_balance"].__doc__
print(f"```text\n{doc}\n```\n")
```

``` text
Set borvelia_primary_balance.

Updates workbook inputs from Records.
Records match cells by key fields.

Required record fields:
    TIME_PERIOD: Value for TIME_PERIOD.
    OBS_VALUE: Value for OBS_VALUE.

Optional record fields:
    REF_AREA: REF_AREA If supplied, expected value: "Borvelia".
    INDICATOR: INDICATOR If supplied, expected value: "Primary balance (% of GDP)".
    UNIT_MEASURE: UNIT_MEASURE If supplied, expected value: "PC_GDP".

Source binding:
    Workbook range: Sheet1!F5:J5
    Layout: row_series
    Value type: float

Example:
    set_borvelia_primary_balance(ctx, [
        {'TIME_PERIOD': 1, 'OBS_VALUE': -1.0},
        {'TIME_PERIOD': 2, 'OBS_VALUE': -0.5},
    ])
```

The setter writes into the graph’s input map via
`EvalContext.set_inputs`. In this notebook, the generated code is
executed dynamically so the example can verify the same effect:

``` text
Sheet1!I5 after setter: 7.5
Sheet1!J5 after setter: 8.0
```

Periods **4** and **5** correspond to columns I and J; the setter
updates those leaves without requiring sheet addresses in the records.

### Output series

Bindings with an `output.compute` block can be projected into **output
series**: one item per binding with graph overlap on cells in
`data_range` (any graph node, not only leaves). Each cell carries a
static dimension map; `OBS_VALUE` is filled at runtime when the
generated `compute_*` function runs.

``` python
output_series = derive_output_series(graph, bindings, workbook=workbook_path)
print(
    f"```text\n{pformat({'id': output_series[0]['id'], 'compute_name': output_series[0]['compute_name'], 'cell_count': len(output_series[0]['cells'])}, indent=2)}\n```\n"
)
resolved_output = resolve_series_binding(
    graph, workbook_path, series, direction="output"
)
leaf_out = next(
    leaf for leaf in resolved_output["leaves"] if leaf["key"]["TIME_PERIOD"] == 3
)
print(
    f"```text\n{pformat({'address': leaf_out['address'], 'record': leaf_out['record']}, indent=2)}\n```\n"
)
```

``` text
{ 'cell_count': 5,
  'compute_name': 'compute_borvelia_primary_balance',
  'id': 'borvelia_primary_balance'}
```

``` text
{ 'address': 'Sheet1!H5',
  'record': { 'INDICATOR': 'Primary balance',
              'OBS_VALUE': None,
              'REF_AREA': 'Borvelia',
              'TIME_PERIOD': 3,
              'UNIT_MEASURE': 'PC_GDP'}}
```

Output records include **all declared dimensions** (and attributes such
as `UNIT_MEASURE`), not only the `key` fields required by the setter.
Here `OBS_VALUE` in the resolved record is a placeholder until
evaluation; the generated function overwrites it with `xl_cell`.

### Generated compute function (Records API)

The same `CodeGenerator.generate` call that emitted the setter also
appends `compute_borvelia_primary_balance` when `output.compute` is
present. It returns **`Records`** (`list[Record]`) with one dict per
graph cell in the series, each including `OBS_VALUE` and the bound
dimensions.

``` python
compute_sig_start = code.index("def compute_borvelia_primary_balance(")
compute_sig_end = code.index(") -> Records:", compute_sig_start) + len(") -> Records:")
print(f"```text\n{code[compute_sig_start:compute_sig_end]}\n```\n")
assert "Record = dict[str, object]" in code
assert "Records = list[Record]" in code
```

``` text
def compute_borvelia_primary_balance(inputs=None, *, ctx=None) -> Records:
```

### Calling the compute function

In an exported model, import the compute function alongside the setter.
Pass an optional shared `ctx` (for example after applying inputs with
the setter):

``` python
from tabulate import tabulate

from exported_model import (
    make_context,
    set_borvelia_primary_balance,
    compute_borvelia_primary_balance,
)

ctx = make_context()
set_borvelia_primary_balance(
    ctx,
    [
        {"TIME_PERIOD": 4, "OBS_VALUE": 7.5},
        {"TIME_PERIOD": 5, "OBS_VALUE": 8.0},
    ],
)
records = compute_borvelia_primary_balance(ctx=ctx)
print(tabulate(records, headers='keys', tablefmt='pipe'))
```

The compute function evaluates each bound cell via `xl_cell` and does
not require callers to pass sheet addresses unless
`include_address: true` is set on `output.compute`.

| INDICATOR       | REF_AREA | TIME_PERIOD | UNIT_MEASURE | OBS_VALUE |
|-----------------|----------|-------------|--------------|-----------|
| Primary balance | Borvelia | 1           | PC_GDP       | -1        |
| Primary balance | Borvelia | 2           | PC_GDP       | -0.5      |
| Primary balance | Borvelia | 3           | PC_GDP       | 0         |
| Primary balance | Borvelia | 4           | PC_GDP       | 7.5       |
| Primary balance | Borvelia | 5           | PC_GDP       | 8.0       |

After the setter runs, period **4** and **5** records reflect the
updated input values (`7.5` and `8.0`) while still carrying dimensions
such as `REF_AREA` and `UNIT_MEASURE` from the binding. Address-keyed
`compute_all(ctx=ctx)` remains available for the full target map;
`compute_borvelia_primary_balance` is the tabular, dimension-keyed view
of this series only.

### Modular exports

For package-style exports, generated series setters and output compute
functions are emitted from the package entrypoint and re-exported from
the package root. This keeps the callable surface next to `make_context`
and `compute_all`:

``` python
from exported_series import (
    make_context,
    set_borvelia_primary_balance,
    compute_borvelia_primary_balance,
)

ctx = make_context()
set_borvelia_primary_balance(ctx, [{"TIME_PERIOD": 4, "OBS_VALUE": 7.5}])
records = compute_borvelia_primary_balance(ctx=ctx)
```
