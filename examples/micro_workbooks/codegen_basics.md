# Consolidated Micro-Workbook Basic Code Generation Examples


Each row of
[examples/micro_workbooks/codegen_basics.xlsx](codegen_basics.xlsx)
contains a self-contained example that can be extracted as a graph and
then exported to standalone Python code. This workbook demonstrates the
workflow and application behavior for different Excel dependency
scenarios.

``` python
from pathlib import Path

from excel_grapher.grapher import (
    create_dependency_graph, DependencyGraph
)
from excel_grapher.exporter import CodeGenerator

# Load the example workbook
workbook_path = Path("codegen_basics.xlsx")
```

## 01. Formula with no dependencies

The first example is a single-cell formula with no dependencies. We can
extract the graph with the `create_dependency_graph` function (see
[Extraction
Basics](examples/micro_workbooks/extraction_basics.qmd%20for%20more%20details));
then, instead of running the graph using the `FormulaEvaluator` Excel
emulator, we can transpile the graph to standalone Python code using the
`CodeGenerator` class. We’ll write the code to a file called
`formula_with_no_dependencies.py` in the `codegen_outputs` folder.

``` python
graph: DependencyGraph = create_dependency_graph(
    workbook_path,
    ["Sheet1!B1"]
)

with CodeGenerator(graph) as gen:
    code = gen.generate()
with open("codegen_outputs/formula_with_no_dependencies.py", "w", encoding="utf-8") as f:
    f.write(code)
```

Note that `CodeGenerator`’s `generate` method exports a miniature Excel
runtime, with error handling and formula/operator implementations for
the Excel functions used in the graph. So while the implementation of
our `1+1` function on line 344-346 is brief, the full code output is
nearly 400 lines of code, which feels excessive. There is enormous room
to optimize this code generation process to reduce the final output
length.

We can then append the file as a module to our current session and run
the code:

``` python
import sys
sys.path.append("codegen_outputs")
from formula_with_no_dependencies import compute_all

result = compute_all()
print(f"```text\n{result}\n```")
```

``` text
{'Sheet1!B1': 2.0}
```

The return value of the `compute_all` function is a dictionary of the
target cell address and its computed value.

## 02. Linear dependency

The second example consists of two cells: one hardcoded (“Sheet1!B2”)
and one a formula that depends on the hardcoded cell (“Sheet1!C2”).
Let’s generate the code for this example and run it with the default
input:

``` python
graph: DependencyGraph = create_dependency_graph(
    workbook_path,
    ["Sheet1!C2"]
)

with CodeGenerator(graph) as gen:
    code = gen.generate()
with open("codegen_outputs/linear_dependency.py", "w", encoding="utf-8") as f:
    f.write(code)

from linear_dependency import compute_all

result = compute_all()
print(f"```text\n{str(result['Sheet1!C2'])}\n```")
```

``` text
2.0
```

To recompute the value of “Sheet1!C2” with a different input value for
“Sheet1!B2”, we can create a `context` object with our desired inputs
and call `compute_all` with it:

``` python
from linear_dependency import make_context

context = make_context(inputs={"Sheet1!B2": 2})
result = compute_all(ctx=context)
print(f"```text\n{str(result['Sheet1!C2'])}\n```")
```

``` text
3.0
```

The implementation of the C2 function is defined on lines 345-347. The
hardcoded input value for B2 is set in lines 338-340. Note that the
exported code is object-oriented rather than functional, with inputs and
computation caching stored in a mutable `Context` object, so you must
take care not to share the same `Context` instance if running multiple
scenarios in parallel in the same session.

## 03. Multiple non-adjacent targets

If there are multiple target cells that are not adjacent to each other,
`compute_all` simply returns a dictionary keyed by cell address for each
target cell.

``` python
graph: DependencyGraph = create_dependency_graph(
    workbook_path,
    ["Sheet1!C3", "Sheet1!E3"]
)

with CodeGenerator(graph) as gen:
    code = gen.generate()
with open("codegen_outputs/multiple_non_adjacent_targets.py", "w", encoding="utf-8") as f:
    f.write(code)

from multiple_non_adjacent_targets import compute_all

result = compute_all()
print(f"```text\n{str(result)}\n```")
```

``` text
{'Sheet1!C3': 2.0, 'Sheet1!E3': 3.0}
```

An unordered dictionary does seem like the right output shape for
`compute_all` outputs representing non-adjacent target cells, though I’d
like to be able to override/customize the naming of the keys.

`excel-grapher` currently supports a `entrypoints` argument to
`CodeGenerator.generate` that allows you to define explicit entrypoints
for the target cells, which causes `compute_*` functions to be generated
for each entrypoint name:

``` python
from typing import Mapping, Sequence

entrypoints: Mapping[str, Sequence[str]] | None = {
    "c3_cell": ["Sheet1!C3"],
    "e3_cell": ["Sheet1!E3"]
}

with CodeGenerator(graph) as gen:
    code = gen.generate(entrypoints=entrypoints)
with open("codegen_outputs/multiple_non_adjacent_targets_with_entrypoints.py", "w", encoding="utf-8") as f:
    f.write(code)

from multiple_non_adjacent_targets_with_entrypoints import (
    compute_all,
    compute_c3_cell,
    compute_e3_cell
)

import re

compute_c3_cell_match = re.search(
    r"def compute_c3_cell\([\s\S]*?(?=^def |\Z)",
    code,
    re.MULTILINE,
)

print(f"```python\n{compute_c3_cell_match.group(0).strip()}\n```")

c3_result = compute_c3_cell()
e3_result = compute_e3_cell()
print(f"```text\n{str(c3_result)}\n```")
print(f"```text\n{str(e3_result)}\n```")
```

``` python
def compute_c3_cell(inputs=None, *, ctx=None):
    """Compute c3_cell target cells and return results."""
    if ctx is None:
        ctx = make_context(inputs)
    elif inputs is not None:
        warnings.warn("inputs will be ignored because ctx was provided", UserWarning, stacklevel=2)
    return {target: handler(ctx, target) for target, handler in TARGETS_C3_CELL.items()}


TARGETS_E3_CELL = {
    'Sheet1!E3': xl_cell,
}
```

``` text
{'Sheet1!C3': 2.0}
```

``` text
{'Sheet1!E3': 3.0}
```

But this doesn’t currently affect the shape of the dictionary returned
by `compute_all`, and maybe it should:

``` python
result = compute_all()
print(f"```text\n{str(result)}\n```")
```

``` text
{'Sheet1!C3': 2.0, 'Sheet1!E3': 3.0}
```

## 04. Multiple adjacent targets

The next example demonstrates what happens when we export the code with
multiple adjacent targets.

``` python
graph: DependencyGraph = create_dependency_graph(
    workbook_path,
    ["Sheet1!C4", "Sheet1!D4"]
)

with CodeGenerator(graph) as gen:
    code = gen.generate()
with open("codegen_outputs/multiple_adjacent_targets.py", "w", encoding="utf-8") as f:
    f.write(code)

from multiple_adjacent_targets import compute_all

result = compute_all()
print(f"```text\n{str(result)}\n```")
```

``` text
{'Sheet1!C4:Sheet1!D4': array([[2.0, 3.0]], dtype=object)}
```

Here, the adjacent cell addresses are automatically aggregated into a
single Excel *range address*, which is used as the key in the returned
dictionary. The value for that key is a NumPy array of the computed
values for each cell in the range.

This is probably a good default in *most* cases, because contiguous
cells will often comprise a logical unit. However, I would like to be
able to override this behavior by defining explicit entrypoints for the
target cells, which currently has no effect on the shape of the
dictionary returned by `compute_all`:

``` python
from typing import Mapping, Sequence

entrypoints: Mapping[str, Sequence[str]] | None = {
    "c4_d4_range": ["Sheet1!C4", "Sheet1!D4"],
    "c4_cell": ["Sheet1!C4"]
}

with CodeGenerator(graph) as gen:
    code = gen.generate(entrypoints=entrypoints)
with open("codegen_outputs/multiple_adjacent_targets_with_entrypoints.py", "w", encoding="utf-8") as f:
    f.write(code)

from multiple_adjacent_targets_with_entrypoints import (
    compute_all,
    compute_c4_cell,
    compute_c4_d4_range
)

result = compute_all()
print(f"```text\n{str(result)}\n```")
```

``` text
{'Sheet1!C4:Sheet1!D4': array([[2.0, 3.0]], dtype=object)}
```

Note that there is no obstacle to defining entrypoints with overlapping
cell addresses. Here we define an entrypoint for one of the individual
cells as well as one for the range that contains it.

``` python
c4_result = compute_c4_cell()
c4_d4_result = compute_c4_d4_range()
print(f"```text\nCell result: {str(c4_result)}\nRange result: {str(c4_d4_result)}\n```")
```

``` text
Cell result: {'Sheet1!C4': 2.0}
Range result: {'Sheet1!C4:Sheet1!D4': array([[2.0, 3.0]], dtype=object)}
```

While I think it’s fine that `compute_all` returns a `dict` keyed by
cell address, I am skeptical of that return value shape for the
entrypoint functions. Here we are only returning a single value or
series, so we shouldn’t need keyed access. Instead, maybe we return a
scalar value or tuple. The one downside of this is that it makes it
harder to associate, for example, values in a time series with years. A
dict is not the right shape for a time series, because a dict is
unordered, but neither does it seem quite right to have a tuple with no
labels. A tuple of tuples, tuple/list of dicts, or array may be clearer
options.

Here are a few minimal sketch options for entrypoint signatures and
output shapes for an economic time series:

``` python
# Option A: tuple[tuple[int, float], ...] for immutable (year, value) pairs
def compute_gdp_series(...) -> tuple[tuple[int, float], ...]:
    return ((2021, 23115.0), (2022, 25440.0), (2023, 27361.0))

# Option B: tuple[dict[str, float], ...] for per-point labels
def compute_gdp_series(...) -> tuple[dict[str, int | float], ...]:
    return (
        {"year": 2021, "gdp": 23115.0},
        {"year": 2022, "gdp": 25440.0},
        {"year": 2023, "gdp": 27361.0},
    )

# Option C: list[dict[str, float]] for JSON-friendly downstream use
def compute_gdp_series(...) -> list[dict[str, int | float]]:
    return [
        {"year": 2021, "gdp": 23115.0},
        {"year": 2022, "gdp": 25440.0},
        {"year": 2023, "gdp": 27361.0},
    ]

# Option D: tuple[float, ...] for strictly positional values
def compute_gdp_values_tuple(...) -> tuple[float, ...]:
    # index 0 -> 2021, index 1 -> 2022, index 2 -> 2023
    return (23115.0, 25440.0, 27361.0)

# Option E: list[float] for strictly positional values
def compute_gdp_values_list(...) -> list[float]:
    # index position implies year based on documented ordering
    return [23115.0, 25440.0, 27361.0]

# Option F: NumPy array for contiguous/range-like outputs
import numpy as np

def compute_gdp_series_array(...) -> np.ndarray:
    # shape: (n, 2), each row is [year, gdp]
    return np.asarray([[2021.0, 23115.0], [2022.0, 25440.0], [2023.0, 27361.0]])
```

## 05. Must cycle

In the fourth example, The B5 and C5 formula cells make a cycle. Excel’s
internal behavior with respect to cycles is different depending on
workbook settings. If `iterate` is enabled in the workbook, Excel will
iterate over the cycle until it converges on a value or hits a maximum
number of iterations. Otherwise, it will stop and return 0 from any cell
already seen before in a formula chain. Like `FormulaEvaluator`,
standalone Python code generated and exported from `CodeGenerator`
replicates this behavior with a `CircularReferenceWarning` unless the
workbook is configured to allow cycles:

``` python
graph: DependencyGraph = create_dependency_graph(
    workbook_path,
    ["Sheet1!B5", "Sheet1!C5"]
)

with CodeGenerator(graph) as gen:
    code = gen.generate()
with open("codegen_outputs/must_cycle.py", "w", encoding="utf-8") as f:
    f.write(code)

from must_cycle import compute_all

result = compute_all()
print(f"```text\n{str(result)}\n```")
```

``` text
{'Sheet1!B5:Sheet1!C5': array([[2.0, 1.0]], dtype=object)}
```
