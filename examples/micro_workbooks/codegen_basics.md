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
graph: DependencyGraph = create_dependency_graph(workbook_path, ["Sheet1!B1"], load_values=False)

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
    workbook_path, ["Sheet1!C2"], load_values=True)

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

## 03. Multiple targets

The third example demonstrates what happens when we export the code with
multiple targets.

``` python
graph: DependencyGraph = create_dependency_graph(workbook_path, ["Sheet1!C3", "Sheet1!D3"], load_values=False)

with CodeGenerator(graph) as gen:
    code = gen.generate()
with open("codegen_outputs/multiple_targets.py", "w", encoding="utf-8") as f:
    f.write(code)

from multiple_targets import compute_all

result = compute_all()
print(f"```text\n{str(result)}\n```")
```

``` text
{'Sheet1!C3:Sheet1!D3': array([[2.0, 3.0]], dtype=object)}
```

Here, the output is a dictionary keyed by the Excel range address of the
target cells, and the value is a NumPy array of the computed values for
each cell in the range.

This is pretty unergonomic. At minimum, I think we should return a
dictionary keyed by cells, not by the aggregated range. But I don’t love
that dictionaries are unordered, so maybe we should return a list or
namedtuple or something ordered instead. We also need to think about how
we want to transform the code in postprocessing, and/or what output
options to support. (E.g., what if we want to key by labels instead of
addresses? But if we use a namedtuple, we can’t use numeric labels
because property names can’t start with numbers.)

## 04. Must cycle

In the fourth example, The B4 and C4 formula cells make a cycle. Excel’s
internal behavior with respect to cycles is different depending on
workbook settings. If `iterate` is enabled in the workbook, Excel will
iterate over the cycle until it converges on a value or hits a maximum
number of iterations. Otherwise, it will stop and return 0 from any cell
already seen before in a formula chain. Like `FormulaEvaluator`,
standalone Python code generated and exported from `CodeGenerator`
replicates this behavior with a `CircularReferenceWarning` unless the
workbook is configured to allow cycles:

``` python
graph: DependencyGraph = create_dependency_graph(workbook_path, ["Sheet1!B4", "Sheet1!C4"], load_values=False)

with CodeGenerator(graph) as gen:
    code = gen.generate()
with open("codegen_outputs/must_cycle.py", "w", encoding="utf-8") as f:
    f.write(code)

from must_cycle import compute_all

result = compute_all()
print(f"```text\n{str(result)}\n```")
```

``` text
{'Sheet1!B4:Sheet1!C4': array([[2.0, 1.0]], dtype=object)}
```

    C:\Users\chris\Software\excel-grapher\examples\micro_workbooks\codegen_outputs\must_cycle.py:295: CircularReferenceWarning: Circular reference detected; returning 0 (iterative calculation is disabled).
      return xl_circular_reference()
