# Consolidated Micro-Workbook Evaluation Examples


Each row of
[examples/micro_workbooks/codegen_basics.xlsx](codegen_basics.xlsx)
contains a self-contained example that can be extracted as a graph and
evaluated with `FormulaEvaluator`. For a standalone Python
**package** (`compute_*` over named series), provide series bindings
and use `CodeGenerator.generate_modules` — see
[Series bindings](series_bindings.md).

``` python
from pathlib import Path

from excel_grapher.grapher import create_dependency_graph, DependencyGraph
from excel_grapher.evaluator import FormulaEvaluator

# Load the example workbook
workbook_path = Path("codegen_basics.xlsx")
```

## 01. Formula with no dependencies

The first example is a single-cell formula with no dependencies. Extract
the graph with `create_dependency_graph` (see
[Extraction Basics](extraction_basics.md) for more details), then
evaluate it with `FormulaEvaluator`.

``` python
graph: DependencyGraph = create_dependency_graph(workbook_path, ["Sheet1!B1"])

with FormulaEvaluator(graph) as ev:
    result = ev.evaluate(["Sheet1!B1"])
print(f"```text\n{result}\n```")
```

``` text
{'Sheet1!B1': 2.0}
```

The return value is a dictionary of target cell addresses and computed
values.

## 02. Linear dependency

The second example consists of two cells: one hardcoded
(“Sheet1!B2”) and one a formula that depends on the hardcoded cell
(“Sheet1!C2”). Evaluate with the workbook defaults, then change the
input leaf and re-evaluate:

``` python
graph: DependencyGraph = create_dependency_graph(workbook_path, ["Sheet1!C2"])

with FormulaEvaluator(graph) as ev:
    result = ev.evaluate(["Sheet1!C2"])
print(f"```text\n{str(result['Sheet1!C2'])}\n```")
```

``` text
2.0
```

``` python
graph.set_node_value("Sheet1!B2", 2)

with FormulaEvaluator(graph) as ev:
    result = ev.evaluate(["Sheet1!C2"])
print(f"```text\n{str(result['Sheet1!C2'])}\n```")
```

``` text
3.0
```

`set_node_value` mutates the graph. Subsequent evaluations on that graph
see the new leaf value. For a named, records-shaped output API
(dimensions plus `OBS_VALUE`), export an inverted-tree **package** with
series bindings (`generate_modules`) — see
[Series bindings](series_bindings.md).

## 03. Multiple non-adjacent targets

If there are multiple target cells that are not adjacent to each other,
`evaluate` returns a dictionary keyed by cell address for each target
cell.

``` python
graph: DependencyGraph = create_dependency_graph(workbook_path, ["Sheet1!C3", "Sheet1!E3"])

with FormulaEvaluator(graph) as ev:
    result = ev.evaluate(["Sheet1!C3", "Sheet1!E3"])
print(f"```text\n{str(result)}\n```")
```

``` text
{'Sheet1!C3': 2.0, 'Sheet1!E3': 3.0}
```

## 04. Multiple adjacent targets

The next example evaluates two adjacent target cells.
`FormulaEvaluator` still returns one dictionary entry per address (it
does not collapse contiguous cells into a range key).

``` python
graph: DependencyGraph = create_dependency_graph(workbook_path, ["Sheet1!C4", "Sheet1!D4"])

with FormulaEvaluator(graph) as ev:
    result = ev.evaluate(["Sheet1!C4", "Sheet1!D4"])
print(f"```text\n{str(result)}\n```")
```

``` text
{'Sheet1!C4': 2.0, 'Sheet1!D4': 3.0}
```

For stable tabular contracts with named `compute_*` functions that
return tuples (and `as_records` for a Records view), export an
inverted-tree **package** with series bindings — see
[Series bindings](series_bindings.md).

## 05. Must cycle

In this example, the B5 and C5 formula cells make a cycle. Excel’s
internal behavior with respect to cycles is different depending on
workbook settings. If `iterate` is enabled in the workbook, Excel will
iterate over the cycle until it converges on a value or hits a maximum
number of iterations. Otherwise, it will stop and return 0 from any
cell already seen before in a formula chain. `FormulaEvaluator`
replicates this behavior with a `CircularReferenceWarning` unless the
workbook is configured to allow cycles:

``` python
graph: DependencyGraph = create_dependency_graph(workbook_path, ["Sheet1!B5", "Sheet1!C5"])

with FormulaEvaluator(graph) as ev:
    result = ev.evaluate(["Sheet1!B5", "Sheet1!C5"])
print(f"```text\n{str(result)}\n```")
```

``` text
{'Sheet1!B5': 2.0, 'Sheet1!C5': 1.0}
```

    CircularReferenceWarning: Circular reference detected; returning 0 (iterative calculation is disabled).
