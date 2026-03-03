## `excel-grapher`

> **Proprietary software** — © 2026 Teal Insights. All rights reserved.

Build and analyze dependency graphs from Excel workbooks.

### Why this exists

- **Transpilation support**: trace formula dependencies to enable Excel → Python translation
- **Interpretability**: visualize and sanity-check spreadsheet logic (GraphViz, Mermaid, NetworkX)
- **Performance-minded**: focuses on targeted dependency closure from specific output cells/ranges

### Key design decisions

- **Node identity**: nodes are keyed by sheet-qualified A1 strings like `Sheet1!A1` (`NodeKey`)
- **Edge direction**: an edge `A -> B` means **A depends on B** (dependency-first evaluation)
- **Leaf definition**: a leaf is any non-formula cell (`Node.is_leaf=True`)
- **Values are optional**: `load_values=True` loads cached Excel results (second workbook load); otherwise formula nodes have `value=None`
- **Extensible metadata**: each `Node` has a `metadata: dict[str, Any]` that hooks can mutate
- **Range expansion**: ranges like `A1:A10` are expanded to individual cell dependencies (bounded by `max_range_cells`)
- **Normalized formulas**: each formula node has a `normalized_formula` field with sheet-qualified refs, resolved named ranges, and stripped `$` markers — ready for transpilation

### Installation

This is a proprietary package. Install from the private GitHub repository:

**Using uv (recommended):**

```bash
# Basic install
uv add git+https://github.com/Teal-Insights/excel-grapher

# With NetworkX support
uv add "excel-grapher[networkx] @ git+https://github.com/Teal-Insights/excel-grapher"

# With all optional dependencies
uv add "excel-grapher[all] @ git+https://github.com/Teal-Insights/excel-grapher"
```

**Using pip:**

```bash
pip install git+https://github.com/Teal-Insights/excel-grapher

# With extras:
pip install "excel-grapher[networkx] @ git+https://github.com/Teal-Insights/excel-grapher"
```

> **Note:** You must have access to the Teal-Insights GitHub organization and appropriate SSH keys or tokens configured.

### Quick start

```python
from excel_grapher import create_dependency_graph, to_graphviz

g = create_dependency_graph("model.xlsx", ["Sheet1!A10"], load_values=False)
print(len(g))  # number of visited nodes
print(to_graphviz(g))
```

### Exports

- **GraphViz DOT**:

```python
from excel_grapher import to_graphviz
dot = to_graphviz(g, rankdir="LR")
```

- **Mermaid**:

```python
from excel_grapher import to_mermaid
mm = to_mermaid(g, max_nodes=100)
```

- **NetworkX** (optional dependency):

```python
from excel_grapher import to_networkx
G = to_networkx(g)
```

### Validation (calcChain.xml)

```python
from excel_grapher import validate_graph

res = validate_graph(g, Path("model.xlsx"), scope={"Sheet1"})
print(res.is_valid, res.messages)
```

If `xl/calcChain.xml` is missing (common for generated files), validation returns `is_valid=True` with an informational message.

### Working with cell data (for transpilation)

The `DependencyGraph` provides direct O(1) access to cell data via `get_node()`, plus filter methods for iterating over formula vs leaf cells.

```python
from pathlib import Path
from excel_grapher import create_dependency_graph, discover_formula_cells_in_rows

# Discover formula cells in specific rows
targets = discover_formula_cells_in_rows(Path("model.xlsx"), "Sheet1", [10, 11, 12])

# Build the dependency graph
graph = create_dependency_graph(Path("model.xlsx"), targets, load_values=True)

# Access cells by normalized address (O(1) lookup)
node = graph.get_node("Sheet1!A10")
print(node.formula)             # Original formula
print(node.normalized_formula)  # Sheet-qualified for transpilation
print(node.value)               # Cached value from Excel

# Iterate over formula cells
for key, node in graph.formula_nodes():
    print(key, node.normalized_formula)

# Iterate over leaf (value) cells
for key, node in graph.leaf_node_items():
    print(key, node.value)

# Get sorted keys
formula_keys = graph.formula_keys()
leaf_keys = graph.leaf_keys()
```

#### `DependencyGraph` filter methods

| Method | Returns | Description |
|--------|---------|-------------|
| `get_node(key)` | `Node \| None` | O(1) lookup by cell address |
| `formula_nodes()` | `Iterator[tuple[NodeKey, Node]]` | Cells with formulas |
| `leaf_node_items()` | `Iterator[tuple[NodeKey, Node]]` | Leaf cells (no formula) |
| `formula_keys()` | `list[NodeKey]` | Sorted keys for formula cells |
| `leaf_keys()` | `list[NodeKey]` | Sorted keys for leaf cells |

#### `Node` fields

| Field | Type | Description |
|-------|------|-------------|
| `formula` | `str \| None` | Original formula (None for leaf cells) |
| `normalized_formula` | `str \| None` | Sheet-qualified formula for transpilation |
| `value` | `Any` | Cached or hardcoded value |
| `is_leaf` | `bool` | True if cell has no formula |
| `sheet` | `str` | Sheet name |
| `column` | `str` | Column letter |
| `row` | `int` | Row number |

#### `discover_formula_cells_in_rows()`

Utility for scanning rows to find formula cells with numeric cached values:

```python
def discover_formula_cells_in_rows(
    wb_path: Path,
    sheet_name: str,
    rows: list[int],
) -> list[str]
```

Returns sheet-qualified cell addresses (e.g., `"'Sheet Name'!A1"`) for formula cells.

# `excel-evaluator`

## Objective

Create a Python library that can expand Excel formulas by substitution and translate or transpile to optimized, interpretable, functionally equivalent Python code.

## Context

We have an Excel workbook, `example/data/lic-dsf-template.xlsm`, that will serve as the primary target for the library (though the intent is to build a general-purpose tool.) Within the workbook, we are targeting particular indicator rows for conversion to Python code: "B1_GDP_ext" rows 35, 36, 39, 40; "B3_Exports_ext" rows 35, 36, 39, 40; "B4_other flows_ext" rows 35, 36, 39, 40.

The `excel_grapher` dependency implements logic to extract all cells required for the computation of these indicators as an `excel_grapher.DependencyGraph` object. This object already represents Excel addresses in a normalized format, with all addresses sheet-qualified and all named ranges expanded to their canonical references. We control this dependency, so any bugs related to graph extraction should be reported and fixed there.

When executing the computation(s), we want to respect Excel's internal logic. We cache the computed value for each cell so that if the cell is referenced twice in the same formula or in different branches of the call stack, we can return the cached value instead of recalculating it. The principle is, "compute a cell only once." For purposes of this project, we can assume `iterate=False` and any cycles that occur are due to implementation errors on our part.

For function implementations, we borrow heavily from the `formulas` library. For notes on `formulas`, see the [project wiki](https://github.com/Teal-Insights/excel-formula-expander/wiki/Notes-on-the-architecture-of-the-%60formulas%60-library).

## Demo

### What we’re demonstrating

This demo builds a **synthetic two-cell workbook**:

- `S!A1` is a leaf value (`10`)
- `S!B1` is a formula (`=A1*2`) that references `S!A1`

We then run:

- **`FormulaEvaluator`**: interprets a `DependencyGraph` (graph-driven
  runtime evaluation)
- **`CodeGenerator`**: emits standalone Python where dependencies are
  encoded as function calls

### Setup

For demonstration purposes, this code block creates a workbook with two
cells, where one references the other.

``` python
from __future__ import annotations

import sys
from pathlib import Path

import openpyxl


# Setup (create workbook, run demo)
def _find_repo_root(start: Path) -> Path:
    for p in [start, *start.parents]:
        if (p / "pyproject.toml").exists():
            return p
    raise RuntimeError("Could not find repo root (missing pyproject.toml)")


def create_synthetic_workbook(path: Path, *, sheet_name: str = "S") -> None:
    """Create a tiny workbook with two cells where one references the other."""
    path.parent.mkdir(parents=True, exist_ok=True)

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = sheet_name

    ws["A1"].value = 10
    ws["B1"].value = "=A1*2"

    wb.save(path)


ROOT = _find_repo_root(Path.cwd())
sys.path.insert(0, str(ROOT))

workbook_path = ROOT / "demo" / "_artifacts" / "two_cell_demo.xlsx"
create_synthetic_workbook(workbook_path, sheet_name="S")
```

### DependencyGraph (dict representation)

When processing an Excel workbook, we start with a target cell (or
cells) and traverse the formula dependencies to build a
`DependencyGraph` (by which we really just mean a flat dictionary of
“nodes”/cells and a separate dictionary of “edges” to show the
dependencies between them).

``` python
import json
from dataclasses import asdict
from excel_grapher import create_dependency_graph
from excel_grapher import DependencyGraph

from excel_evaluator import FormulaEvaluator
from excel_evaluator.codegen import CodeGenerator

targets = ["S!B1"]
graph = create_dependency_graph(
    workbook_path,
    targets,
    load_values=True,
    max_depth=10,
)

def serialize_graph(graph: DependencyGraph) -> dict:
    return {
        "nodes": {k: asdict(v) for k, v in graph._nodes.items()},
        # Adjacency list: node -> dependencies (edges point from node to its deps)
        "edges": {k: sorted(v) for k, v in graph._edges.items()}
    }

print(json.dumps(serialize_graph(graph), indent=2, sort_keys=True))
```

    {
      "edges": {
        "S!A1": [],
        "S!B1": [
          "S!A1"
        ]
      },
      "nodes": {
        "S!A1": {
          "column": "A",
          "formula": null,
          "is_leaf": true,
          "metadata": {},
          "normalized_formula": null,
          "row": 1,
          "sheet": "S",
          "value": 10
        },
        "S!B1": {
          "column": "B",
          "formula": "=A1*2",
          "is_leaf": false,
          "metadata": {},
          "normalized_formula": "=S!A1*2",
          "row": 1,
          "sheet": "S",
          "value": null
        }
      }
    }

### Evaluator results

The `FormulaEvaluator` is a wrapper around the `DependencyGraph` that
provides translation and evaluation logic to turn Excel formulas into
equivalent Python code and run them. It has, for instance, a dictionary
of Python equivalents for Excel functions, operators, and error types,
as well as circular-reference handling (Excel’s default: warn + return 0)
and a cache to enforce that we only evaluate a cell once. It’s
essentially a full Excel emulator in Python.

In practice, this provides very fast, accurate, and repeatable
extraction/execution for any given Excel workbook, but the actual logic
of the computation is non-transparent and non-modular, plus we would
have to leak our Excel emulation engine in order for end-users to run
the extracted workbook formulas using this engine.

An advantage of the `FormulaEvaluator` is that it provides a native
interface for extraction, so it’s super easy to re-extract and re-run
the computation if users update their workbook. This doesn’t even
require that they respect the original template workbook’s structure,
because the re-extraction will automatically re-map the formula
hierarchy of the altered workbook. If we want to remain tightly coupled
to Excel and support user workbook import, we’d probably want to stick
with this approach.

A disadvantage is that we are translating Excel formulas to Python at
runtime, which adds a bit of computational overhead vs. translating them
to Python ahead of time.

``` python
with FormulaEvaluator(graph) as ev:
    evaluator_results = ev.evaluate(targets)
print(evaluator_results)
```

    {'S!B1': 20.0}

### Exported code

The code we export today has a different set of tradeoffs:

- **Advantages**
  - **Standalone artifact**: the output is plain Python that can be
    shipped to end-users without bundling `excel_evaluator` or
    `excel_grapher`.
  - **Partial obfuscation**: we are not exposing our Excel
    parsing/extraction engine directly; while a motivated reader might
    infer pieces of the approach from the emitted structure, it’s
    materially less transparent than shipping the evaluator/extractor.
  - **Minimal runtime surface**: we embed **only** the Excel-equivalent
    `xl_*` functions/operators/helpers that are actually needed by the
    exported dependency graph (plus a small amount of scaffolding).
  - **Repeatable execution**: the exported code “freezes” the workbook
    logic at a point in time; downstream runs are deterministic and
    don’t depend on Excel.
- **Disadvantages**
  - **Still Excel-shaped**: the output is still fundamentally organized
    around cell-level functions and Excel semantics, so interpretability
    and modularity are only modestly improved (it’s more a foundation
    for later refactors than a final “interpretable model”).
  - **Regeneration required**: if the workbook changes, you need to
    re-extract and re-export; the code is not a live view over a mutable
    workbook.

``` python
code = CodeGenerator(graph).generate(targets)
print("\n".join(code.splitlines()[:120]))
```

You can also emit named entrypoints by passing a mapping of names to target lists:

``` python
code = CodeGenerator(graph).generate(
    targets,
    entrypoints={
        "outputs": ["S!B1", "S!C1"],
        "checks": ["S!D1"],
    },
)
```

This generates `compute_outputs(...)` and `compute_checks(...)` alongside `compute_all(...)`.

    """Standalone runtime for generated Excel formula code."""

    from __future__ import annotations

    from enum import Enum


    class XlError(str, Enum):
        """Excel error values."""
        VALUE = "#VALUE!"
        REF = "#REF!"
        DIV = "#DIV/0!"
        NA = "#N/A"
        NAME = "#NAME?"
        NUM = "#NUM!"
        NULL = "#NULL!"


    def to_number(value) -> float | XlError:
        if value is None:
            return 0.0
        if isinstance(value, XlError):
            return value
        if isinstance(value, bool):
            return 1.0 if value else 0.0
        if isinstance(value, (int, float)):
            return float(value)
        if isinstance(value, str):
            s = value.strip()
            if s == "":
                return 0.0
            try:
                return float(s)
            except ValueError:
                return XlError.VALUE
        return XlError.VALUE


    def xl_mul(left, right) -> float | XlError:
        """Safe multiplication that propagates XlError values."""
        if isinstance(left, XlError):
            return left
        if isinstance(right, XlError):
            return right
        left_num = to_number(left)
        right_num = to_number(right)
        if isinstance(left_num, XlError):
            return left_num
        if isinstance(right_num, XlError):
            return right_num
        return left_num * right_num

    from functools import lru_cache


    # --- Cell functions ---

    @lru_cache(maxsize=None)
    def cell_s_a1():
        """Leaf cell: S!A1"""
        return 10


    @lru_cache(maxsize=None)
    def cell_s_b1():
        """Formula: =A1*2"""
        return xl_mul(cell_s_a1(), 2.0)


    def compute_all():
        """Compute all target cells and return results."""
        return {
            'S!B1': cell_s_b1(),
        }

### Exported code results

When run, the exported code produces the same results as the
`FormulaEvaluator`:

``` python
namespace: dict = {}
exec(code, namespace)
generated_results = namespace["compute_all"]()
print(generated_results)
```

    {'S!B1': 20.0}

## Roadmap

- Eventually we'll want to modularize or even vectorize parts of the computation in the exported code. Static analysis might provide insight into natural opportunities for this. Modularizable parts of the code will have a pyramidal shape, with either a single-cell entrypoint or a group of entrypoints that all belong to the same row or range. Perhaps we look at cells that are often referenced as ranges, and see if we can vectorize them. And perhaps we look at the top_n most frequently referenced cells as natural candidates for modularization, since frequently referenced cells are likely to represent core abstractions. This is a longer term goal, aimed at producing a "more interpretable" Python representation of the economic model (as opposed to using a dict of Excel-addressed formulas that don't make much sense when abstracted from the workbook.)
- To inform the item above, we will want to extract the row and column labels from the Excel workbook for each referenced cell, so that we can store them in our data model. I've made a start on this in the [lic-dsf-programmatic-extraction](https://github.com/Teal-Insights/lic-dsf-programmatic-extraction) repo.