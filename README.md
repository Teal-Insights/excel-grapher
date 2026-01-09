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

