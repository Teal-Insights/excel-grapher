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

### Extractor (flattening for transpilation)

The extractor module provides a simplified interface for extracting cell data into flat dictionaries, ready for formula expansion and Python transpilation.

```python
from pathlib import Path
from excel_grapher import build_cell_dict

# Define which sheets/rows contain your output formulas
sheet_rows = {
    "Sheet1": [10, 11, 12],
    "Sheet2": [5, 6],
}

# Build the cell dictionary (traces all dependencies)
cells = build_cell_dict(Path("model.xlsx"), sheet_rows, load_values=True)

# Access cells by normalized address
cell = cells["Sheet1!A10"]
print(cell.formula)             # Original formula
print(cell.normalized_formula)  # Sheet-qualified for transpilation
print(cell.value)               # Cached value from Excel

# Filter by cell type
formula_cells = cells.formula_cells()
value_cells = cells.value_cells()

# Get sorted keys
for key in cells.formula_keys():
    print(key, cells[key].normalized_formula)
```

#### `CellInfo`

Dataclass representing a single cell:

| Field | Type | Description |
|-------|------|-------------|
| `formula` | `str \| None` | Original formula (None for value cells) |
| `normalized_formula` | `str \| None` | Sheet-qualified formula for transpilation |
| `value` | `Any` | Cached or hardcoded value |
| `is_formula` | `bool` (property) | True if cell contains a formula |

#### `CellDict`

Dictionary subclass (`dict[str, CellInfo]`) with helper methods:

| Method | Returns | Description |
|--------|---------|-------------|
| `formula_cells()` | `dict[str, CellInfo]` | Only cells with formulas |
| `value_cells()` | `dict[str, CellInfo]` | Only hardcoded value cells |
| `formula_keys()` | `list[str]` | Sorted keys for formula cells |
| `value_keys()` | `list[str]` | Sorted keys for value cells |

#### `build_cell_dict()`

Main entry point for building a cell dictionary:

```python
def build_cell_dict(
    workbook_path: Path,
    sheet_rows: dict[str, list[int]],
    load_values: bool = True,
    max_depth: int = 50,
) -> CellDict
```

| Parameter | Description |
|-----------|-------------|
| `workbook_path` | Path to the Excel file |
| `sheet_rows` | Dict mapping sheet names to output row numbers |
| `load_values` | Whether to load cached values (default: True) |
| `max_depth` | Maximum dependency traversal depth (default: 50) |

#### Lower-level functions

- `discover_formula_cells_in_rows(wb_path, sheet_name, rows)` - Scan rows for formula cells with numeric values
- `graph_to_cell_dict(graph)` - Convert a `DependencyGraph` to a `CellDict`

