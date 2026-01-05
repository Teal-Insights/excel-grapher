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

