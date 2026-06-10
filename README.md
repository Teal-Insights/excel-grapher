## `excel-grapher`

Build and analyze dependency graphs from Excel workbooks, **evaluate formulas with Excel semantics**, and **export standalone Python code**.

**Documentation:** [https://teal-insights.github.io/excel-grapher/](https://teal-insights.github.io/excel-grapher/)

### Why this exists

- **Transpilation support**: trace formula dependencies to enable Excel → Python translation.
- **Interpretability**: visualize and sanity-check spreadsheet logic (GraphViz, Mermaid, NetworkX).
- **Performance-minded**: focuses on targeted dependency closure from specific output cells/ranges.
- **Excel semantics in Python**: run workbook logic in-process with a full Excel-like evaluator.
- **Exportable**: emit standalone Python packages that embed only the runtime surface you need.

---

### Library layout

The unified distribution is `excel-grapher` and exposes a single import package, `excel_grapher`, with five main subpackages:

- `excel_grapher/core/` — shared semantic types, coercions, and scalar operators.
- `excel_grapher/runtime/` — Excel-equivalent function implementations and runtime semantics.
- `excel_grapher/grapher/` — workbook loading, graph extraction, and visualization logic.
- `excel_grapher/evaluator/` — `FormulaEvaluator`: an Excel emulator for recomputing formulas in the extracted graph in Python.
- `excel_grapher/exporter/` — `CodeGenerator`: an transpiler for exporting the extracted graph as a standalone Python library.

Typical imports:

```python
from excel_grapher.grapher import create_dependency_graph, DependencyGraph
from excel_grapher.evaluator import FormulaEvaluator
from excel_grapher.exporter import CodeGenerator
from excel_grapher.core import XlError  # and other shared types, if needed
```

---

### Installation

This is a proprietary package. Install from the private GitHub repository:

**Using `uv` (recommended):**

```bash
# Basic install
uv add git+https://github.com/Teal-Insights/excel-grapher

# With NetworkX support
uv add "excel-grapher[networkx] @ git+https://github.com/Teal-Insights/excel-grapher"

# With all optional dependencies
uv add "excel-grapher[all] @ git+https://github.com/Teal-Insights/excel-grapher"
```

**Using `pip`:**

```bash
pip install git+https://github.com/Teal-Insights/excel-grapher

# With extras:
pip install "excel-grapher[networkx] @ git+https://github.com/Teal-Insights/excel-grapher"
```

> **Note:** You must have access to the Teal-Insights GitHub organization and appropriate SSH keys or tokens configured.

---

### High-level usage

The library supports a three-stage pipeline:

1. **Build a dependency graph** from an Excel workbook (`excel_grapher.grapher`).
2. **Evaluate formulas with Excel semantics** over that graph (`excel_grapher.evaluator.FormulaEvaluator`).
3. **Export standalone Python code** that embeds only the required runtime surface (`excel_grapher.exporter.CodeGenerator`).

A minimal end‑to‑end example:

```python
from pathlib import Path

from excel_grapher.grapher import create_dependency_graph
from excel_grapher.evaluator import FormulaEvaluator
from excel_grapher.exporter import CodeGenerator

workbook_path = Path("model.xlsx")
targets = ["Sheet1!A10"]

# 1) Build a dependency graph
graph = create_dependency_graph(workbook_path, targets, load_values=True)
print(len(graph))  # number of visited nodes

# 2) Evaluate with Excel semantics
with FormulaEvaluator(graph) as ev:
    results = ev.evaluate(targets)

# 3) Export standalone Python code
code = CodeGenerator(graph).generate(targets)
```

### Series bindings

Optional sidecar manifests (`.bindings.yaml`) declare structured input/output APIs for **exported**
code — `set_*` and `compute_*` functions over `Records`. See the
[Series bindings guide](https://teal-insights.github.io/excel-grapher/user-guide/series-bindings.html)
and [Code export](https://teal-insights.github.io/excel-grapher/user-guide/export.html) guide.

### User guide

Detailed documentation lives in the [User Guide](https://teal-insights.github.io/excel-grapher/user-guide/dependency-graphs.html):

| Topic | Page |
| --- | --- |
| Dependency graphs | [Read guide](https://teal-insights.github.io/excel-grapher/user-guide/dependency-graphs.html) |
| Visualization | [Read guide](https://teal-insights.github.io/excel-grapher/user-guide/visualization.html) |
| Formula evaluation | [Read guide](https://teal-insights.github.io/excel-grapher/user-guide/evaluator.html) |
| End-to-end demo | [Read guide](https://teal-insights.github.io/excel-grapher/user-guide/end-to-end-demo.html) |
| Series bindings | [Read guide](https://teal-insights.github.io/excel-grapher/user-guide/series-bindings.html) |
| Code export | [Read guide](https://teal-insights.github.io/excel-grapher/user-guide/export.html) |
| Parity testing | [Read guide](https://teal-insights.github.io/excel-grapher/user-guide/parity-testing.html) |
| Contributing | [Read guide](https://teal-insights.github.io/excel-grapher/user-guide/contributing.html) |

### Examples

Hands-on walkthroughs (Markdown on GitHub):

| Topic | Walkthrough |
| --- | --- |
| Graph extraction | [extraction_basics.md](https://github.com/Teal-Insights/excel-grapher/blob/main/examples/micro_workbooks/extraction_basics.md) |
| Code generation | [codegen_basics.md](https://github.com/Teal-Insights/excel-grapher/blob/main/examples/micro_workbooks/codegen_basics.md) |
| Series bindings | [series_bindings.md](https://github.com/Teal-Insights/excel-grapher/blob/main/examples/micro_workbooks/series_bindings.md) |
| Induced graph | [induced_graph.md](https://github.com/Teal-Insights/excel-grapher/blob/main/examples/induced_graph/induced_graph.md) |
