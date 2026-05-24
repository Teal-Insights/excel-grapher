## `excel-grapher`

Build and analyze dependency graphs from Excel workbooks, **evaluate formulas with Excel semantics**, and **export standalone Python code**.

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

There are no compatibility shims for the old `excel_evaluator` / standalone `excel_grapher` packages; callers should update to the new import paths.

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

The sections below go into more detail.

---

## 1. Dependency graphs (`excel_grapher.grapher`)

### Key design decisions

- **Node identity**: nodes are keyed by sheet-qualified A1 strings like `Sheet1!A1` (`NodeKey`).
- **Edge direction**: an edge `A -> B` means **A depends on B** (dependency-first evaluation).
- **Leaf definition**: a leaf is any node with no cell dependencies (`Node.is_leaf=True`), including non-formula cells and literal-only formulas (e.g. `=1+1`).
- **Values are optional**: `load_values=True` loads cached Excel results (second workbook load); otherwise formula nodes have `value=None`.
- **Extensible metadata**: each `Node` has a `metadata: dict[str, Any]` that hooks can mutate; optional `label_detection` can populate `row_labels` / `column_labels` at build time.
- **Range expansion**: ranges like `A1:A10` are expanded to individual cell dependencies (bounded by `max_range_cells`).
- **Normalized formulas**: each formula node has a `normalized_formula` field with sheet-qualified refs, resolved named ranges, and stripped `$` markers — ready for transpilation.

### Quick start: building a graph

```python
from pathlib import Path

from excel_grapher.grapher import create_dependency_graph
from excel_grapher.grapher import to_graphviz, to_mermaid, to_networkx  # optional

wb_path = Path("model.xlsx")
targets = ["Sheet1!A10"]

g = create_dependency_graph(wb_path, targets, load_values=False)
print(len(g))          # number of visited nodes
print(to_graphviz(g))  # GraphViz DOT
```

### Optional row/column label detection

When ``label_detection`` is a ``LabelDetectionConfig`` with ``enabled=True``, each node gets
``metadata["row_labels"]`` and ``metadata["column_labels"]`` (lists of strings) at graph build time.
Built-in behaviors include left/up heuristic scans and region-scoped rules
(``left_edge_then_up_scan``, ``top_edge_then_left_scan``, ``region_header_rows``);
pass custom implementations via ``label_behaviors``.

Include the same settings in graph cache ``extraction_params`` using
``label_detection_config_to_jsonable(...)`` so cached graphs invalidate when rules change.

```python
from excel_grapher import LabelDetectionConfig, create_dependency_graph

graph = create_dependency_graph(
    wb_path,
    targets,
    load_values=True,
    label_detection=LabelDetectionConfig(enabled=True),
)
row_labels = graph.get_node(targets[0]).metadata.get("row_labels", [])
```

### Target forms

`targets` accepts any mix of:

- sheet-qualified single cells: `"Sheet1!A1"`, `"'My Sheet'!B2"`
- sheet-qualified ranges: `"Sheet1!B12:F12"`, `"Sheet1!A1:Sheet1!B2"`,
  `"'My Sheet'!A1:B2"`
- defined names that resolve to a single cell or rectangular range:
  `"MyInput"`, `"DataRange"`

Range and named-range targets expand to one root per cell (subject to
`max_range_cells`) and the BFS proceeds from the deduplicated union of roots.
Targets that are neither sheet-qualified nor a known defined name raise
`ValueError`.

### Dynamic OFFSET/INDIRECT configuration

Dynamic references (e.g. `OFFSET`, `INDIRECT`) can be handled in three ways via the `create_dependency_graph` API:

```python
from excel_grapher.grapher import create_dependency_graph, DynamicRefConfig, DynamicRefLimits

# Signature (simplified):
# create_dependency_graph(
#     workbook_path,
#     targets,
#     *,
#     dynamic_refs: DynamicRefConfig | None = None,
#     use_cached_dynamic_refs: bool = False,
#     ...
# )
```

- **`use_cached_dynamic_refs=True`**  
  Use the existing cached-workbook path for `OFFSET`/`INDIRECT`. This preserves the legacy behavior and **ignores** `dynamic_refs`.

- **`use_cached_dynamic_refs=False` (default) and `dynamic_refs is None`**  
  When the builder encounters dynamic refs that require resolution, it raises **`DynamicRefError`**. This is the safe “no silent fallback” default.

- **`use_cached_dynamic_refs=False` and `dynamic_refs is a DynamicRefConfig`**  
  Resolve dynamic refs using static constraints (cell types and limits). Missing or invalid domains still raise `DynamicRefError`.

You typically build a `DynamicRefConfig` from a `dict[str, type]` constraints schema (sheet-qualified addresses mapped to typing annotations):

```python
from typing import Annotated, Literal

from excel_grapher.grapher import DynamicRefConfig, create_dependency_graph
from excel_grapher.core.cell_types import Between

constraints_schema = {
    "Sheet1!B1": Literal["ROW_INDEX"],
    "Sheet1!C1": Annotated[float, Between(0, 10)],
}

constraints_data: dict[str, object] = {}

config = DynamicRefConfig.from_constraints(constraints_schema, constraints_data)

graph = create_dependency_graph(
    "model_with_dynamic_refs.xlsx",
    ["Sheet1!D10"],
    load_values=False,
    dynamic_refs=config,
    # use_cached_dynamic_refs=False is the default
)
```

Key points:

- Constraint keys use **address-style** strings (e.g. `"Sheet1!B1"`).
- `DynamicRefConfig` is immutable and carries both the `cell_type_env` and `DynamicRefLimits`.
- From the top-level package, you can import `DynamicRefConfig`, `DynamicRefLimits`, and `DynamicRefError`.

### Working with cell data (for transpilation)

The `DependencyGraph` provides direct O(1) access to cell data via `get_node()`, plus filter methods for iterating over formula vs. leaf cells.

```python
from pathlib import Path

from excel_grapher.grapher import create_dependency_graph, discover_formula_cells_in_rows
from excel_grapher.grapher import DependencyGraph

# Discover formula cells in specific rows
targets = discover_formula_cells_in_rows(Path("model.xlsx"), "Sheet1", [10, 11, 12])

# Build the dependency graph
graph: DependencyGraph = create_dependency_graph(Path("model.xlsx"), targets, load_values=True)

# Access cells by normalized address (O(1) lookup)
node = graph.get_node("Sheet1!A10")
print(node.formula)             # Original formula
print(node.normalized_formula)  # Sheet-qualified for transpilation
print(node.value)               # Cached value from Excel
print(node.is_target)           # True if this node came from original targets

# Iterate over formula cells
for key, node in graph.formula_nodes():
    print(key, node.normalized_formula)

# Iterate over leaf (value) cells
for key, node in graph.leaf_node_items():
    print(key, node.value)

# Get sorted keys
formula_keys = graph.formula_keys()
leaf_keys = graph.leaf_keys()
target_keys = graph.target_keys()
```

#### `DependencyGraph` filter methods

| Method              | Returns                                 | Description                           |
|---------------------|------------------------------------------|---------------------------------------|
| `get_node(key)`     | `NodeView \| None`                       | O(1) immutable lookup by cell address |
| `formula_nodes()`   | `Iterator[tuple[NodeKey, Node]]`         | Cells with formulas                   |
| `leaf_node_items()` | `Iterator[tuple[NodeKey, Node]]`         | Leaf cells (no cell dependencies)     |
| `formula_keys()`    | `list[NodeKey]`                          | Sorted keys for formula cells         |
| `leaf_keys()`       | `list[NodeKey]`                          | Sorted keys for leaf cells            |
| `target_keys()`     | `list[NodeKey]`                          | Sorted keys marked as original targets |

#### `Node` fields

| Field                | Type         | Description                               |
|----------------------|--------------|-------------------------------------------|
| `formula`            | `str \| None` | Original formula (``None`` for value-only cells)  |
| `normalized_formula` | `str \| None` | Sheet-qualified formula for transpilation |
| `value`              | `Any`         | Cached or hardcoded value               |
| `is_leaf`            | `bool`        | True if the node has no cell dependencies  |
| `is_target`          | `bool`        | True if the node was one of the original graph targets |
| `sheet`              | `str`         | Sheet name                              |
| `column`             | `str`         | Column letter                           |
| `row`                | `int`         | Row number                              |

#### `discover_formula_cells_in_rows()`

Utility for scanning rows to find formula cells with numeric cached values:

```python
def discover_formula_cells_in_rows(
    wb_path: Path,
    sheet_name: str,
    rows: list[int],
) -> list[str]:
    ...
```

Returns sheet-qualified cell addresses (e.g., `"'Sheet Name'!A1"`) for formula cells.

---

## 2. Visualizing and exporting graphs

### GraphViz DOT

```python
from excel_grapher.grapher import to_graphviz

dot = to_graphviz(g, rankdir="LR")
```

### Mermaid

```python
from excel_grapher.grapher import to_mermaid

mm = to_mermaid(g, max_nodes=100)
```

### Path-induced subgraphs for focused visualization

Use `select_path_induced_subgraph(...)` to isolate only nodes on directed dependency paths between source and target node sets, then pass the smaller graph to any exporter.

```python
from excel_grapher.grapher import select_path_induced_subgraph, to_graphviz

focused = select_path_induced_subgraph(
    g,
    source_keys=["Sheet1!F1"],
    target_keys=["Sheet1!A1"],
    max_path_length=10,  # optional safety cutoff
    max_paths=1000,      # optional safety cutoff
)
dot = to_graphviz(focused, rankdir="LR")
```

The path search follows graph edge direction (`A -> B` means `A` depends on `B`), validates that all requested keys exist in the graph, and preserves edge guards/provenance in the returned induced subgraph.

### NetworkX (optional dependency)

```python
from excel_grapher.grapher import to_networkx

G = to_networkx(g)
```

### Formula text on nodes

For `to_graphviz`, `to_mermaid`, and `to_networkx`, formula cells get a **second line** in the node label (cell address, then the formula). This is on by default (`include_formula_on_nodes=True`). Set `include_formula_on_nodes=False` to use only the cell address. Long formulas are truncated for display: `max_formula_length` defaults to `120` characters; use `None` for no limit. Truncated text ends with `...`.

### Large graphs and module inference: NetworkX visualization

For graphs that are too large for Graphviz or Mermaid, the current recommended workflow is:

- `DependencyGraph` -> `to_networkx(...)` -> `to_web_viz_payload(...)`
- `write_web_viz_html(...)` for rendering

This builds a NetworkX graph, converts it to a visualization payload, generates a static HTML viewer that wraps the payload, and opens the generated HTML in a browser. The payload stores node coordinates, module/rank metadata, and edge columns; the browser viewer consumes those directly.

```python
from pathlib import Path

from excel_grapher.exporter import to_web_viz_payload
from excel_grapher.grapher import create_dependency_graph, to_networkx, write_web_viz_html

g = create_dependency_graph("model.xlsx", ["Sheet1!A1"], load_values=False)
payload = to_web_viz_payload(to_networkx(g))
write_web_viz_html(payload, Path("model.html"), data_mode="auto")
```

![Lightweight visualization overview](README_files/lightweight_viewer.png)

#### Customizing NetworkX visualization

The payload builder (`to_web_viz_payload`) and the viewer (`write_web_viz_html`) are powered by plugins. A single default plugin is provided for each, but you are encouraged to write your own. Configure `to_web_viz_payload(...)` via:

- `layout="..."`: layout plugin id (`stratified_multipartite` default; see `list_web_viz_layouts()`).
- `layout_config={...}`: plugin-specific optional config.
- `include_guarded_edges`: include/exclude guarded edges in core/local edge export.
- `include_guarded_edges_for_partition`: include/exclude guarded edges in module partitioning.
- `include_module_overlay`: include partition overlay metadata and module color semantics.

#### Default layout: `stratified_multipartite`

The default layout is `stratified_multipartite`, which uses SCC-condensation longest-path rank on the vertical axis and Louvain community ordering on the horizontal axis.

```python
from excel_grapher.exporter import to_web_viz_payload, list_web_viz_layouts
from excel_grapher.grapher import write_web_viz_html

payload = to_web_viz_payload(
    G,
    layout="stratified_multipartite",  # default
    layout_config=None,                # plugin-specific options
    include_module_overlay=True,       # include Louvain partition overlay
)
write_web_viz_html(payload, "model-web.html", data_mode="auto")
```

This layout is partly intended to inform refactoring and modularization of generated code. Interpret the viewer with this mental model:

- **Position**: `x`/`y` - SCC-rank vertical strata and Louvain-based horizontal ordering.
- **Color**: node color maps to `module_id` (partition/community id). With no module overlay, all nodes are one module color.
- **Edges**: overview prefers node-level `local_edges`; if unavailable, it falls back to module-centroid edges.
- **Label `Module edges` / `Graph edges`**: reflects which edge set is currently drawn in overview.
- **Rank** in tooltip is node rank metadata and may differ from visual layering for force-directed layouts.

`excel_grapher.exporter.to_web_viz_payload(...)` includes the partition overlay by default and is the canonical payload entrypoint for this visualization workflow. Open the exported HTML directly in a browser. The interface supports panning, zooming, hover tooltips, and a local force-layout mode for inspecting a neighborhood around a selected node.

### Validation via `calcChain.xml`

You can validate the graph against Excel’s `calcChain.xml`:

```python
from pathlib import Path

from excel_grapher.grapher import create_dependency_graph, validate_graph

g = create_dependency_graph("model.xlsx", ["Sheet1!A10"], load_values=False)
res = validate_graph(g, Path("model.xlsx"), scope={"Sheet1"})
print(res.is_valid, res.messages)
```

If `xl/calcChain.xml` is missing (common for generated files), validation returns `is_valid=True` with an informational message.

---

## 3. Evaluating formulas (`excel_grapher.evaluator`)

The evaluator implements Excel’s semantics in Python and runs over a `DependencyGraph`.

### Conceptually

- `FormulaEvaluator` is a wrapper around `DependencyGraph` that:
  - Translates Excel formulas to Python at runtime.
  - Provides Python equivalents for Excel functions, operators, and error types.
  - Handles circular references in a way compatible with Excel’s defaults (warn + return `0`, etc.).
  - Caches results to ensure each cell is computed at most once in a given evaluation.

This gives **fast, accurate, repeatable** execution for any given workbook, but keeps the logic in an Excel-shaped representation. It’s the easiest path when you want to:

- Re-extract and re-run a computation whenever the workbook changes.
- Keep a tight coupling to Excel while still running logic in Python.

### Minimal evaluator example

```python
from excel_grapher.grapher import create_dependency_graph
from excel_grapher.evaluator import FormulaEvaluator

targets = ["Sheet1!B10"]

graph = create_dependency_graph(
    "model.xlsx",
    targets,
    load_values=True,
    max_depth=10,
)

with FormulaEvaluator(graph) as ev:
    evaluator_results = ev.evaluate(targets)

print(evaluator_results)
# {'Sheet1!B10': ...}
```

---

## 4. End-to-end demo: synthetic two-cell workbook

This example builds a **synthetic two-cell workbook** and runs it through the full pipeline.

- `S!A1` is a leaf value (`10`).
- `S!B1` is a formula (`=A1*2`) that references `S!A1`.

### Setup: create the workbook

```python
from __future__ import annotations

import sys
from pathlib import Path

import fastpyxl


def _find_repo_root(start: Path) -> Path:
    for p in [start, *start.parents]:
        if (p / "pyproject.toml").exists():
            return p
    raise RuntimeError("Could not find repo root (missing pyproject.toml)")


def create_synthetic_workbook(path: Path, *, sheet_name: str = "S") -> None:
    path.parent.mkdir(parents=True, exist_ok=True)

    wb = fastpyxl.Workbook()
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

### Build the `DependencyGraph` (dict representation)

```python
import json
from dataclasses import asdict

from excel_grapher.grapher import create_dependency_graph, DependencyGraph
from excel_grapher.evaluator import FormulaEvaluator
from excel_grapher.exporter import CodeGenerator

targets = ["S!B1"]
graph: DependencyGraph = create_dependency_graph(
    workbook_path,
    targets,
    load_values=True,
    max_depth=10,
)

def serialize_graph(graph: DependencyGraph) -> dict:
    return {
        "nodes": {k: asdict(v) for k, v in graph._nodes.items()},
        # Adjacency list: node -> dependencies (edges point from node to its deps)
        "edges": {k: sorted(v) for k, v in graph._edges.items()},
    }

print(json.dumps(serialize_graph(graph), indent=2, sort_keys=True))
```

Example output:

```json
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
```

### Evaluator results

```python
with FormulaEvaluator(graph) as ev:
    evaluator_results = ev.evaluate(targets)

print(evaluator_results)
# {'S!B1': 20.0}
```

### Caching an extracted graph (optional)

If graph extraction is expensive and you expect to re-use the same workbook + targets + extraction settings,
you can cache the `DependencyGraph` to disk as JSON.

Strict caching (requires access to the workbook file to validate fingerprints):

```python
from pathlib import Path

from excel_grapher import (
    CacheValidationPolicy,
    build_graph_cache_meta,
    create_dependency_graph,
    save_graph_cache,
    try_load_graph_cache,
)

workbook_path = Path("workbook.xlsx")
targets = ["S!B1"]
extraction_params = {"load_values": True, "max_depth": 50}

expected = build_graph_cache_meta(workbook_path, targets, extraction_params=extraction_params)
graph = try_load_graph_cache(Path("graph-cache.json"), expected_meta=expected)
if graph is None:
    graph = create_dependency_graph(workbook_path, targets, **extraction_params)
    save_graph_cache(Path("graph-cache.json"), graph, expected)
```

Portable caching (for `FormulaEvaluator` on machines without the workbook file):

```python
from excel_grapher import (
    CacheValidationPolicy,
    build_graph_cache_meta_portable,
    try_load_graph_cache,
)

targets = ["S!B1"]
expected = build_graph_cache_meta_portable(targets, extraction_params={"load_values": True, "max_depth": 50})

graph = try_load_graph_cache(
    Path("graph-cache.json"),
    expected_meta=expected,
    policy=CacheValidationPolicy.PORTABLE,
)
if graph is None:
    raise FileNotFoundError("No valid cached graph found for the requested targets/settings.")
```

**Tradeoffs for the evaluator approach:**

- **Advantages**
  - **Native interface for extraction**: easy to re-extract and re-run if the workbook changes.
  - **Template flexibility**: users can alter workbook structure; re-extraction will follow the new formula graph.
- **Disadvantages**
  - **Runtime translation**: Excel → Python translation happens at runtime for each evaluation.
  - **Coupled to Excel**: still conceptually “driven by Excel” rather than a fully normalized Python model.

---

## 5. Exporting standalone Python (`excel_grapher.exporter`)

The exporter turns a `DependencyGraph` into a standalone Python module:

```python
from excel_grapher.exporter import CodeGenerator

code = CodeGenerator(graph).generate(targets)
print("\n".join(code.splitlines()[:120]))
```

When graph nodes are target-marked (`Node.is_target=True`), `generate()` and `generate_modules()` can infer export targets directly from the graph:

```python
with CodeGenerator(graph) as gen:
    code = gen.generate()  # defaults to graph.target_keys()
```

If neither explicit targets nor target-marked nodes are available, code generation raises a `ValueError`.

You can also emit named entrypoints by passing a mapping of names to target lists:

```python
code = CodeGenerator(graph).generate(
    targets,
    entrypoints={
        "outputs": ["S!B1", "S!C1"],
        "checks": ["S!D1"],
    },
)
```

This generates `compute_outputs(...)` and `compute_checks(...)` alongside `compute_all(...)`.

A (truncated) sketch of the exported code:

```python
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
    ...


def xl_mul(left, right) -> float | XlError:
    ...


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


def compute_all(inputs=None, *, ctx=None):
    """Compute all target cells and return Records."""
    ...
    return _targets_to_records(ctx, TARGETS, TARGET_RECORD_LAYOUT)
```

Generated `compute_all(...)` and `compute_{name}(...)` entrypoints return **`Records`**: a `list[dict]` where each record includes a required `"value"` field and, by default, an `"address"` field with the sheet-qualified cell address. Rectangular targets emit one record per cell in deterministic row-major order.

```python
namespace: dict = {}
exec(code, namespace)
generated_results = namespace["compute_all"]()
print(generated_results)
# [{'address': 'S!B1', 'value': 20.0}]
```

Convert records to an address-keyed dict when needed:

```python
by_address = {rec["address"]: rec["value"] for rec in generated_results if "address" in rec}
```

### Input groups and optional setters

Discover semantic input groups from graph leaf inputs, inspect or edit the payload, then optionally generate setter functions:

```python
from excel_grapher.exporter import (
    CodeGenerator,
    GroupingOptions,
    GroupingOverride,
    SetterGenerationOptions,
)

gen = CodeGenerator(graph)

# 1. Discover groups (label-free by default)
payload = gen.discover_input_groups(
    targets,
    grouping=GroupingOptions(
        include_labels=False,
        overrides=(GroupingOverride(range_spec="S!A1:B3", orientation="columnwise"),),
    ),
)

# 2. Inspect/edit: serialize, edit JSON, or pass explicit groups
edited_groups = payload.groups

# 3. Generate modular package with optional setters.py
files = gen.generate_modules(
    targets,
    setters=SetterGenerationOptions(),
    input_groups=edited_groups,  # skips rediscovery when provided
)
```

When `setters` is omitted, modular export behavior is unchanged (no `setters.py`). When provided, generated `set_*` functions accept `Records` with required `"value"` and apply inputs via `ctx.set_inputs(...)`.

**Migration:** prior exports returned `dict[str, value]` from `compute_*`. New exports return `Records`; use the comprehension above or `tests/integration/utils/parity_harness.records_to_address_dict` during migration.

**Tradeoffs for the exported-code approach:**

- **Advantages**
  - **Standalone artifact**: output is plain Python; no need to distribute `excel_grapher` or the evaluator with it.
  - **Partial obfuscation**: does not expose the extraction engine directly.
  - **Minimal runtime surface**: embeds only the Excel-equivalent `xl_*` helpers actually needed by the exported graph.
  - **Repeatable execution**: freezes workbook logic at a point in time; downstream runs are deterministic and Excel-free.
- **Disadvantages**
  - **Still Excel-shaped**: the structure is still cell-centric and Excel-like; interpretability gains are incremental.
  - **Regeneration required**: changes to the workbook require re-extracting and re-exporting.

---

## 6. Parity testing

**Behavioral parity** is defined across three layers: **Excel** (reference), **`FormulaEvaluator`**, and **standalone exported code**. Shared semantics live in `excel_grapher/exporter/export_runtime/` so the evaluator and generated code stay aligned (**evaluator ↔ export** checks use `tests/integration/utils/parity_harness.py`). **Evaluator ↔ Excel** checks use values saved in the workbook and, when automation is available, **live recalculation** via xlwings that require Excel should **run when automation works** and **`pytest.skip`** otherwise (see `tests/integration/evaluator/test_golden_master.py`).

---

## 7. Roadmap

- Continue expanding **three-way parity** coverage: evaluator ↔ export runtime, evaluator ↔ Excel (cache and live automation where available), especially for representation-sensitive areas such as `OFFSET`, `INDIRECT`, `LOOKUP`, `MATCH`, and `INDEX`.
- Refine the dynamic-reference configuration API and constraints tooling (e.g., validation helpers) as more real-world models and templates are integrated.