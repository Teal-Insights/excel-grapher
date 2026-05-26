# Series bindings

Declarative **series bindings** describe how spreadsheet input cells map to named dimensions (SDMX-style concepts) and how generated Python **setters** accept `Records` — `list[dict]` with a required `value` field plus key columns.

They replace the retired label-detection heuristics (`row_labels` / `column_labels` on graph nodes). The machine source of truth is a **sidecar manifest** next to the workbook, not inferred geometry.

**Schema:** [`series_binding.schema.json`](series_binding.schema.json)  
**Example workbook:** [`../examples/micro_workbooks/series_bindings.xlsx`](../examples/micro_workbooks/series_bindings.xlsx) with [`series_bindings.bindings.yaml`](../examples/micro_workbooks/series_bindings.bindings.yaml)  
**Walkthrough:** [`../examples/micro_workbooks/series_bindings.qmd`](../examples/micro_workbooks/series_bindings.qmd)

---

## Authoring conventions

| Topic | Convention |
|--------|------------|
| **Format** | YAML by default (`.bindings.yaml`); JSON (`.bindings.json`) also accepted |
| **Location** | One sidecar per workbook, colocated with the `.xlsx` |
| **Granularity (MVP)** | One `series[]` entry per **indicator row** → one `set_*` function (`layout: row_series`) |
| **Naming** | Stable `id` (snake_case); `setter.name` must match `set_[a-z][a-z0-9_]*` |
| **Non-executable prose** | `sdmx_notes`, `notes` — documentation only |

### Default file layout

```text
lic_inputs.xlsx
lic_inputs.bindings.yaml    # schema_version, workbook, series[]
```

All setters for that workbook live in a single `series[]` array.

### Optional per-sheet shards

For large workbooks, split by sheet and merge at load time:

```text
lic_inputs.xlsx
lic_inputs.bindings/
  Inputs.bindings.yaml
  Assumptions.bindings.yaml
```

Rules:

- Same `schema_version` and `workbook` in every shard (error otherwise).
- Union `series[]`; **reject duplicate `series[].id`** across shards.
- `load_series_bindings(path)` accepts a **file or directory** of `*.bindings.{yaml,yml,json}`.

**Non-goals:** bindings embedded inside the xlsx; one file per series; repo-wide manifest without a `workbook` field.

### `data_range` addressing

Expansion uses the same path as graph **targets** (`expand_targets_to_roots`), including:

- Local range: `Inputs!F5:J5`
- **Both-end sheet-qualified** (project standard): `Inputs!F5:Inputs!J5`, `'My Sheet'!A1:'My Sheet'!B2`
- **Defined names**: `PrimaryRow` (requires `workbook=` or a graph built from that workbook so named-range maps are available)

```python
expand_data_range("Inputs!F5:Inputs!J5")
expand_data_range("PrimaryRow", workbook=path_to_xlsx)
expand_data_range_for_graph(graph, "PrimaryRow", workbook=path_to_xlsx)
```

---

## LIC manifest table → sidecar workflow

Analysts often maintain a **manifest table** (Excel or markdown) listing inputs: country, indicator, units, sheet/range, and notes. Treat that table as human QA; the **sidecar YAML** is what tooling validates and codegen consumes.

Suggested workflow:

1. **Inventory** inputs in the manifest (one row per logical series / indicator row).
2. **Draft** `series_bindings.bindings.yaml` — one `series[]` item per manifest row that should become a setter.
3. **Map columns** from the manifest into binding fields:

| Manifest column (typical) | Binding field |
|---------------------------|---------------|
| Sheet name | `sheet` (required today; may become inferred from `data_range`) |
| Cell range (e.g. `F5:J5`) | `data_range` |
| Country / REF_AREA | `series_context.REF_AREA` or `bind.kind: cell` |
| Indicator label | `bind.kind: row_label` + `series_context.INDICATOR` |
| Year / period axis | `bind.kind: column_header` + `key: [TIME_PERIOD]` |
| Unit | `structure.attributes[]` with `value:` |
| Setter name | `setter.name` |

4. **Validate** against the dependency graph (see [Python API](#python-api)).
5. **Export** with `CodeGenerator(..., series_bindings=..., bindings_workbook=...)` to emit `set_*` functions.

Offset projection years (1..5) vs calendar years (1999..2000) are distinguished only by what you bind in `column_header` — declare the intended semantics in `sdmx_notes`.

---

## Records contract

Generated setters accept:

```python
records: list[dict[str, object]]
```

Each record must include:

- **`value`** — observation (`OBS_VALUE`) written to the matched leaf cell.
- **Key fields** — every concept listed in `key` (e.g. `TIME_PERIOD: 3`).

Optional:

- **`address`** or **`cell_address`** — when `setter.allow_address` is true or keys are ambiguous (`requires_address`).
- Other fields allowed when `setter.strict` is false; strict mode rejects unknown keys.

`series_context` values are attached to emitted records for documentation; they are not used for matching unless also listed in `key`.

---

## Bind kinds

| `bind.kind` | Scope | Status | Purpose |
|-------------|--------|--------|---------|
| `data_cell` | leaf | **supported** | Read numeric input from the leaf itself |
| `cell` | series | **supported** | Fixed address (e.g. country in `A2`) |
| `column_header` | cell | **supported** | Header row for the data column |
| `row_label` | series/cell | **supported** | Label text on the data row |
| `constant` | series | **supported** | Fixed scalar (e.g. scalar layout key) |

Normalization: `strip`, `strip_trailing_unit` (removes trailing `(…)` unit clauses from indicator text).

---

## Layouts

| `layout` | Status | Description |
|----------|--------|-------------|
| `row_series` | **supported** | One indicator row × many columns (MVP) |
| `scalar` | **supported** | Single editable cell |
| `matrix` | **schema only** | Multi-row blocks — not implemented yet |
| `column_series` | **planned** | Not in schema yet; track via follow-up issues |

---

## Schema versions

| Version | Contents |
|---------|----------|
| **1.0.0** | `row_series`, `scalar`; supported bind kinds above. Use for production workbooks. |
| **1.1.0** | Draft extensions (e.g. `matrix` in schema). Not fully implemented in tooling. |

---

## Python API

```python
from pathlib import Path

from excel_grapher.grapher import create_dependency_graph
from excel_grapher.series_bindings import (
    bindings_canonical_sha256,
    derive_input_series,
    expand_data_range,
    load_series_bindings,
    validate_series_bindings,
)
from excel_grapher.exporter import CodeGenerator

workbook = Path("lic_inputs.xlsx")
bindings = load_series_bindings(workbook.with_suffix(".bindings.yaml"))

targets = [
    addr
    for s in bindings["series"]
    for addr in expand_data_range(s["data_range"], workbook=workbook)
]
graph = create_dependency_graph(workbook, targets, load_values=True)

report = validate_series_bindings(graph, bindings, workbook=workbook)
assert report["ok"]

bindings_canonical_sha256(bindings)
input_series = derive_input_series(graph, bindings, workbook=workbook)

with CodeGenerator(graph) as gen:
    code = gen.generate(
        targets,
        series_bindings=bindings,
        bindings_workbook=workbook,
    )
```

After export:

```python
ctx = make_context()
set_borvelia_primary_balance(ctx, [{"TIME_PERIOD": 3, "value": -0.5}])
results = compute_all(ctx=ctx)
```

For modular package exports, generated series setters are emitted from the package entrypoint and re-exported from the package root, alongside `make_context` and `compute_all`.

---

## Relation to issue #185

Issue #185 is now framed around **input series** derived from series bindings, not independently discovered input groups. One binding id corresponds to one generated setter and one input-series view over the participating graph leaves in `data_range`.

Design history: [GitHub issue #192](https://github.com/Teal-Insights/excel-grapher/issues/192).
