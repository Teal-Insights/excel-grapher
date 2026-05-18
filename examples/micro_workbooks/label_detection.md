# Micro-Workbook Label Detection Examples


[examples/micro_workbooks/label_detection.xlsx](label_detection.xlsx)
holds five small layouts for optional graph **label detection**
(`create_dependency_graph(..., label_detection=...)`). Labels are
heuristics over cached cell **values** (not formulas), so
`load_values=True` is required for meaningful results. Detected text is
stored on each node as `metadata["row_labels"]` and
`metadata["column_labels"]` (lists of strings).

``` python
from pathlib import Path
from pprint import pformat

from excel_grapher.grapher import (
    LabelDetectionConfig,
    create_dependency_graph,
    DependencyGraph,
)

workbook_path = Path("label_detection.xlsx")


def print_text(text: str):
    print("```text")
    print(text)
    print("```\n")
```

## 01. Simple row and column labels in neighboring cells

On `Sheet1`, the first layout is a 2×2 block: **B1** names the column,
**A2** names the row, and **B2** holds a constant (`0`). For **B2**, the
default heuristics scan **left** along the row and **up** along the
column, stopping at the first gap. Here the immediate neighbors are
non-empty strings, so they are the only labels collected.

With label detection **disabled** (the default), node metadata stays
empty:

``` python
graph_off: DependencyGraph = create_dependency_graph(
    workbook_path,
    ["Sheet1!B2"],
    load_values=True,
)

node_metadata = graph_off.get_node("Sheet1!B2").metadata
print_text(
    f"Row labels: {node_metadata.get('row_labels', [])}\n"
    f"Column labels: {node_metadata.get('column_labels', [])}"
)
```

``` text
Row labels: []
Column labels: []
```

Pass `LabelDetectionConfig(enabled=True)` to opt in. The same graph now
annotates **B2** with one row label and one column label:

``` python
graph_on: DependencyGraph = create_dependency_graph(
    workbook_path,
    ["Sheet1!B2"],
    load_values=True,
    label_detection=LabelDetectionConfig(enabled=True),
)

node = graph_on.get_node("Sheet1!B2")
node_metadata = node.metadata
print_text(
    f"Row labels: {node_metadata.get('row_labels', [])}\n"
    f"Column labels: {node_metadata.get('column_labels', [])}"
)
```

``` text
Row labels: ['Row 1']
Column labels: ['Column 1']
```

So **B2** is described as belonging to row **“Row 1”** and column
**“Column 1”** in graph metadata, which downstream tools (for example
HTML exports) can use for display names.

## 02. Simple row and column labels in non-neighboring cells

The next block puts numeric entries in the 2×2 block **B5:C6**, with row
labels in **A5:A6** and column labels in **B4:C4**. For target **C6**,
the default left and up scans **skip** the intervening numeric cells
until a string label is found; empty cells end the run.

``` python
graph: DependencyGraph = create_dependency_graph(
    workbook_path,
    ["Sheet1!C6"],
    load_values=True,
    label_detection=LabelDetectionConfig(enabled=True),
)

md = dict(graph.get_node("Sheet1!C6").metadata)
print_text(
    f"Row labels: {md.get('row_labels', [])}\n"
    f"Column labels: {md.get('column_labels', [])}"
)
```

``` text
Row labels: ['Second Row']
Column labels: ['Second Column']
```

## 03. No labels

Since blank cells end the scan, and B8 has blank cells in both
directions, no labels are collected.

``` python
graph: DependencyGraph = create_dependency_graph(
    workbook_path,
    ["Sheet1!B8"],
    load_values=True,
    label_detection=LabelDetectionConfig(enabled=True),
)

md = dict(graph.get_node("Sheet1!B8").metadata)
print_text(
    f"Row labels: {md.get('row_labels', [])}\n"
    f"Column labels: {md.get('column_labels', [])}"
)
```

``` text
Row labels: []
Column labels: []
```

## 04. Nested row and column labels

We allow an arbitrary nested row and column labels, such as when a
single table has two header rows. Scanning left from **D13**, we find
two label columns. Scanning up from the cell, we find two label rows.
The heuristic scan collects the row labels in both A13 and B13, and the
column labels in both D10 and D11.

``` python
graph: DependencyGraph = create_dependency_graph(
    workbook_path,
    ["Sheet1!D13"],
    load_values=True,
    label_detection=LabelDetectionConfig(enabled=True),
)

md = dict(graph.get_node("Sheet1!D13").metadata)
print_text(
    f"Row labels: {md.get('row_labels', [])}\n"
    f"Column labels: {md.get('column_labels', [])}"
)
```

``` text
Row labels: ['Row  2.2', 'Row 1.2']
Column labels: ['Column 2.2', 'Column 2.1']
```

## 05. Merged cells

Scanning left from **D18**, we find a “Row 2.2” label in **B18**.
Immediately left of that, we encounter a merged cell **A17** created
from **A17:A18**. The merged cell **A17** is labeled “Row 1.1”. In
Excel, a merged cell like A17:A18 is visually one big cell, but
internally only the top-left cell (A17) stores the value. The covered
cells (like A18) are placeholders with no value. When `left_edge_scan`
lands on one of those placeholder cells, it resolves the merged range’s
anchor cell (A17) and uses that text, so this label is included in the
detected row labels.

``` python
graph: DependencyGraph = create_dependency_graph(
    workbook_path,
    ["Sheet1!D18"],
    load_values=True,
    label_detection=LabelDetectionConfig(enabled=True),
)

md = dict(graph.get_node("Sheet1!D18").metadata)
print_text(
    f"Row labels: {md.get('row_labels', [])}\n"
    f"Column labels: {md.get('column_labels', [])}"
)
```

``` text
Row labels: ['Row  2.2', 'Row 1.1']
Column labels: ['Column 2.2', 'Column 1.1']
```

## 06. Intervening blank cells

Suppose we have two tables side-by-side, or one above the other. We
don’t want heuristic left-scan or up-scan to accidentally collect labels
from the other table. As long as there are blank cells between the
tables, the scan will stop at the edge of the current table.

The example in rows 20-24 illustrates this. We’ve arranged lists above
and left of our table, separated from the table by blank cells. The text
cells comprising the lists are not detected as labels for the table.

``` python
graph: DependencyGraph = create_dependency_graph(
    workbook_path,
    ["Sheet1!E24"],
    load_values=True,
    label_detection=LabelDetectionConfig(enabled=True),
)

md = dict(graph.get_node("Sheet1!E24").metadata)
print_text(
    f"Row labels: {md.get('row_labels', [])}\n"
    f"Column labels: {md.get('column_labels', [])}"
)
```

``` text
Row labels: ['Row_2']
Column labels: []
```

## 07. Intervening text cells

`left_edge_scan` and `top_edge_scan` prioritize edge labels rather than
all intervening text in a scan path.

If a scan encounters a non-year numeric cell after collecting text
labels, it clears the collected labels and continues scanning. This
filters out intervening text fields while preserving edge labels such as
“Year 1”.

For this example, we have the same table in tall format and wide format,
with a numeric field for population, a text field for country, and a
numeric field for GDP. When we extract the labels for GDP, we want to
skip the country field and only collect the “Year 1” label.

``` python
graph: DependencyGraph = create_dependency_graph(
    workbook_path,
    ["Sheet1!D27", "Sheet1!B32"],
    load_values=True,
    label_detection=LabelDetectionConfig(enabled=True),
)

md_d27 = dict(graph.get_node("Sheet1!D27").metadata)
md_b32 = dict(graph.get_node("Sheet1!B32").metadata)
print_text(
    f"Row labels for D27: {md_d27.get('row_labels', [])}\n"
    f"Column labels for B32: {md_b32.get('column_labels', [])}"
)
```

``` text
Row labels for D27: ['Year 1']
Column labels for B32: ['Year 1']
```

While the default behavior will be to take only left- or top-edge text
cells as labels, there may be times when we want all text cells
connected to the target cell to be considered as labels. For example,
“USA” in this wide-format time series acts like a secondary label
because it is an identifier for the row, so we probably want to collect
it.

The default `LabelDetectionBehavior`s used by `LabelDetectionConfig` are
`left_edge_scan` and `top_edge_scan`, but there are also others
registered on the default registry: `full_row_scan`, `full_column_scan`,
`region_header_rows`, `left_edge_then_up_scan`, and
`top_edge_then_left_scan`. `full_row_scan` and `full_column_scan` are
bidirectional from the reference cell (scan both directions until
blanks), with default output ordering of right-to-left for rows and
bottom-to-top for columns.

You can configure a `BehaviorRule` with a `RegionSelector` to use these
behaviors for a specified spreadsheet region, and pass this rule to
`LabelDetectionConfig`:

``` python
from excel_grapher.grapher import (
    BehaviorRule,
    RegionSelector,
    region_specs_from_ranges,
)

cfg = LabelDetectionConfig(
    enabled=True,
    rules=(
        BehaviorRule(
            name="wideGdpBlock",
            selector=RegionSelector(
                include=region_specs_from_ranges(["Sheet1!A29:B32"]),
            ),
            behaviors=("full_row_scan", "full_column_scan"),
            stop_after_match=True,
        ),
    ),
    fallback_behaviors=(),
)
graph = create_dependency_graph(
    workbook_path,
    ["Sheet1!B32"],
    load_values=True,
    label_detection=cfg,
)
md = dict(graph.get_node("Sheet1!B32").metadata)
print_text(
    f"Row labels: {md.get('row_labels', [])}\n"
    f"Column labels: {md.get('column_labels', [])}"
)
```

``` text
Row labels: ['GDP']
Column labels: ['USA', 'Year 1']
```

## 08. Duplicate labels

Labels are deduplicated by default, so in the next example, “Dupe Label”
is only collected once each for row and column:

``` python
graph: DependencyGraph = create_dependency_graph(
    workbook_path,
    ["Sheet1!C36"],
    load_values=True,
    label_detection=LabelDetectionConfig(enabled=True),
)

md = dict(graph.get_node("Sheet1!C36").metadata)
print_text(
    f"Row labels: {md.get('row_labels', [])}\n"
    f"Column labels: {md.get('column_labels', [])}"
)
```

``` text
Row labels: ['Dupe Label']
Column labels: ['Dupe Label']
```

One reason for deduplicating results is that the user might run multiple
`LabelDetectionBehavior`s (for instance, `left_edge_scan` and
`full_row_scan`) that discover the same label cell. So without
deduplication, we might collect not only multiple cells with the same
label, but also multiple copies of the same label from a single cell.
Note that first-seen order is preserved after deduplication.

## 09. Year labels

Rows 38–40 are a small table: text headers in **A38:B38**, year row ids
**1999** and **2000** in **A39:A40**, and values in **B39:B40**. The
default heuristic scan behaviors only keep numeric cells that look like
calendar years (1900–2100) For **B40**, the default left scan sees
**2000** and adds it to the row labels list.

``` python
graph: DependencyGraph = create_dependency_graph(
    workbook_path,
    ["Sheet1!B40"],
    load_values=True,
    label_detection=LabelDetectionConfig(enabled=True),
)

md = dict(graph.get_node("Sheet1!B40").metadata)
print_text(
    f"Row labels: {md.get('row_labels', [])}\n"
    f"Column labels: {md.get('column_labels', [])}"
)
```

``` text
Row labels: ['2000']
Column labels: ['Revenue']
```

## 10. Numeric labels

Rows 40–42 show the same table with single-digit year offset ids, **1**
and **2**, rather than year-like numbers in **A41:A42**. For **B42**,
the default left scan sees **2** in **A42** and does not keep it because
it is not a calendar year, so the row label list stays empty.

``` python
graph: DependencyGraph = create_dependency_graph(
    workbook_path,
    ["Sheet1!B42"],
    load_values=True,
    label_detection=LabelDetectionConfig(enabled=True),
)

md = dict(graph.get_node("Sheet1!B42").metadata)
print_text(
    f"Row labels: {md.get('row_labels', [])}\n"
    f"Column labels: {md.get('column_labels', [])}"
)
```

``` text
Row labels: ['Year']
Column labels: []
```

If we wanted to collect the year offset labels, we could register a
custom behavior: read column **A** on the target row and coerce the
cached value to a string (any non-empty result counts as a row label).

Here we define a custom behavior class that implements the
`LabelDetectionBehavior` protocol (with a `str` attribute `name` and a
`detect` method that takes a `LabelDetectionContext` and returns a
`LabelResult`). The `LabelDetectionContext` contains the location of the
cell whose labels are being detected, and values and formulas for the
`fastpyxl` worksheet on which the cell resides. `detect` can implement
any detection logic we want that makes use of these values.

``` python
from dataclasses import dataclass

from excel_grapher.grapher import (
    BehaviorRule,
    LabelDetectionBehavior,
    LabelDetectionContext,
    LabelResult,
    RegionLabelParams,
    RegionSelector,
    region_specs_from_ranges,
)


# custom class that implements `LabelDetectionBehavior` protocol
@dataclass
class ColumnARowLabel(LabelDetectionBehavior):
    name: str = "column_a_row_label"

    def detect(self, ctx: LabelDetectionContext) -> LabelResult:
        if ctx.ws_values is None:
            return LabelResult()
        value = ctx.ws_values.cell(row=ctx.row, column=1).value
        if value is None:
            return LabelResult()
        text = str(value).strip()
        if not text:
            return LabelResult()
        return LabelResult(row_labels=(text,))


cfg = LabelDetectionConfig(
    enabled=True,
    rules=(
        BehaviorRule(
            name="numericRowIds",
            selector=RegionSelector(
                include=region_specs_from_ranges(["Sheet1!A42:B44"]),
            ),
            # reference the custom behavior by name
            behaviors=("column_a_row_label", "top_edge_scan"),
            stop_after_match=True,
        ),
    ),
    fallback_behaviors=(),
)

graph = create_dependency_graph(
    workbook_path,
    ["Sheet1!B44"],
    load_values=True,
    label_detection=cfg,
    # register the custom behavior
    label_behaviors=[ColumnARowLabel()],
)

md = dict(graph.get_node("Sheet1!B44").metadata)
print_text(
    f"Row labels: {md.get('row_labels', [])}\n"
    f"Column labels: {md.get('column_labels', [])}"
)
```

``` text
Row labels: ['2']
Column labels: ['Revenue']
```

Note that when we define a BehaviorRule that uses our custom behavior,
we apply it only to the worksheet region (“Sheet1!A42:B44”) that we want
to collect year offset labels from, and we still apply “top_edge_scan”
to collect the column labels. We can apply any combination of behaviors
we want to different parts of the workbook.

## 11. Transforming years to offsets

`year_offset_headers` is not in the default registry because this
behavior is domain-specific. Register it explicitly (or implement a
custom equivalent) when needed.

The next snippet shows one way to implement that as a custom behavior
for a **tall-format** block (`Sheet1!A38:B40`), where years are in
column **A** and values are in column **B**. For **B40**, this converts
year **2000** to `offset:1` relative to the first year in the block
(**1999**), while still collecting the top column label (`"Revenue"`).

``` python
from dataclasses import dataclass

from excel_grapher.grapher import (
    BehaviorRule,
    LabelDetectionBehavior,
    LabelDetectionContext,
    LabelResult,
    RegionSelector,
    region_specs_from_ranges,
)


@dataclass
class YearOffsetRowLabel(LabelDetectionBehavior):
    name: str = "year_offset_row_label"

    def detect(self, ctx: LabelDetectionContext) -> LabelResult:
        rp = ctx.region_params
        if rp is None or ctx.ws_values is None or rp.min_row is None or rp.max_row is None:
            return LabelResult()

        # We use the first numeric year in column A of the configured region as the baseline.
        base_year: int | None = None
        for row in range(rp.min_row, rp.max_row + 1):
            v = ctx.ws_values.cell(row=row, column=1).value
            if isinstance(v, int) and 1900 <= v <= 2100:
                base_year = v
                break
        if base_year is None:
            return LabelResult()

        current = ctx.ws_values.cell(row=ctx.row, column=1).value
        if not isinstance(current, int) or not (1900 <= current <= 2100):
            return LabelResult()

        return LabelResult(row_labels=(f"offset:{current - base_year}",))


cfg = LabelDetectionConfig(
    enabled=True,
    rules=(
        BehaviorRule(
            name="tallYearOffsets",
            selector=RegionSelector(
                include=region_specs_from_ranges(["Sheet1!A38:B40"]),
            ),
            behaviors=("year_offset_row_label", "top_edge_scan"),
            stop_after_match=True,
            # Region bounds are consumed by the custom behavior.
            region_params=RegionLabelParams(min_row=39, max_row=40),
        ),
    ),
    fallback_behaviors=(),
)

graph = create_dependency_graph(
    workbook_path,
    ["Sheet1!B40"],
    load_values=True,
    label_detection=cfg,
    label_behaviors=[YearOffsetRowLabel()],
)

md = dict(graph.get_node("Sheet1!B40").metadata)
print_text(
    f"Row labels: {md.get('row_labels', [])}\n"
    f"Column labels: {md.get('column_labels', [])}"
)
```

``` text
Row labels: ['offset:1']
Column labels: ['Revenue']
```

This keeps the year-offset behavior explicit and local to the region and
model that need it, without relying on a globally registered default.

## 12. Rightward or downward scans

Rightward and downward scans are available in the default registry.

In rows 46-49 is a table with units to the right of a numeric column,
and with a source field at the bottom of the column. To collect these
labels for the cell **A48** (while still collecting the column label),
we can use `right_edge_scan` and `bottom_edge_scan` in combination with
`top_edge_scan`:

``` python
from excel_grapher.grapher import (
    BehaviorRule,
    RegionSelector,
    region_specs_from_ranges,
)

cfg = LabelDetectionConfig(
    enabled=True,
    rules=(
        BehaviorRule(
            name="unitsAndSourceBlock",
            selector=RegionSelector(
                include=region_specs_from_ranges(["Sheet1!A46:C49"]),
            ),
            behaviors=("right_edge_scan", "bottom_edge_scan", "top_edge_scan"),
            stop_after_match=True,
        ),
    ),
    fallback_behaviors=(),
)

graph = create_dependency_graph(
    workbook_path,
    ["Sheet1!A48"],
    load_values=True,
    label_detection=cfg,
)

md = dict(graph.get_node("Sheet1!A48").metadata)
print_text(
    f"Row labels: {md.get('row_labels', [])}\n"
    f"Column labels: {md.get('column_labels', [])}"
)
```

``` text
Row labels: ['million', 'dollars', 'Source: CIA Factbook, 2012']
Column labels: ['Value']
```

## 13. Left, then up scans

Rows 51-56 show a small wide-format time series with grouped rows and
hierarchical row labels. A simple heuristic left-scan from **B56** would
only collect “Real” from **A56**, but this label alone isn’t very
informative. The double-indentation of **A56** indicates that it is a
child of the single-indented **A54** label, “GDP”, and the unindented
**A53** label, “United States”. To understand what **B56** represents,
we need to scan left to collect the row label, then up to collect parent
labels. This is implemented by the built-in `left_edge_then_up_scan`
behavior.

``` python
graph: DependencyGraph = create_dependency_graph(
    workbook_path,
    ["Sheet1!B56"],
    load_values=True,
    label_detection=LabelDetectionConfig(
        enabled=True,
        rules=(
            BehaviorRule(
                name="indentHierarchy",
                selector=RegionSelector(
                    include=region_specs_from_ranges(["Sheet1!A51:B56"]),
                ),
                behaviors=("left_edge_then_up_scan", "top_edge_scan"),
            ),
        ),
    )
)

md = dict(graph.get_node("Sheet1!B56").metadata)
print_text(
    f"Row labels: {md.get('row_labels', [])}\n"
    f"Column labels: {md.get('column_labels', [])}"
)
```

``` text
Row labels: ['United States', 'GDP', 'Real']
Column labels: []
```

This behavior uses indentation as the primary hierarchy signal and text
style rank (`bold > italic > normal`) as a secondary signal.

What if we wanted to implement a custom behavior that uses font weight
to infer hierarchy while ignoring indentation? The `ws_values` and
`ws_formulas` attributes of `LabelDetectionContext` expose a `cell`
selector method that can be used to get a
[`fastpyxl.cell.cell.Cell`](https://fastpyxl.readthedocs.io/en/latest/api/fastpyxl.cell.cell.html#fastpyxl.cell.cell.Cell)
object, and this `Cell` object (a subclass of
[`StylableObject`](https://fastpyxl.readthedocs.io/en/latest/api/fastpyxl.styles.styleable.html#fastpyxl.styles.styleable.StyleableObject))
exposes style fields such as `alignment` and `font` that can be used in
label detection logic. Here’s an example of how to implement such logic:

``` python
from dataclasses import dataclass
from fastpyxl.utils.cell import column_index_from_string

from excel_grapher.grapher import (
    BehaviorRule,
    LabelDetectionBehavior,
    LabelDetectionContext,
    LabelDetectionConfig,
    LabelResult,
    RegionLabelParams,
    RegionSelector,
    create_dependency_graph,
    region_specs_from_ranges,
)


@dataclass
class FontWeightHierarchyRowLabels(LabelDetectionBehavior):
    name: str = "font_weight_hierarchy_row_labels"

    def detect(self, ctx: LabelDetectionContext) -> LabelResult:
        rp = ctx.region_params
        if (
            ctx.ws_values is None
            or rp is None
            or not rp.label_columns
            or rp.min_row is None
            or rp.max_row is None
        ):
            return LabelResult()

        # Demo assumption: one label column with hierarchy based on bold vs non-bold text.
        col_letter = rp.label_columns[0]
        col_idx = column_index_from_string(col_letter)

        current_parent: str | None = None
        row_labels: list[str] = []
        for row in range(rp.min_row, min(ctx.row, rp.max_row) + 1):
            cell = ctx.ws_values.cell(row=row, column=col_idx)
            if not isinstance(cell.value, str):
                continue
            text = cell.value.strip()
            if not text:
                continue

            is_bold = bool(cell.font and cell.font.bold)
            if is_bold:
                current_parent = text

            if row == ctx.row:
                if is_bold:
                    row_labels = [text]
                elif current_parent is not None:
                    row_labels = [current_parent, text]
                else:
                    row_labels = [text]

        return LabelResult(row_labels=tuple(row_labels))


cfg = LabelDetectionConfig(
    enabled=True,
    rules=(
        BehaviorRule(
            name="fontWeightHierarchyDemo",
            selector=RegionSelector(
                include=region_specs_from_ranges(["Sheet1!A51:B56"]),
            ),
            behaviors=("font_weight_hierarchy_row_labels", "top_edge_scan"),
            stop_after_match=True,
            region_params=RegionLabelParams(
                label_columns=("A",),
                min_row=51,
                max_row=56,
            ),
        ),
    ),
    fallback_behaviors=(),
)

graph = create_dependency_graph(
    workbook_path,
    ["Sheet1!B56"],
    load_values=True,
    label_detection=cfg,
    label_behaviors=[FontWeightHierarchyRowLabels()],
)

md = dict(graph.get_node("Sheet1!B56").metadata)
print_text(
    f"Row labels: {md.get('row_labels', [])}\n"
    f"Column labels: {md.get('column_labels', [])}"
)
```

``` text
Row labels: ['United States', 'Real']
Column labels: []
```

## 14. Merge policy: replace for canonicalization

Sometimes table labels are technically correct but too verbose for
display. In this example, row labels include units/qualifiers (for
example, `"GDP ($, current prices)"`), and we want a canonical display
label (for example, `"GDP"`, stripping the parenthetical trailing text).

We run two behaviors in order on the same region:

1.  `left_edge_scan` collects the raw row label from the table.
2.  `canonicalize_col_a_label` rewrites that label to a canonical form.

With `merge_policy="replace"`, the canonical behavior overwrites the
earlier raw label when it returns a non-empty result.

``` python
import re
from dataclasses import dataclass

from excel_grapher.grapher import (
    BehaviorRule,
    LabelDetectionBehavior,
    LabelDetectionConfig,
    LabelDetectionContext,
    LabelResult,
    RegionSelector,
    region_specs_from_ranges,
)

_TRAILING_PARENS_RE = re.compile(r"\s*\([^)]*\)\s*$")


@dataclass
class CanonicalizeColALabel(LabelDetectionBehavior):
    name: str = "canonicalize_col_a_label"

    def detect(self, ctx: LabelDetectionContext) -> LabelResult:
        if ctx.ws_values is None:
            return LabelResult()
        value = ctx.ws_values.cell(row=ctx.row, column=1).value
        if not isinstance(value, str):
            return LabelResult()
        text = value.strip()
        if not text:
            return LabelResult()
        canonical = _TRAILING_PARENS_RE.sub("", text).strip()
        if not canonical:
            return LabelResult()
        return LabelResult(row_labels=(canonical,))

cfg = LabelDetectionConfig(
    enabled=True,
    merge_policy="replace",
    rules=(
        BehaviorRule(
            name="rawLabels",
            selector=RegionSelector(
                include=region_specs_from_ranges(["Sheet1!A58:B60"]),
            ),
            behaviors=("left_edge_scan",),
        ),
        BehaviorRule(
            name="canonicalLabels",
            selector=RegionSelector(
                include=region_specs_from_ranges(["Sheet1!A58:B60"]),
            ),
            behaviors=("canonicalize_col_a_label",),
            stop_after_match=True,
        ),
    ),
    fallback_behaviors=(),
)

graph = create_dependency_graph(
    workbook_path,
    ["Sheet1!B59", "Sheet1!B60"],
    load_values=True,
    label_detection=cfg,
    label_behaviors=[CanonicalizeColALabel()],
)

md_b59 = dict(graph.get_node("Sheet1!B59").metadata)
md_b60 = dict(graph.get_node("Sheet1!B60").metadata)
print_text(
    f"B59 row labels: {md_b59.get('row_labels', [])}\n"
    f"B60 row labels: {md_b60.get('row_labels', [])}"
)
```

``` text
B59 row labels: ['GDP']
B60 row labels: ['Debt service']
```

Another use case for `merge_policy="replace"` is when you want to apply
one behavior to a large spreadsheet region, but then replace it with a
custom behavior override for a subregion.

## 15. Merge policy: append_dedupe_reverse

`append_dedupe` keeps first-seen labels first. If you want parent or
higher-priority labels to appear first after all behaviors run, you can
use `append_dedupe_reverse`.

``` python
cfg_append = LabelDetectionConfig(
    enabled=True,
    merge_policy="append_dedupe",
    fallback_behaviors=("full_row_scan", "left_edge_scan"),
)
cfg_reverse = LabelDetectionConfig(
    enabled=True,
    merge_policy="append_dedupe_reverse",
    fallback_behaviors=("full_row_scan", "left_edge_scan"),
)

graph_append = create_dependency_graph(
    workbook_path,
    ["Sheet1!D63"],
    load_values=True,
    label_detection=cfg_append,
)
graph_reverse = create_dependency_graph(
    workbook_path,
    ["Sheet1!D63"],
    load_values=True,
    label_detection=cfg_reverse,
)

md_append = dict(graph_append.get_node("Sheet1!D63").metadata)
md_reverse = dict(graph_reverse.get_node("Sheet1!D63").metadata)
print_text(
    f"append_dedupe row labels: {md_append.get('row_labels', [])}\n"
    f"append_dedupe_reverse row labels: {md_reverse.get('row_labels', [])}"
)
```

``` text
append_dedupe row labels: ['C-label', 'A-label']
append_dedupe_reverse row labels: ['A-label', 'C-label']
```
