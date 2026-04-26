# Consolidated Micro-Workbook Graph Examples


# Consolidated Micro-Workbook Graph Examples

This walkthrough uses runnable Python cells to extract each example
graph from a single row-formatted workbook.

``` python
from pathlib import Path
from pprint import pformat

import fastpyxl
from IPython.display import Markdown, display

from excel_grapher.grapher import create_dependency_graph, to_mermaid

workbook_path = Path("example-cases.xlsx")
workbook = fastpyxl.load_workbook(workbook_path, data_only=False)
sheet = workbook["Sheet1"]

cases = [
    ("Formula with no dependencies", "Sheet1!B1", True),
    ("Linear dependency", "Sheet1!C2", False),
    ("Conditional branches", "Sheet1!E3", True),
    ("Nested conditional in a cell", "Sheet1!D4", False),
    ("Nested conditional across cells", "Sheet1!E5", False),
    ("Will cycle", "Sheet1!C6", False),
    ("Won't cycle", "Sheet1!C7", False),
    ("May cycle", "Sheet1!C8", False),
]
```

``` python
for idx, (label, target, show_graph_object) in enumerate(cases, start=1):
    graph = create_dependency_graph(workbook, [target], load_values=False)
    mermaid = to_mermaid(graph)

    print(f"## {idx:02d}. {label}\n")
    print(f"Target: `{target}`\n")

    if show_graph_object:
        print("Graph type:\n")
        print("```text")
        print(type(graph))
        print("```\n")

        print("Graph object:\n")
        print("```text")
        print(pformat(graph, indent=4, width=100))
        print("```\n")

    print("```{mermaid}")
    print(mermaid)
    print("```\n")
```

## 01. Formula with no dependencies

Target: `Sheet1!B1`

Graph type:

``` text
<class 'excel_grapher.grapher.graph.DependencyGraph'>
```

Graph object:

``` text
DependencyGraph(_nodes={   'Sheet1!B1': Node(sheet='Sheet1',
                                             column='B',
                                             row=1,
                                             formula='=1+1',
                                             normalized_formula='=1+1',
                                             value=None,
                                             is_leaf=False,
                                             metadata={})},
                _edges={'Sheet1!B1': set()},
                _reverse_edges={'Sheet1!B1': set()},
                _guards={},
                _edge_extra={},
                _hooks=[],
                leaf_classification=None)
```

``` mermaid
flowchart TD
  Sheet1_B1("Sheet1!B1<br>=1+1")
```

## 02. Linear dependency

Target: `Sheet1!C2`

``` mermaid
flowchart TD
  Sheet1_B2["Sheet1!B2"]
  Sheet1_C2("Sheet1!C2<br>=B2+1")
  Sheet1_C2 --> Sheet1_B2
```

## 03. Conditional branches

Target: `Sheet1!E3`

Graph type:

``` text
<class 'excel_grapher.grapher.graph.DependencyGraph'>
```

Graph object:

``` text
DependencyGraph(_nodes={   'Sheet1!B3': Node(sheet='Sheet1',
                                             column='B',
                                             row=3,
                                             formula=None,
                                             normalized_formula=None,
                                             value=1,
                                             is_leaf=True,
                                             metadata={}),
                           'Sheet1!C3': Node(sheet='Sheet1',
                                             column='C',
                                             row=3,
                                             formula=None,
                                             normalized_formula=None,
                                             value=10,
                                             is_leaf=True,
                                             metadata={}),
                           'Sheet1!D3': Node(sheet='Sheet1',
                                             column='D',
                                             row=3,
                                             formula=None,
                                             normalized_formula=None,
                                             value=20,
                                             is_leaf=True,
                                             metadata={}),
                           'Sheet1!E3': Node(sheet='Sheet1',
                                             column='E',
                                             row=3,
                                             formula='=IF(B3=1,C3,D3)',
                                             normalized_formula='=IF(Sheet1!B3=1,Sheet1!C3,Sheet1!D3)',
                                             value=None,
                                             is_leaf=False,
                                             metadata={})},
                _edges={   'Sheet1!B3': set(),
                           'Sheet1!C3': set(),
                           'Sheet1!D3': set(),
                           'Sheet1!E3': {'Sheet1!C3', 'Sheet1!B3', 'Sheet1!D3'}},
                _reverse_edges={   'Sheet1!B3': {'Sheet1!E3'},
                                   'Sheet1!C3': {'Sheet1!E3'},
                                   'Sheet1!D3': {'Sheet1!E3'},
                                   'Sheet1!E3': set()},
                _guards={   ('Sheet1!E3', 'Sheet1!C3'): Compare(left=CellRef(key='Sheet1!B3'),
                                                                op='=',
                                                                right=Literal(value=1)),
                            ('Sheet1!E3', 'Sheet1!D3'): Not(operand=Compare(left=CellRef(key='Sheet1!B3'),
                                                                            op='=',
                                                                            right=Literal(value=1)))},
                _edge_extra={},
                _hooks=[],
                leaf_classification=None)
```

``` mermaid
flowchart TD
  Sheet1_B3["Sheet1!B3"]
  Sheet1_C3["Sheet1!C3"]
  Sheet1_D3["Sheet1!D3"]
  Sheet1_E3("Sheet1!E3<br>=IF(B3=1,C3,D3)")
  Sheet1_E3 --> Sheet1_B3
  Sheet1_E3 -.->|"Sheet1!B3=1"| Sheet1_C3
  Sheet1_E3 -.->|"NOT(Sheet1!B3=1)"| Sheet1_D3
```

## 04. Nested conditional in a cell

Target: `Sheet1!D4`

``` mermaid
flowchart TD
  Sheet1_B4["Sheet1!B4"]
  Sheet1_C4["Sheet1!C4"]
  Sheet1_D4("Sheet1!D4<br>=IF(NOT(B4=1),IF(B4=0,C4,1),0)")
  Sheet1_D4 --> Sheet1_B4
  Sheet1_D4 -.->|"NOT(Sheet1!B4=1)"| Sheet1_C4
```

## 05. Nested conditional across cells

Target: `Sheet1!E5`

``` mermaid
flowchart TD
  Sheet1_A5["Sheet1!A5"]
  Sheet1_B5["Sheet1!B5"]
  Sheet1_C5["Sheet1!C5"]
  Sheet1_E5("Sheet1!E5<br>=IF(A5=1,B5,C5)")
  Sheet1_E5 --> Sheet1_A5
  Sheet1_E5 -.->|"Sheet1!A5=1"| Sheet1_B5
  Sheet1_E5 -.->|"NOT(Sheet1!A5=1)"| Sheet1_C5
```

## 06. Will cycle

Target: `Sheet1!C6`

``` mermaid
flowchart TD
  Sheet1_B6("Sheet1!B6<br>=C6+1")
  Sheet1_C6("Sheet1!C6<br>=B6+1")
  Sheet1_B6 --> Sheet1_C6
  Sheet1_C6 --> Sheet1_B6
```

## 07. Won’t cycle

Target: `Sheet1!C7`

``` mermaid
flowchart TD
  Sheet1_B7["Sheet1!B7"]
  Sheet1_C7("Sheet1!C7<br>=IF(B7=0,1,D7)")
  Sheet1_D7("Sheet1!D7<br>=IF(NOT(B7=0),2,C7)")
  Sheet1_C7 --> Sheet1_B7
  Sheet1_C7 -.->|"NOT(Sheet1!B7=0)"| Sheet1_D7
  Sheet1_D7 --> Sheet1_B7
  Sheet1_D7 -.->|"NOT(NOT(Sheet1!B7=0))"| Sheet1_C7
```

## 08. May cycle

Target: `Sheet1!C8`

``` mermaid
flowchart TD
  Sheet1_B8["Sheet1!B8"]
  Sheet1_C8("Sheet1!C8<br>=IF(B8=0,1,D8)")
  Sheet1_D8("Sheet1!D8<br>=IF(B8=1,2,C8)")
  Sheet1_C8 --> Sheet1_B8
  Sheet1_C8 -.->|"NOT(Sheet1!B8=0)"| Sheet1_D8
  Sheet1_D8 --> Sheet1_B8
  Sheet1_D8 -.->|"NOT(Sheet1!B8=1)"| Sheet1_C8
```
