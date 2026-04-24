from pathlib import Path
from pprint import pprint

from excel_grapher.grapher import create_dependency_graph, to_graphviz

if __name__ == "__main__":
    # Path to the workbook to extract the graph from
    WORKBOOK_PATH: Path = Path("examples/micro-workbooks/08-nested-conditional-across-cells.xlsx")
    TARGET: str = "Sheet1!A4"

    # Extract the graph from the workbook with A1 as the target
    print(f"Extracting graph from {WORKBOOK_PATH} with target {TARGET}...")
    graph = create_dependency_graph(WORKBOOK_PATH, [TARGET], load_values=False)

    # Print the graph object's type and field names
    print(f"\nGraph type: {type(graph)}\nGraph fields: {graph.__dataclass_fields__.keys()}\n")

    # Print the graph
    pprint(graph, indent=4)

    # Export to graphviz and write to file
    dot = to_graphviz(graph)
    with open("examples/micro-workbooks/08-nested-conditional-across-cells.qmd", "w") as f:
        f.write(f"```{{dot}}\n{dot}\n```")
