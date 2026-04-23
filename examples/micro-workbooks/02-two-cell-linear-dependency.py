from pathlib import Path
from pprint import pprint

import matplotlib.pyplot as plt
import networkx as nx

from excel_grapher.grapher import create_dependency_graph, to_networkx

if __name__ == "__main__":
    # Path to the workbook to extract the graph from
    WORKBOOK_PATH: Path = Path("examples/micro-workbooks/02-two-cell-linear-dependency.xlsx")
    TARGET: str = "Sheet1!A2"

    # Extract the graph from the workbook with A2 as the target
    print(f"Extracting graph from {WORKBOOK_PATH} with target {TARGET}...")
    graph = create_dependency_graph(WORKBOOK_PATH, [TARGET], load_values=False)

    # Print the graph object's type and field names
    print(f"\nGraph type: {type(graph)}\nGraph fields: {graph.__dataclass_fields__.keys()}\n")

    # Print the graph
    pprint(graph, indent=4)

    # Export to networkx and write to file
    G = to_networkx(graph)
    plt.figure(figsize=(10, 10))
    pos = nx.spring_layout(G, seed=42)
    nx.draw(G, pos, with_labels=True, arrows=True, node_color="lightblue", edge_color="gray")
    plt.title("Dependency Graph")
    plt.axis("off")
    plt.tight_layout()
    plt.savefig("examples/micro-workbooks/02-two-cell-linear-dependency.png", dpi=150)
