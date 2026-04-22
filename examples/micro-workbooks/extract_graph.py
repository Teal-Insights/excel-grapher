from pathlib import Path

from excel_grapher.grapher import create_dependency_graph, to_graphviz


def extract_graph() -> None:
    workbook_path: Path = Path("examples/micro-workbooks/01-single-cell-no-deps.xlsx")
    graph = create_dependency_graph(workbook_path, ["Sheet1!A1"], load_values=False)
    print(graph)


if __name__ == "__main__":
    extract_graph()
