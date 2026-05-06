from __future__ import annotations

from scripts.build_workflow_cytoscape_preset import (
    build_cytoscape_preset_payload,
    build_index_html,
)
from scripts.build_workflow_dot_from_github import WorkflowIssue


def test_build_cytoscape_preset_payload_maps_clusters_positions_and_edges() -> None:
    issues = [
        WorkflowIssue(
            repo="excel-grapher",
            number=1,
            title="One",
            pre_group_key="major",
            pre_group_label="Major",
            post_group_key=None,
            post_group_label=None,
        ),
        WorkflowIssue(
            repo="excel-grapher",
            number=2,
            title="Two",
            pre_group_key=None,
            pre_group_label=None,
            post_group_key=None,
            post_group_label=None,
        ),
    ]

    graphviz_json = {
        "bb": "0,0,100,200",
        "objects": [
            {"_gvid": 0, "name": "cluster_repo_excel_grapher", "label": "excel", "nodes": [2, 3]},
            {
                "_gvid": 1,
                "name": "cluster_repo_excel_grapher_major",
                "label": "Major",
                "nodes": [2],
            },
            {
                "_gvid": 2,
                "name": "excel_grapher_1",
                "label": "#1 · One",
                "pos": "10,190",
                "width": "0.75",
                "height": "0.5",
            },
            {
                "_gvid": 3,
                "name": "excel_grapher_2",
                "label": "#2 · Two",
                "pos": "30,50",
                "width": "0.75",
                "height": "0.5",
            },
        ],
        "edges": [
            {"tail": 2, "head": 3},
        ],
    }

    payload = build_cytoscape_preset_payload(
        owner="Teal-Insights",
        workflow_issues=issues,
        graphviz_json=graphviz_json,
    )

    nodes = payload["elements"]["nodes"]
    edges = payload["elements"]["edges"]
    node_by_id = {node["data"]["id"]: node for node in nodes}

    assert "cluster::cluster_repo_excel_grapher" in node_by_id
    assert "cluster::cluster_repo_excel_grapher_major" in node_by_id
    assert (
        node_by_id["excel_grapher_1"]["data"]["parent"]
        == "cluster::cluster_repo_excel_grapher_major"
    )
    assert node_by_id["excel_grapher_2"]["data"]["parent"] == "cluster::cluster_repo_excel_grapher"
    assert node_by_id["excel_grapher_1"]["data"]["gv_width"] == 54.0  # 0.75 in × 72 pt/in
    assert node_by_id["excel_grapher_1"]["data"]["gv_height"] == 36.0
    # y-axis flips from Graphviz coords (max_y - y)
    assert node_by_id["excel_grapher_1"]["position"] == {"x": 10.0, "y": 10.0}
    assert edges[0]["data"]["source"] == "excel_grapher_1"
    assert edges[0]["data"]["target"] == "excel_grapher_2"


def test_build_index_html_references_preset_json() -> None:
    html = build_index_html(json_filename="workflow.json")
    assert "workflow.json" in html
    assert "Graphviz preset" in html
    assert "layout: {" in html and "name: 'preset'" in html
