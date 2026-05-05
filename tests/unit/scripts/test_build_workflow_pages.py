from __future__ import annotations

from scripts.build_workflow_dot_from_github import WorkflowIssue
from scripts.build_workflow_pages import build_index_html, build_workflow_json


def test_build_workflow_json_includes_repo_group_and_issue_nodes() -> None:
    issues = [
        WorkflowIssue(
            repo="excel-grapher",
            number=51,
            title="Support type-annotated code export from CodeGenerator",
            pre_group_key="major",
            pre_group_label="Major/ambitious new features",
            post_group_key=None,
            post_group_label=None,
        ),
        WorkflowIssue(
            repo="excel-grapher",
            number=135,
            title="Redesign constraints API",
            pre_group_key="api_audit",
            pre_group_label="High-level API design audits and improvements",
            post_group_key="leaf_config_api_audit",
            post_group_label="Audit configuration API and storage data model",
        ),
        WorkflowIssue(
            repo="lic-dsf-programmatic-extraction",
            number=19,
            title="Update extraction script",
            pre_group_key=None,
            pre_group_label=None,
            post_group_key=None,
            post_group_label=None,
        ),
    ]
    edges = [("excel-grapher", 135, "lic-dsf-programmatic-extraction", 19)]

    payload = build_workflow_json(owner="Teal-Insights", issues=issues, blocks_edges=edges)

    nodes = payload["elements"]["nodes"]
    edges_json = payload["elements"]["edges"]

    node_ids = {node["data"]["id"] for node in nodes}
    assert "repo::excel-grapher" in node_ids
    assert "repo::lic-dsf-programmatic-extraction" in node_ids
    assert "group::excel-grapher::major" in node_ids
    assert "group::excel-grapher::api_audit" in node_ids
    assert "group::excel-grapher::api_audit/leaf_config_api_audit" in node_ids
    assert "issue::excel-grapher#51" in node_ids
    assert "issue::excel-grapher#135" in node_ids
    assert "issue::lic-dsf-programmatic-extraction#19" in node_ids

    major_issue = next(node for node in nodes if node["data"]["id"] == "issue::excel-grapher#51")
    assert major_issue["data"]["parent"] == "group::excel-grapher::major"

    cross_repo_issue = next(
        node for node in nodes if node["data"]["id"] == "issue::lic-dsf-programmatic-extraction#19"
    )
    assert cross_repo_issue["data"]["parent"] == "repo::lic-dsf-programmatic-extraction"
    assert edges_json[0]["data"]["source"] == "issue::excel-grapher#135"
    assert edges_json[0]["data"]["target"] == "issue::lic-dsf-programmatic-extraction#19"


def test_build_index_html_points_to_json_file() -> None:
    html = build_index_html(json_filename="workflow.json")

    assert "cytoscape" in html
    assert 'fetch("workflow.json")' in html
    assert "Click an issue node to open it in GitHub." in html
    assert "'width': 'label'" in html
    assert "'height': 'label'" in html
    assert "'text-valign': 'top'" in html
    assert "'text-margin-y': -8" in html
    assert "cytoscape-fcose" in html
    assert "cytoscape-klay" in html
    assert 'option value="klay" selected' in html
    assert "layoutSelect" in html
