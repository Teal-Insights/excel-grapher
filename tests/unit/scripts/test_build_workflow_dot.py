from __future__ import annotations

from pathlib import Path

from scripts.build_workflow import (
    WorkflowIssue,
    build_dot,
    build_workflow_issues,
    render_dot_to_svg,
)


def test_build_workflow_issues_uses_slash_labels_and_descriptions() -> None:
    issues_by_repo = {
        "excel-grapher": [
            {
                "number": 135,
                "title": "Redesign constraints API",
                "labels": [
                    {"name": "api_audit/leaf_config_api_audit"},
                    {"name": "priority/high"},
                ],
            },
            {
                "number": 118,
                "title": "No grouping label yet",
                "labels": [],
            },
            {
                "number": 119,
                "title": "Single-level grouped issue",
                "labels": [{"name": "major"}],
            },
            {
                "number": 117,
                "title": "By default, enforce behavioral parity",
                "labels": [{"name": "deliverables/whitepaper"}],
            },
        ]
    }
    label_descriptions_by_repo = {
        "excel-grapher": {
            "api_audit": "High-level API design audits and improvements",
            "api_audit/leaf_config_api_audit": "Audit configuration API and storage data model",
            "priority": "Priority",
            "priority/high": "High priority",
            "major": "Major/ambitious new features",
            "deliverables": "Prototypes and deliverables",
            "deliverables/whitepaper": "Whitepaper deliverable",
        }
    }

    workflow_issues = build_workflow_issues(issues_by_repo, label_descriptions_by_repo)

    assert sorted(i.number for i in workflow_issues) == [118, 119, 135]
    issue_by_number = {issue.number: issue for issue in workflow_issues}

    unlabeled_issue = issue_by_number[118]
    assert unlabeled_issue.pre_group_key is None
    assert unlabeled_issue.pre_group_label is None
    assert unlabeled_issue.post_group_key is None
    assert unlabeled_issue.post_group_label is None

    single_level_grouped = issue_by_number[119]
    assert single_level_grouped.pre_group_key == "major"
    assert single_level_grouped.pre_group_label == "Major/ambitious new features"
    assert single_level_grouped.post_group_key is None
    assert single_level_grouped.post_group_label is None

    grouped_issue = issue_by_number[135]
    assert grouped_issue.pre_group_key == "api_audit"
    assert grouped_issue.pre_group_label == "High-level API design audits and improvements"
    assert grouped_issue.post_group_key == "leaf_config_api_audit"
    assert grouped_issue.post_group_label == "Audit configuration API and storage data model"


def test_build_dot_renders_repo_pre_post_hierarchy_and_blocks_edges() -> None:
    issues = [
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
            repo="lic-dsf-extraction-pipeline",
            number=19,
            title="Update extraction script",
            pre_group_key="api_audit",
            pre_group_label="High-level API design audits and improvements",
            post_group_key="leaf_config_api_audit",
            post_group_label="Audit configuration API and storage data model",
        ),
    ]
    blocks_edges = [
        ("excel-grapher", 135, "lic-dsf-extraction-pipeline", 19),
    ]

    dot = build_dot(issues, blocks_edges)

    assert "rankdir=TB" in dot
    assert "nodesep=0.5" in dot
    assert "ranksep=0.8" in dot
    assert 'subgraph "cluster_repo_excel_grapher"' in dot
    assert 'subgraph "cluster_repo_lic_dsf_extraction_pipeline"' in dot
    assert "Redesign constraints API" in dot
    assert "Update extraction script" in dot
    assert 'WIDTH="200"' in dot
    assert "<TABLE" in dot
    assert '[label="blocks"]' in dot
    assert "cluster_repo_excel_grapher_unlabeled" not in dot


def test_build_dot_adds_clickable_issue_urls_when_owner_provided() -> None:
    issues = [
        WorkflowIssue(
            repo="excel-grapher",
            number=135,
            title="Redesign constraints API",
            pre_group_key=None,
            pre_group_label=None,
            post_group_key=None,
            post_group_label=None,
        ),
    ]
    dot = build_dot(issues, [], owner="Teal-Insights")

    assert 'URL="https://github.com/Teal-Insights/excel-grapher/issues/135"' in dot
    assert 'target="_blank"' in dot


def test_build_dot_supports_left_right_layout() -> None:
    dot = build_dot([], [], rankdir="LR")
    assert "rankdir=LR" in dot


def test_render_dot_to_svg_invokes_graphviz(monkeypatch, tmp_path: Path) -> None:
    dot_path = tmp_path / "workflow.dot"
    svg_path = tmp_path / "workflow.svg"
    dot_path.write_text("digraph workflow {}", encoding="utf-8")

    calls: list[list[str]] = []

    class CompletedProcess:
        returncode = 0

    def fake_run(cmd, capture_output, text, check):  # noqa: ANN001
        calls.append(cmd)
        return CompletedProcess()

    monkeypatch.setattr("subprocess.run", fake_run)

    render_dot_to_svg(dot_path=dot_path, svg_path=svg_path)

    assert calls == [["dot", "-Tsvg", str(dot_path), "-o", str(svg_path)]]
