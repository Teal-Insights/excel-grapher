#!/usr/bin/env python3
from __future__ import annotations

import argparse
import json
import os
import re
import subprocess
import time
import urllib.error
import urllib.parse
import urllib.request
from dataclasses import dataclass
from pathlib import Path
from typing import Any

GITHUB_API_VERSION = "2026-03-10"
DEFAULT_REPOS = [
    "excel-grapher",
    "lic-dsf-extraction-pipeline",
    "tiny-dsa-extraction-pipeline",
    "qcraft-extraction-pipeline",
    "qcraft-v2-planning",
]
ISSUE_LABEL_WIDTH = 200
GRAPH_FONT_NAME = "Arial"
GRAPH_FONT_SIZE = 10
GRAPH_NODE_SEP = 0.5
GRAPH_RANK_SEP = 0.8


@dataclass(frozen=True)
class WorkflowIssue:
    repo: str
    number: int
    title: str
    pre_group_key: str | None
    pre_group_label: str | None
    post_group_key: str | None
    post_group_label: str | None

    @property
    def node_id(self) -> str:
        repo_key = self.repo.replace("-", "_")
        return f"{repo_key}_{self.number}"


def _quote(value: str) -> str:
    escaped = value.replace("\\", "\\\\").replace('"', '\\"')
    return f'"{escaped}"'


def _escape_html(value: str) -> str:
    return (
        value.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;").replace('"', "&quot;")
    )


def _normalize_repo_key(repo: str) -> str:
    return repo.replace("-", "_")


def _request_with_retry(request: urllib.request.Request) -> Any:
    backoff = 1.0
    max_attempts = 5
    for attempt in range(max_attempts):
        try:
            with urllib.request.urlopen(request, timeout=30) as response:
                body = response.read()
                return json.loads(body.decode()) if body else None
        except urllib.error.HTTPError as error:
            if error.code in {429, 502, 503, 504} and attempt < max_attempts - 1:
                retry_after = error.headers.get("Retry-After")
                wait_for = float(retry_after) if retry_after else backoff
                time.sleep(wait_for)
                backoff *= 2
                continue
            detail = error.read().decode(errors="replace")
            raise RuntimeError(f"GitHub API request failed ({error.code}): {detail}") from error

    raise RuntimeError("GitHub API retry loop exhausted")


def _github_request(
    owner: str,
    repo: str,
    path: str,
    token: str,
    *,
    query: dict[str, str] | None = None,
) -> Any:
    url = f"https://api.github.com/repos/{owner}/{repo}{path}"
    if query:
        url = f"{url}?{urllib.parse.urlencode(query)}"

    request = urllib.request.Request(
        url=url,
        headers={
            "Accept": "application/vnd.github+json",
            "Authorization": f"Bearer {token}",
            "X-GitHub-Api-Version": GITHUB_API_VERSION,
        },
    )
    return _request_with_retry(request)


def _resolve_token(token: str | None) -> str:
    if token:
        return token
    env_token = os.environ.get("GITHUB_TOKEN")
    if env_token:
        return env_token
    result = subprocess.run(
        ["gh", "auth", "token"],
        capture_output=True,
        text=True,
        check=False,
    )
    if result.returncode != 0 or not result.stdout.strip():
        raise RuntimeError("No GitHub token provided. Set GITHUB_TOKEN or pass --github-token.")
    return result.stdout.strip()


def fetch_open_issues(owner: str, repo: str, token: str) -> list[dict[str, Any]]:
    page = 1
    issues: list[dict[str, Any]] = []
    while True:
        rows = _github_request(
            owner,
            repo,
            "/issues",
            token,
            query={
                "state": "open",
                "per_page": "100",
                "page": str(page),
            },
        )
        assert isinstance(rows, list)
        if not rows:
            break
        for row in rows:
            if "pull_request" in row:
                continue
            issues.append(row)
        page += 1
    return issues


def fetch_label_descriptions(owner: str, repo: str, token: str) -> dict[str, str]:
    page = 1
    descriptions: dict[str, str] = {}
    while True:
        rows = _github_request(
            owner,
            repo,
            "/labels",
            token,
            query={"per_page": "100", "page": str(page)},
        )
        assert isinstance(rows, list)
        if not rows:
            break
        for row in rows:
            descriptions[row["name"]] = row.get("description") or row["name"]
        page += 1
    return descriptions


def _choose_group_label(label_names: list[str]) -> str | None:
    slash_candidates = []
    single_level_candidates = []
    for name in label_names:
        if name == "deliverables" or name.startswith("deliverables/"):
            continue
        if "/" in name:
            parts = name.split("/", 1)
            if len(parts) != 2 or not parts[0] or not parts[1]:
                continue
            slash_candidates.append(name)
            continue
        if name:
            single_level_candidates.append(name)

    if slash_candidates:
        return sorted(slash_candidates)[0]
    if single_level_candidates:
        return sorted(single_level_candidates)[0]
    return None


def _has_deliverables_group_label(label_names: list[str]) -> bool:
    return any(name == "deliverables" or name.startswith("deliverables/") for name in label_names)


def build_workflow_issues(
    issues_by_repo: dict[str, list[dict[str, Any]]],
    label_descriptions_by_repo: dict[str, dict[str, str]],
) -> list[WorkflowIssue]:
    workflow_issues: list[WorkflowIssue] = []
    for repo, issues in issues_by_repo.items():
        repo_descriptions = label_descriptions_by_repo.get(repo, {})
        for issue in issues:
            label_names = [label["name"] for label in issue.get("labels", [])]
            group_label = _choose_group_label(label_names)
            if group_label is None:
                if _has_deliverables_group_label(label_names):
                    continue
                pre_key = None
                post_key = None
                pre_label = None
                post_label = None
            else:
                if "/" in group_label:
                    pre_key, post_key = group_label.split("/", 1)
                    pre_label = repo_descriptions.get(pre_key, pre_key)
                    post_label = repo_descriptions.get(group_label, post_key)
                else:
                    pre_key = group_label
                    post_key = None
                    pre_label = repo_descriptions.get(group_label, group_label)
                    post_label = None
            workflow_issues.append(
                WorkflowIssue(
                    repo=repo,
                    number=int(issue["number"]),
                    title=issue["title"],
                    pre_group_key=pre_key,
                    pre_group_label=pre_label,
                    post_group_key=post_key,
                    post_group_label=post_label,
                )
            )
    workflow_issues.sort(
        key=lambda i: (
            i.repo,
            i.pre_group_key or "",
            i.post_group_key or "",
            i.number,
        )
    )
    return workflow_issues


def fetch_blocks_edges(
    owner: str,
    token: str,
    workflow_issues: list[WorkflowIssue],
) -> list[tuple[str, int, str, int]]:
    issue_lookup = {(issue.repo, issue.number): issue for issue in workflow_issues}
    edges: set[tuple[str, int, str, int]] = set()
    for issue in workflow_issues:
        blocked_by = _github_request(
            owner,
            issue.repo,
            f"/issues/{issue.number}/dependencies/blocked_by",
            token,
        )
        assert isinstance(blocked_by, list)
        for blocker in blocked_by:
            blocker_repo = blocker["repository"]["name"]
            blocker_number = int(blocker["number"])
            if (blocker_repo, blocker_number) not in issue_lookup:
                continue
            edges.add((blocker_repo, blocker_number, issue.repo, issue.number))
    return sorted(edges)


def build_dot(
    issues: list[WorkflowIssue],
    blocks_edges: list[tuple[str, int, str, int]],
    *,
    owner: str | None = None,
    rankdir: str = "TB",
) -> str:
    by_repo_direct: dict[str, list[WorkflowIssue]] = {}
    by_repo_pre_only: dict[str, dict[str, list[WorkflowIssue]]] = {}
    by_repo_grouped: dict[str, dict[str, dict[str, list[WorkflowIssue]]]] = {}
    for issue in issues:
        if issue.pre_group_key is None:
            by_repo_direct.setdefault(issue.repo, []).append(issue)
            continue
        if issue.post_group_key is None:
            pre_only_map = by_repo_pre_only.setdefault(issue.repo, {})
            pre_only_map.setdefault(issue.pre_group_key, []).append(issue)
            continue
        pre_map = by_repo_grouped.setdefault(issue.repo, {})
        post_map = pre_map.setdefault(issue.pre_group_key, {})
        post_map.setdefault(issue.post_group_key, []).append(issue)

    pre_labels = {
        (issue.repo, issue.pre_group_key): issue.pre_group_label
        for issue in issues
        if issue.pre_group_key is not None and issue.pre_group_label is not None
    }
    post_labels = {
        (issue.repo, issue.pre_group_key, issue.post_group_key): issue.post_group_label
        for issue in issues
        if issue.pre_group_key is not None
        and issue.post_group_key is not None
        and issue.post_group_label is not None
    }
    all_repos = sorted(set(by_repo_direct) | set(by_repo_pre_only) | set(by_repo_grouped))

    lines: list[str] = []
    lines.append("digraph workflow {")
    lines.append(
        f"  graph [compound=true, rankdir={rankdir}, nodesep={GRAPH_NODE_SEP}, ranksep={GRAPH_RANK_SEP}];"
    )
    lines.append(
        f'  node [shape=box, style="rounded", fontname={_quote(GRAPH_FONT_NAME)}, fontsize={GRAPH_FONT_SIZE}];'
    )
    lines.append(f"  edge [fontname={_quote(GRAPH_FONT_NAME)}, fontsize=9];")
    lines.append("")

    def append_issue_node(indent: str, issue: WorkflowIssue) -> None:
        number_html = _escape_html(f"#{issue.number}")
        title_html = _escape_html(issue.title)
        label_html = (
            '<<TABLE BORDER="0" CELLBORDER="0" CELLPADDING="4" CELLSPACING="0">'
            f'<TR><TD WIDTH="{ISSUE_LABEL_WIDTH}" ALIGN="LEFT"><B>{number_html}</B> &#183; {title_html}</TD></TR>'
            "</TABLE>>"
        )
        node_label = f"#{issue.number} · {issue.title}"
        attrs = [f"label={label_html}"]
        if owner is not None:
            issue_url = f"https://github.com/{owner}/{issue.repo}/issues/{issue.number}"
            attrs.extend(
                [
                    f"URL={_quote(issue_url)}",
                    f"href={_quote(issue_url)}",
                    f"tooltip={_quote(node_label)}",
                    'target="_blank"',
                ]
            )
        lines.append(f"{indent}{issue.node_id} [{', '.join(attrs)}];")

    for repo in all_repos:
        repo_key = _normalize_repo_key(repo)
        lines.append(f'  subgraph "cluster_repo_{repo_key}" {{')
        lines.append(f"    label={_quote(f'{repo} GitHub issues')};")
        lines.append("")

        for issue in sorted(by_repo_direct.get(repo, []), key=lambda item: item.number):
            append_issue_node("    ", issue)
        if by_repo_direct.get(repo):
            lines.append("")

        for pre_key in sorted(by_repo_pre_only.get(repo, {})):
            pre_key_safe = re.sub(r"[^A-Za-z0-9_]", "_", pre_key)
            pre_label = pre_labels[(repo, pre_key)] or pre_key
            lines.append(f'    subgraph "cluster_repo_{repo_key}_{pre_key_safe}" {{')
            lines.append(f"      label={_quote(pre_label)};")
            for issue in sorted(by_repo_pre_only[repo][pre_key], key=lambda item: item.number):
                append_issue_node("      ", issue)
            lines.append("    }")
            lines.append("")

        for pre_key in sorted(by_repo_grouped.get(repo, {})):
            pre_key_safe = re.sub(r"[^A-Za-z0-9_]", "_", pre_key)
            pre_label = pre_labels[(repo, pre_key)] or pre_key
            lines.append(f'    subgraph "cluster_repo_{repo_key}_{pre_key_safe}" {{')
            lines.append(f"      label={_quote(pre_label)};")
            lines.append("")

            for post_key in sorted(by_repo_grouped[repo][pre_key]):
                post_key_safe = re.sub(r"[^A-Za-z0-9_]", "_", post_key)
                post_label = post_labels[(repo, pre_key, post_key)] or post_key
                lines.append(
                    f'      subgraph "cluster_repo_{repo_key}_{pre_key_safe}_{post_key_safe}" {{'
                )
                lines.append(f"        label={_quote(post_label)};")
                for issue in sorted(
                    by_repo_grouped[repo][pre_key][post_key], key=lambda item: item.number
                ):
                    append_issue_node("        ", issue)
                lines.append("      }")
                lines.append("")

            lines.append("    }")
            lines.append("")

        lines.append("  }")
        lines.append("")

    for blocker_repo, blocker_number, blocked_repo, blocked_number in blocks_edges:
        blocker_id = f"{_normalize_repo_key(blocker_repo)}_{blocker_number}"
        blocked_id = f"{_normalize_repo_key(blocked_repo)}_{blocked_number}"
        lines.append(f'  {blocker_id} -> {blocked_id} [label="blocks"];')

    lines.append("}")
    lines.append("")
    return "\n".join(lines)


def render_dot_to_svg(*, dot_path: Path, svg_path: Path) -> None:
    dot_binary = os.environ.get("DOT_BIN", "dot")
    result = subprocess.run(
        [dot_binary, "-Tsvg", str(dot_path), "-o", str(svg_path)],
        capture_output=True,
        text=True,
        check=False,
    )
    if result.returncode != 0:
        detail = result.stderr.strip() or result.stdout.strip() or "unknown graphviz error"
        raise RuntimeError(f"Graphviz render failed: {detail}")


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Build DOT workflow graph from open GitHub issues and labels."
    )
    parser.add_argument("--owner", default="Teal-Insights")
    parser.add_argument(
        "--repos",
        nargs="+",
        default=DEFAULT_REPOS,
        help="Repository names under --owner.",
    )
    parser.add_argument(
        "--output",
        type=Path,
        default=Path("artifacts/workflow.dot"),
    )
    parser.add_argument(
        "--layout",
        choices=["TB", "LR"],
        default="TB",
        help="Graph layout direction: TB (top-down, default) or LR (left-right).",
    )
    parser.add_argument("--github-token")
    parser.add_argument(
        "--render-svg",
        action="store_true",
        help="Render SVG from generated DOT using Graphviz `dot`.",
    )
    parser.add_argument(
        "--svg-output",
        type=Path,
        help="Path for rendered SVG output (default: <output with .svg suffix>).",
    )
    args = parser.parse_args()

    token = _resolve_token(args.github_token)

    issues_by_repo: dict[str, list[dict[str, Any]]] = {}
    label_descriptions_by_repo: dict[str, dict[str, str]] = {}
    for repo in args.repos:
        issues_by_repo[repo] = fetch_open_issues(args.owner, repo, token)
        label_descriptions_by_repo[repo] = fetch_label_descriptions(args.owner, repo, token)

    workflow_issues = build_workflow_issues(issues_by_repo, label_descriptions_by_repo)
    blocks_edges = fetch_blocks_edges(args.owner, token, workflow_issues)
    dot = build_dot(workflow_issues, blocks_edges, owner=args.owner, rankdir=args.layout)

    args.output.parent.mkdir(parents=True, exist_ok=True)
    args.output.write_text(dot, encoding="utf-8")
    print(f"Wrote DOT graph to {args.output}")
    print(f"Issues included: {len(workflow_issues)}")
    print(f"Blocks edges included: {len(blocks_edges)}")

    if args.render_svg:
        svg_output = (
            args.svg_output if args.svg_output is not None else args.output.with_suffix(".svg")
        )
        svg_output.parent.mkdir(parents=True, exist_ok=True)
        render_dot_to_svg(dot_path=args.output, svg_path=svg_output)
        print(f"Rendered SVG graph to {svg_output}")


if __name__ == "__main__":
    main()
