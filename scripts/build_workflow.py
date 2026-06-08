#!/usr/bin/env python3
"""Build the workflow graph site for great-docs pre-render.

Runs from the ``great-docs/`` build directory. Fetches live GitHub issue data when
a token is available; otherwise copies a checked-in snapshot from ``custom/workflow/``
or ``docs/``. Registers ``workflow/**`` as a Quarto resource so ``index.html`` and
``workflow.json`` deploy to ``_site/workflow/``.
"""

from __future__ import annotations

import json
import os
import re
import shutil
import subprocess
import time
import urllib.error
import urllib.parse
import urllib.request
from dataclasses import dataclass
from pathlib import Path
from typing import Any

import yaml

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
WORKFLOW_DIR = "workflow"
WORKFLOW_RESOURCE = f"{WORKFLOW_DIR}/**"
WORKFLOW_HTML_EXCLUDE = f"!{WORKFLOW_DIR}/index.html"


def _find_project_root(start: Path) -> Path:
    for candidate in (start, *start.parents):
        if (candidate / "pyproject.toml").is_file():
            return candidate
    return start


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
    for env_name in ("GITHUB_TOKEN", "ISSUE_READ_PAT"):
        env_token = os.environ.get(env_name)
        if env_token:
            return env_token
    result = subprocess.run(
        ["gh", "auth", "token"],
        capture_output=True,
        text=True,
        check=False,
    )
    if result.returncode != 0 or not result.stdout.strip():
        raise RuntimeError(
            "No GitHub token provided. Set GITHUB_TOKEN or ISSUE_READ_PAT, or pass --github-token."
        )
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


def _parse_graphviz_json(dot_text: str, *, dot_bin: str = "dot") -> dict[str, Any]:
    result = subprocess.run(
        [dot_bin, "-Tjson"],
        input=dot_text,
        capture_output=True,
        text=True,
        check=False,
    )
    if result.returncode != 0:
        detail = result.stderr.strip() or result.stdout.strip() or "unknown graphviz error"
        raise RuntimeError(f"Graphviz JSON render failed: {detail}")
    return json.loads(result.stdout)


def _parse_xy(pos: str) -> tuple[float, float]:
    x_text, y_text = pos.split(",", 1)
    return float(x_text), float(y_text)


def _graphviz_size_inches_to_points(value: Any) -> float | None:
    """Graphviz JSON emits node width/height as inch amounts (strings or numbers)."""
    if value is None:
        return None
    try:
        inches = float(value)
    except (TypeError, ValueError):
        return None
    return inches * 72.0


def _graphviz_max_y(graph_json: dict[str, Any]) -> float:
    bb = graph_json.get("bb")
    if not isinstance(bb, str):
        return 0.0
    # bb is "x0,y0,x1,y1"
    parts = bb.split(",")
    if len(parts) != 4:
        return 0.0
    return float(parts[3])


def _cluster_depths(cluster_parents: dict[int, int | None]) -> dict[int, int]:
    cache: dict[int, int] = {}

    def depth(cluster_id: int) -> int:
        if cluster_id in cache:
            return cache[cluster_id]
        parent = cluster_parents.get(cluster_id)
        if parent is None:
            cache[cluster_id] = 0
        else:
            cache[cluster_id] = depth(parent) + 1
        return cache[cluster_id]

    for cluster_id in cluster_parents:
        depth(cluster_id)
    return cache


def build_cytoscape_preset_payload(
    *,
    owner: str,
    workflow_issues: list[WorkflowIssue],
    graphviz_json: dict[str, Any],
) -> dict[str, Any]:
    objects = graphviz_json.get("objects", [])
    if not isinstance(objects, list):
        raise RuntimeError("Graphviz JSON missing objects list")

    objects_by_gvid: dict[int, dict[str, Any]] = {}
    node_gvid_by_name: dict[str, int] = {}
    cluster_gvids: set[int] = set()
    for obj in objects:
        gvid = obj.get("_gvid")
        name = obj.get("name")
        if not isinstance(gvid, int) or not isinstance(name, str):
            continue
        objects_by_gvid[gvid] = obj
        node_gvid_by_name[name] = gvid
        if name.startswith("cluster_"):
            cluster_gvids.add(gvid)

    cluster_parents: dict[int, int | None] = {cluster_id: None for cluster_id in cluster_gvids}
    for cluster_id in cluster_gvids:
        cluster_obj = objects_by_gvid[cluster_id]
        for child in cluster_obj.get("subgraphs", []) or []:
            if isinstance(child, int) and child in cluster_gvids:
                cluster_parents[child] = cluster_id

    # Fallback parent inference when Graphviz JSON lacks explicit `subgraphs`.
    cluster_nodes: dict[int, set[int]] = {
        cluster_id: {
            n for n in (objects_by_gvid[cluster_id].get("nodes", []) or []) if isinstance(n, int)
        }
        for cluster_id in cluster_gvids
    }
    for cluster_id in cluster_gvids:
        if cluster_parents[cluster_id] is not None:
            continue
        my_nodes = cluster_nodes[cluster_id]
        if not my_nodes:
            continue
        candidates: list[tuple[int, int]] = []
        for other_id in cluster_gvids:
            if other_id == cluster_id:
                continue
            other_nodes = cluster_nodes[other_id]
            if my_nodes < other_nodes:
                candidates.append((len(other_nodes), other_id))
        if candidates:
            candidates.sort()
            cluster_parents[cluster_id] = candidates[0][1]
    cluster_depth = _cluster_depths(cluster_parents)

    node_to_clusters: dict[int, set[int]] = {}
    for cluster_id in cluster_gvids:
        cluster_obj = objects_by_gvid[cluster_id]
        for node_gvid in cluster_obj.get("nodes", []) or []:
            if not isinstance(node_gvid, int):
                continue
            node_to_clusters.setdefault(node_gvid, set()).add(cluster_id)

    issue_by_node_name = {issue.node_id: issue for issue in workflow_issues}

    elements_nodes: list[dict[str, Any]] = []
    elements_edges: list[dict[str, Any]] = []
    max_y = _graphviz_max_y(graphviz_json)

    cluster_node_id_by_gvid = {
        gvid: f"cluster::{objects_by_gvid[gvid]['name']}" for gvid in cluster_gvids
    }
    for cluster_id in sorted(cluster_gvids):
        cluster_obj = objects_by_gvid[cluster_id]
        label = cluster_obj.get("label") or cluster_obj.get("name")
        data: dict[str, Any] = {
            "id": cluster_node_id_by_gvid[cluster_id],
            "label": label,
            "type": "cluster",
            "cluster_name": cluster_obj.get("name"),
        }
        parent_id = cluster_parents.get(cluster_id)
        if parent_id is not None:
            data["parent"] = cluster_node_id_by_gvid[parent_id]
        elements_nodes.append({"data": data})

    for issue in workflow_issues:
        node_gvid = node_gvid_by_name.get(issue.node_id)
        if node_gvid is None:
            continue
        node_obj = objects_by_gvid[node_gvid]
        pos = node_obj.get("pos")
        if not isinstance(pos, str):
            continue
        x, y = _parse_xy(pos)
        y = max_y - y

        data: dict[str, Any] = {
            "id": issue.node_id,
            "label": f"#{issue.number} · {issue.title}",
            "type": "issue",
            "repo": issue.repo,
            "number": issue.number,
            "title": issue.title,
            "url": f"https://github.com/{owner}/{issue.repo}/issues/{issue.number}",
        }
        w_pt = _graphviz_size_inches_to_points(node_obj.get("width"))
        h_pt = _graphviz_size_inches_to_points(node_obj.get("height"))
        if w_pt is not None and h_pt is not None:
            data["gv_width"] = w_pt
            data["gv_height"] = h_pt
        containing_clusters = node_to_clusters.get(node_gvid, set())
        if containing_clusters:
            leaf_cluster = max(containing_clusters, key=lambda cid: cluster_depth.get(cid, 0))
            data["parent"] = cluster_node_id_by_gvid[leaf_cluster]
        elements_nodes.append({"data": data, "position": {"x": x, "y": y}})

    issue_node_names = set(issue_by_node_name)
    for edge in graphviz_json.get("edges", []) or []:
        tail = edge.get("tail")
        head = edge.get("head")
        if not isinstance(tail, int) or not isinstance(head, int):
            continue
        tail_obj = objects_by_gvid.get(tail)
        head_obj = objects_by_gvid.get(head)
        if tail_obj is None or head_obj is None:
            continue
        source = tail_obj.get("name")
        target = head_obj.get("name")
        if not isinstance(source, str) or not isinstance(target, str):
            continue
        if source not in issue_node_names or target not in issue_node_names:
            continue
        elements_edges.append(
            {
                "data": {
                    "id": f"edge::{source}->{target}",
                    "source": source,
                    "target": target,
                    "type": "blocks",
                    "label": "blocks",
                }
            }
        )

    return {
        "meta": {
            "owner": owner,
            "repos": sorted({issue.repo for issue in workflow_issues}),
            "issue_count": len(
                [node for node in elements_nodes if node["data"]["type"] == "issue"]
            ),
            "cluster_count": len(
                [node for node in elements_nodes if node["data"]["type"] == "cluster"]
            ),
            "edge_count": len(elements_edges),
        },
        "elements": {
            "nodes": elements_nodes,
            "edges": elements_edges,
        },
    }


def build_index_html(*, json_filename: str = "workflow.json") -> str:
    return f"""<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <title>Workflow Graph (Graphviz preset + Cytoscape)</title>
  <style>
    :root {{
      color-scheme: light dark;
      font-family: Inter, Segoe UI, Roboto, Helvetica, Arial, sans-serif;
    }}
    body {{
      margin: 0;
      display: grid;
      grid-template-columns: 320px 1fr;
      height: 100vh;
    }}
    #sidebar {{
      border-right: 1px solid #8884;
      padding: 12px;
      overflow: auto;
    }}
    #cy {{
      width: 100%;
      height: 100%;
      display: block;
    }}
    input, select, button {{
      width: 100%;
      margin: 0.3rem 0;
      padding: 0.45rem;
      box-sizing: border-box;
    }}
    .muted {{
      opacity: 0.8;
      font-size: 0.9rem;
    }}
  </style>
</head>
<body>
  <aside id="sidebar">
    <h3>Workflow (Graphviz preset)</h3>
    <div id="summary" class="muted">Loading...</div>
    <label for="repoFilter">Repo filter</label>
    <select id="repoFilter">
      <option value="__all__">All repos</option>
    </select>
    <label for="searchBox">Search issues</label>
    <input id="searchBox" type="text" placeholder="issue #, title..." />
    <button id="resetView">Reset view</button>
    <p class="muted">
      Positions come from Graphviz (`dot -Tjson`) and are rendered with Cytoscape preset layout.
      Click issue nodes to open GitHub.
    </p>
  </aside>
  <main id="cy"></main>

  <script src="https://unpkg.com/cytoscape@3.30.2/dist/cytoscape.min.js"></script>
  <script>
    async function init() {{
      const res = await fetch({json.dumps(json_filename)});
      if (!res.ok) {{
        throw new Error(`Failed to load {json_filename}: ${{res.status}}`);
      }}
      const graph = await res.json();
      const nodes = graph.elements.nodes || [];
      const edges = graph.elements.edges || [];

      const cy = cytoscape({{
        container: document.getElementById('cy'),
        elements: [...nodes, ...edges],
        style: [
          {{
            selector: 'node',
            style: {{
              'label': 'data(label)',
              'font-size': 11,
              'text-wrap': 'wrap',
              'text-max-width': 220,
              'text-valign': 'center',
              'text-halign': 'center',
              'shape': 'round-rectangle',
            }}
          }},
          {{
            selector: 'node[type = "cluster"]',
            style: {{
              'background-opacity': 0.06,
              'border-width': 1.4,
              'border-style': 'dashed',
              'border-color': '#666',
              'font-size': 12,
              'font-weight': 600,
              'text-valign': 'top',
              'text-halign': 'center',
              'text-wrap': 'wrap',
              'text-max-width': 260,
              'text-margin-y': -8,
              'padding': '14px',
            }}
          }},
          {{
            selector: 'node[type = "issue"]',
            style: {{
              'background-opacity': 1,
              'background-color': '#6DA6FF',
              'font-size': 10,
              'text-wrap': 'wrap',
              // Match Graphviz layout: node bbox from `dot -Tjson` (gv_width/gv_height), not wrapped-label bounds.
              'text-max-width': (ele) => {{
                const w = ele.data('gv_width');
                if (w == null || Number.isNaN(Number(w))) return 220;
                return Math.max(40, Number(w) - 24);
              }},
              'width': (ele) => {{
                const w = ele.data('gv_width');
                return (w != null && !Number.isNaN(Number(w))) ? Number(w) : 'label';
              }},
              'height': (ele) => {{
                const h = ele.data('gv_height');
                return (h != null && !Number.isNaN(Number(h))) ? Number(h) : 'label';
              }},
              'padding': '12px',
            }}
          }},
          {{
            selector: 'edge',
            style: {{
              'curve-style': 'bezier',
              'target-arrow-shape': 'triangle',
              'width': 1.6,
              'line-color': '#666',
              'target-arrow-color': '#666',
              'label': 'data(label)',
              'font-size': 9,
              'text-background-opacity': 1,
              'text-background-color': '#fff',
              'text-background-padding': 1,
            }}
          }},
          {{
            selector: '.hidden',
            style: {{
              'display': 'none'
            }}
          }}
        ],
        layout: {{
          name: 'preset',
          fit: true,
          padding: 30
        }}
      }});

      const summary = document.getElementById('summary');
      summary.textContent = `${{graph.meta.issue_count}} issues · ${{graph.meta.edge_count}} blocks edges`;

      const repoFilter = document.getElementById('repoFilter');
      for (const repo of graph.meta.repos || []) {{
        const option = document.createElement('option');
        option.value = repo;
        option.textContent = repo;
        repoFilter.appendChild(option);
      }}

      const applyFilters = () => {{
        const selectedRepo = repoFilter.value;
        const query = document.getElementById('searchBox').value.trim().toLowerCase();
        cy.elements().removeClass('hidden');

        if (selectedRepo !== '__all__') {{
          cy.nodes('[type = "issue"]').forEach((node) => {{
            if (node.data('repo') !== selectedRepo) {{
              node.addClass('hidden');
            }}
          }});
        }}

        if (query) {{
          cy.nodes('[type = "issue"]').forEach((node) => {{
            const hay = `${{node.data('number')}} ${{node.data('title')}}`.toLowerCase();
            if (!hay.includes(query)) {{
              node.addClass('hidden');
            }}
          }});
        }}

        cy.edges().forEach((edge) => {{
          if (edge.source().hasClass('hidden') || edge.target().hasClass('hidden')) {{
            edge.addClass('hidden');
          }}
        }});
      }};

      repoFilter.addEventListener('change', applyFilters);
      document.getElementById('searchBox').addEventListener('input', applyFilters);
      document.getElementById('resetView').addEventListener('click', () => {{
        repoFilter.value = '__all__';
        document.getElementById('searchBox').value = '';
        cy.elements().removeClass('hidden');
        cy.fit();
      }});

      cy.on('tap', 'node[type = "issue"]', (evt) => {{
        const url = evt.target.data('url');
        if (url) {{
          window.open(url, '_blank', 'noopener');
        }}
      }});
    }}

    init().catch((error) => {{
      document.getElementById('summary').textContent = String(error);
      console.error(error);
    }});
  </script>
</body>
</html>
"""


def _fetch_workflow_data(
    owner: str,
    repos: list[str],
    token: str,
) -> tuple[list[WorkflowIssue], list[tuple[str, int, str, int]]]:
    issues_by_repo: dict[str, list[dict[str, Any]]] = {}
    label_descriptions_by_repo: dict[str, dict[str, str]] = {}
    for repo in repos:
        issues_by_repo[repo] = fetch_open_issues(owner, repo, token)
        label_descriptions_by_repo[repo] = fetch_label_descriptions(owner, repo, token)

    workflow_issues = build_workflow_issues(issues_by_repo, label_descriptions_by_repo)
    blocks_edges = fetch_blocks_edges(owner, token, workflow_issues)
    return workflow_issues, blocks_edges


def write_cytoscape_site(
    docs_dir: Path,
    *,
    token: str,
    owner: str = "Teal-Insights",
    repos: list[str] | None = None,
    layout: str = "LR",
    dot_bin: str | None = None,
) -> dict[str, Any]:
    """Write Cytoscape preset JSON and HTML under *docs_dir*.

    Returns:
        The payload ``meta`` block (issue, cluster, and edge counts).
    """
    repo_list = list(DEFAULT_REPOS if repos is None else repos)
    workflow_issues, blocks_edges = _fetch_workflow_data(owner, repo_list, token)
    dot_text = build_dot(workflow_issues, blocks_edges, owner=owner, rankdir=layout)

    dot_binary = dot_bin or "dot"
    graphviz_json = _parse_graphviz_json(dot_text, dot_bin=dot_binary)
    payload = build_cytoscape_preset_payload(
        owner=owner,
        workflow_issues=workflow_issues,
        graphviz_json=graphviz_json,
    )

    json_output = docs_dir / "workflow.json"
    html_output = docs_dir / "index.html"

    docs_dir.mkdir(parents=True, exist_ok=True)
    json_output.write_text(json.dumps(payload, indent=2), encoding="utf-8")
    html_output.write_text(build_index_html(json_filename=json_output.name), encoding="utf-8")
    return dict(payload["meta"])


def register_workflow_resources(quarto_yml: Path) -> bool:
    """Ensure ``_quarto.yml`` copies the workflow site into ``_site/``.

    Returns:
        True when the file was updated.
    """
    if not quarto_yml.is_file():
        return False

    config = yaml.safe_load(quarto_yml.read_text(encoding="utf-8")) or {}
    project = config.setdefault("project", {})

    resources = project.setdefault("resources", [])
    if isinstance(resources, str):
        resources = [resources]
        project["resources"] = resources

    render = project.setdefault("render", ["**"])
    if isinstance(render, str):
        render = [render]
        project["render"] = render

    changed = False
    if WORKFLOW_RESOURCE not in resources:
        resources.append(WORKFLOW_RESOURCE)
        changed = True
    if WORKFLOW_HTML_EXCLUDE not in render:
        render.append(WORKFLOW_HTML_EXCLUDE)
        changed = True

    if changed:
        quarto_yml.write_text(yaml.safe_dump(config, sort_keys=False), encoding="utf-8")
    return changed


def copy_workflow_fallback(*, source_dir: Path, output_dir: Path) -> bool:
    """Copy a checked-in workflow snapshot into the build tree."""
    index_src = source_dir / "index.html"
    json_src = source_dir / "workflow.json"
    if not index_src.is_file() or not json_src.is_file():
        return False

    output_dir.mkdir(parents=True, exist_ok=True)
    shutil.copy2(index_src, output_dir / "index.html")
    shutil.copy2(json_src, output_dir / "workflow.json")
    return True


def _workflow_fallback_dirs(project_root: Path) -> list[Path]:
    return [
        project_root / "custom" / "workflow",
        project_root / "docs",
    ]


def main() -> None:
    build_dir = Path.cwd()
    project_root = _find_project_root(build_dir)
    output_dir = build_dir / WORKFLOW_DIR
    quarto_yml = build_dir / "_quarto.yml"

    generated = False
    try:
        token = _resolve_token(None)
    except RuntimeError:
        token = None

    if token is not None:
        meta = write_cytoscape_site(output_dir, token=token)
        print(f"Generated workflow site in {output_dir.relative_to(build_dir)}/")
        print(f"Issues included: {meta['issue_count']}")
        generated = True
    else:
        for source_dir in _workflow_fallback_dirs(project_root):
            if copy_workflow_fallback(source_dir=source_dir, output_dir=output_dir):
                print(
                    "No GitHub token; copied workflow snapshot from "
                    f"{source_dir.relative_to(project_root)}/"
                )
                generated = True
                break

    if not generated:
        print(
            "Skipping workflow graph (no token and no workflow snapshot fallback). "
            "Set GITHUB_TOKEN or ISSUE_READ_PAT, or commit custom/workflow/ or docs/."
        )
        return

    if register_workflow_resources(quarto_yml):
        print(f"Registered {WORKFLOW_RESOURCE} in {quarto_yml.name}")


if __name__ == "__main__":
    main()
