#!/usr/bin/env python3
from __future__ import annotations

import argparse
import json
import subprocess
from pathlib import Path
from typing import Any

from scripts.build_workflow_dot_from_github import (
    DEFAULT_REPOS,
    WorkflowIssue,
    _resolve_token,
    build_dot,
    build_workflow_issues,
    fetch_blocks_edges,
    fetch_label_descriptions,
    fetch_open_issues,
)


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
        cluster_id: {n for n in (objects_by_gvid[cluster_id].get("nodes", []) or []) if isinstance(n, int)}
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

    cluster_node_id_by_gvid = {gvid: f"cluster::{objects_by_gvid[gvid]['name']}" for gvid in cluster_gvids}
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
            "issue_count": len([node for node in elements_nodes if node["data"]["type"] == "issue"]),
            "cluster_count": len([node for node in elements_nodes if node["data"]["type"] == "cluster"]),
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


def main() -> None:
    parser = argparse.ArgumentParser(description="Generate dot-driven Cytoscape docs workflow page.")
    parser.add_argument("--owner", default="Teal-Insights")
    parser.add_argument("--repos", nargs="+", default=DEFAULT_REPOS)
    parser.add_argument("--layout", choices=["TB", "LR"], default="TB")
    parser.add_argument("--github-token")
    parser.add_argument("--docs-dir", type=Path, default=Path("docs"))
    parser.add_argument("--dot-bin", default=None, help="Path to Graphviz dot executable.")
    args = parser.parse_args()

    token = _resolve_token(args.github_token)
    docs_dir = args.docs_dir
    json_output = docs_dir / "workflow.json"
    html_output = docs_dir / "index.html"

    issues_by_repo: dict[str, list[dict[str, Any]]] = {}
    label_descriptions_by_repo: dict[str, dict[str, str]] = {}
    for repo in args.repos:
        issues_by_repo[repo] = fetch_open_issues(args.owner, repo, token)
        label_descriptions_by_repo[repo] = fetch_label_descriptions(args.owner, repo, token)

    workflow_issues = build_workflow_issues(issues_by_repo, label_descriptions_by_repo)
    blocks_edges = fetch_blocks_edges(args.owner, token, workflow_issues)
    dot_text = build_dot(workflow_issues, blocks_edges, owner=args.owner, rankdir=args.layout)

    dot_bin = args.dot_bin or "dot"
    graphviz_json = _parse_graphviz_json(dot_text, dot_bin=dot_bin)
    payload = build_cytoscape_preset_payload(
        owner=args.owner,
        workflow_issues=workflow_issues,
        graphviz_json=graphviz_json,
    )

    json_output.parent.mkdir(parents=True, exist_ok=True)
    json_output.write_text(json.dumps(payload, indent=2), encoding="utf-8")

    html_output.parent.mkdir(parents=True, exist_ok=True)
    html_output.write_text(build_index_html(json_filename=json_output.name), encoding="utf-8")

    print(f"Wrote workflow JSON to {json_output}")
    print(f"Wrote workflow HTML to {html_output}")
    print(f"Issues included: {payload['meta']['issue_count']}")
    print(f"Clusters included: {payload['meta']['cluster_count']}")
    print(f"Blocks edges included: {payload['meta']['edge_count']}")


if __name__ == "__main__":
    main()
