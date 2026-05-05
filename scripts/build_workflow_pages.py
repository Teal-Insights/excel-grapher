#!/usr/bin/env python3
from __future__ import annotations

import argparse
import json
from pathlib import Path
from typing import Any

from scripts.build_workflow_dot_from_github import (
    DEFAULT_REPOS,
    WorkflowIssue,
    _resolve_token,
    build_workflow_issues,
    fetch_blocks_edges,
    fetch_label_descriptions,
    fetch_open_issues,
)


def issue_node_id(repo: str, number: int) -> str:
    return f"issue::{repo}#{number}"


def repo_node_id(repo: str) -> str:
    return f"repo::{repo}"


def pre_group_node_id(repo: str, pre_key: str) -> str:
    return f"group::{repo}::{pre_key}"


def post_group_node_id(repo: str, pre_key: str, post_key: str) -> str:
    return f"group::{repo}::{pre_key}/{post_key}"


def build_workflow_json(
    *,
    owner: str,
    issues: list[WorkflowIssue],
    blocks_edges: list[tuple[str, int, str, int]],
) -> dict[str, Any]:
    repos = sorted({issue.repo for issue in issues})
    nodes: list[dict[str, Any]] = []
    edges: list[dict[str, Any]] = []

    seen_node_ids: set[str] = set()

    for repo in repos:
        node_id = repo_node_id(repo)
        nodes.append(
            {
                "data": {
                    "id": node_id,
                    "label": f"{repo} GitHub issues",
                    "type": "repo",
                    "repo": repo,
                }
            }
        )
        seen_node_ids.add(node_id)

    for issue in issues:
        parent_id = repo_node_id(issue.repo)

        if issue.pre_group_key:
            pre_id = pre_group_node_id(issue.repo, issue.pre_group_key)
            if pre_id not in seen_node_ids:
                nodes.append(
                    {
                        "data": {
                            "id": pre_id,
                            "label": issue.pre_group_label or issue.pre_group_key,
                            "type": "pre_group",
                            "repo": issue.repo,
                            "group_key": issue.pre_group_key,
                            "parent": parent_id,
                        }
                    }
                )
                seen_node_ids.add(pre_id)
            parent_id = pre_id

        if issue.pre_group_key and issue.post_group_key:
            post_id = post_group_node_id(issue.repo, issue.pre_group_key, issue.post_group_key)
            if post_id not in seen_node_ids:
                nodes.append(
                    {
                        "data": {
                            "id": post_id,
                            "label": issue.post_group_label or issue.post_group_key,
                            "type": "post_group",
                            "repo": issue.repo,
                            "group_key": f"{issue.pre_group_key}/{issue.post_group_key}",
                            "parent": pre_id,
                        }
                    }
                )
                seen_node_ids.add(post_id)
            parent_id = post_id

        issue_id = issue_node_id(issue.repo, issue.number)
        nodes.append(
            {
                "data": {
                    "id": issue_id,
                    "label": f"#{issue.number} · {issue.title}",
                    "type": "issue",
                    "repo": issue.repo,
                    "number": issue.number,
                    "title": issue.title,
                    "url": f"https://github.com/{owner}/{issue.repo}/issues/{issue.number}",
                    "parent": parent_id,
                }
            }
        )
        seen_node_ids.add(issue_id)

    for blocker_repo, blocker_number, blocked_repo, blocked_number in blocks_edges:
        source_id = issue_node_id(blocker_repo, blocker_number)
        target_id = issue_node_id(blocked_repo, blocked_number)
        if source_id not in seen_node_ids or target_id not in seen_node_ids:
            continue
        edges.append(
            {
                "data": {
                    "id": f"edge::{source_id}->{target_id}",
                    "source": source_id,
                    "target": target_id,
                    "type": "blocks",
                    "label": "blocks",
                }
            }
        )

    return {
        "meta": {
            "owner": owner,
            "repos": repos,
            "issue_count": len([node for node in nodes if node["data"]["type"] == "issue"]),
            "edge_count": len(edges),
        },
        "elements": {
            "nodes": nodes,
            "edges": edges,
        },
    }


def build_index_html(*, json_filename: str = "workflow.json") -> str:
    return f"""<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <title>Workflow Graph</title>
  <style>
    :root {{
      color-scheme: light dark;
      font-family: Inter, Segoe UI, Roboto, Helvetica, Arial, sans-serif;
    }}
    body {{
      margin: 0;
      display: grid;
      grid-template-columns: 280px 1fr;
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
    <h3>Workflow</h3>
    <div id="summary" class="muted">Loading...</div>
    <label for="repoFilter">Repo filter</label>
    <select id="repoFilter">
      <option value="__all__">All repos</option>
    </select>
    <label for="searchBox">Search issues</label>
    <input id="searchBox" type="text" placeholder="issue #, title..." />
    <label for="layoutSelect">Layout</label>
    <select id="layoutSelect">
      <option value="klay" selected>klay (default)</option>
      <option value="fcose">fcose</option>
      <option value="breadthfirst">breadthfirst</option>
    </select>
    <button id="resetView">Reset view</button>
    <p class="muted">Click an issue node to open it in GitHub.</p>
  </aside>
  <main id="cy"></main>

  <script src="https://unpkg.com/cytoscape@3.30.2/dist/cytoscape.min.js"></script>
  <script src="https://unpkg.com/layout-base@2.0.1/layout-base.js"></script>
  <script src="https://unpkg.com/cose-base@2.2.0/cose-base.js"></script>
  <script src="https://unpkg.com/cytoscape-fcose@2.2.0/cytoscape-fcose.js"></script>
  <script src="https://unpkg.com/klayjs@0.4.1/klay.js"></script>
  <script src="https://unpkg.com/cytoscape-klay@3.1.4/cytoscape-klay.js"></script>
  <script>
    async function init() {{
      cytoscape.use(cytoscapeFcose);
      cytoscape.use(cytoscapeKlay);

      function runLayout(cy, layoutName) {{
        if (layoutName === 'klay') {{
          cy.layout({{
            name: 'klay',
            klay: {{
              direction: 'DOWN',
              edgeRouting: 'ORTHOGONAL',
              spacing: 36,
              inLayerSpacingFactor: 1.8,
              nodeLayering: 'NETWORK_SIMPLEX',
              thoroughness: 7
            }},
            fit: true,
            padding: 32,
            animate: false
          }}).run();
          return;
        }}

        if (layoutName === 'fcose') {{
          cy.layout({{
            name: 'fcose',
            quality: 'proof',
            randomize: false,
            animate: false,
            packComponents: true,
            nodeRepulsion: 9000,
            idealEdgeLength: 120,
            edgeElasticity: 0.08,
            nestingFactor: 0.9,
            gravity: 0.25,
            gravityCompound: 1.0,
            numIter: 1800
          }}).run();
          return;
        }}

        cy.layout({{
          name: 'breadthfirst',
          directed: true,
          spacingFactor: 1.2,
          animate: false,
          padding: 30
        }}).run();
      }}

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
              'background-color': '#8DA0CB',
              'text-valign': 'center',
              'text-halign': 'center',
              'shape': 'round-rectangle',
              'padding': '8px',
            }}
          }},
          {{
            selector: 'node[type = "repo"]',
            style: {{
              'label': 'data(label)',
              'background-opacity': 0.08,
              'border-width': 2,
              'border-style': 'solid',
              'border-color': '#555',
              'font-size': 14,
              'font-weight': 700,
              'text-valign': 'top',
              'text-halign': 'center',
              'text-wrap': 'wrap',
              'text-max-width': 260,
              'text-margin-y': -8,
              'padding': '18px',
            }}
          }},
          {{
            selector: 'node[type = "pre_group"], node[type = "post_group"]',
            style: {{
              'label': 'data(label)',
              'background-opacity': 0.06,
              'border-width': 1,
              'border-style': 'dashed',
              'border-color': '#666',
              'font-size': 12,
              'font-weight': 600,
              'text-valign': 'top',
              'text-halign': 'center',
              'text-wrap': 'wrap',
              'text-max-width': 240,
              'text-margin-y': -6,
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
              'text-max-width': 220,
              'width': 'label',
              'height': 'label',
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
        ]
      }});

      runLayout(cy, 'klay');

      const summary = document.getElementById('summary');
      summary.textContent = `${{graph.meta.issue_count}} issues · ${{graph.meta.edge_count}} blocks edges`;

      const repoFilter = document.getElementById('repoFilter');
      const layoutSelect = document.getElementById('layoutSelect');
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
          cy.nodes().forEach((node) => {{
            const repo = node.data('repo');
            if (repo && repo !== selectedRepo) {{
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
      layoutSelect.addEventListener('change', () => runLayout(cy, layoutSelect.value));
      document.getElementById('resetView').addEventListener('click', () => {{
        repoFilter.value = '__all__';
        layoutSelect.value = 'klay';
        document.getElementById('searchBox').value = '';
        cy.elements().removeClass('hidden');
        runLayout(cy, layoutSelect.value);
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
    parser = argparse.ArgumentParser(
        description="Generate docs/workflow.json and docs/index.html for GitHub Pages."
    )
    parser.add_argument("--owner", default="Teal-Insights")
    parser.add_argument(
        "--repos",
        nargs="+",
        default=DEFAULT_REPOS,
        help="Repository names under --owner.",
    )
    parser.add_argument("--github-token")
    parser.add_argument("--docs-dir", type=Path, default=Path("docs"))
    parser.add_argument("--json-output", type=Path, default=None)
    parser.add_argument("--html-output", type=Path, default=None)
    args = parser.parse_args()

    token = _resolve_token(args.github_token)
    docs_dir = args.docs_dir
    json_output = args.json_output or docs_dir / "workflow.json"
    html_output = args.html_output or docs_dir / "index.html"

    issues_by_repo: dict[str, list[dict[str, Any]]] = {}
    label_descriptions_by_repo: dict[str, dict[str, str]] = {}
    for repo in args.repos:
        issues_by_repo[repo] = fetch_open_issues(args.owner, repo, token)
        label_descriptions_by_repo[repo] = fetch_label_descriptions(args.owner, repo, token)

    workflow_issues = build_workflow_issues(issues_by_repo, label_descriptions_by_repo)
    blocks_edges = fetch_blocks_edges(args.owner, token, workflow_issues)
    payload = build_workflow_json(owner=args.owner, issues=workflow_issues, blocks_edges=blocks_edges)

    json_output.parent.mkdir(parents=True, exist_ok=True)
    json_output.write_text(json.dumps(payload, indent=2), encoding="utf-8")

    html_output.parent.mkdir(parents=True, exist_ok=True)
    html_output.write_text(
        build_index_html(json_filename=json_output.name),
        encoding="utf-8",
    )

    print(f"Wrote workflow JSON to {json_output}")
    print(f"Wrote workflow HTML to {html_output}")
    print(f"Issues included: {payload['meta']['issue_count']}")
    print(f"Blocks edges included: {payload['meta']['edge_count']}")


if __name__ == "__main__":
    main()
