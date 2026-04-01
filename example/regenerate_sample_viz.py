#!/usr/bin/env python3
"""
Refresh the LIC-DSF sample lightweight viz HTML.

Default (fast): re-embed the current package ``lightweight_viz_template.html`` into
``example/data/lic-dsf-template-sample-exported-viz.html`` while keeping the existing
inline ``window.__VIZ_DATA__`` payload (no graph rebuild).

Use ``--full`` to rebuild from ``example/.cache/...-dependency-graph.pkl`` (slow, large
RAM; re-runs ``to_lightweight_viz`` and serializes ~tens of MB of JSON).

To profile where time goes (with timeout / SIGTERM-safe flush), use
``example/profile_lightweight_viz.py`` (see its docstring).
"""

from __future__ import annotations

import argparse
import pickle
import re
import sys
from importlib import resources
from pathlib import Path

_EXAMPLE_DIR = Path(__file__).resolve().parent
_REPO_ROOT = _EXAMPLE_DIR.parent
_DEFAULT_HTML = _EXAMPLE_DIR / "data" / "lic-dsf-template-sample-exported-viz.html"
_DEFAULT_CACHE = _EXAMPLE_DIR / ".cache" / "lic-dsf-template-2025-08-12-dependency-graph.pkl"


def _package_template() -> str:
    import excel_grapher.grapher.lightweight_viz as lv

    pkg = lv.__package__ or "excel_grapher.grapher"
    return resources.files(pkg).joinpath("lightweight_viz_template.html").read_text(encoding="utf-8")


def _extract_title(html_head: str) -> str:
    m = re.search(r"<title>([^<]*)</title>", html_head, re.I)
    return m.group(1).strip() if m else "Workbook dependency graph"


def _extract_bootstrap_lines(sample_path: Path) -> tuple[str, str]:
    """
    Return (bootstrap_line, sidecar_line) from the first inline script in the sample.
    """
    bootstrap_line = ""
    sidecar_line = ""
    in_script = False
    with sample_path.open("r", encoding="utf-8") as f:
        for line in f:
            s = line.strip()
            if not in_script:
                if s == "<script>":
                    in_script = True
                continue
            if s == "</script>":
                break
            if "window.__VIZ_DATA__" in line:
                bootstrap_line = line.rstrip("\n")
            elif "window.__VIZ_DATA_URL__" in line:
                sidecar_line = line.rstrip("\n")
                break
    if not bootstrap_line or not sidecar_line:
        raise ValueError(
            f"Could not parse bootstrap from {sample_path}: "
            "expected window.__VIZ_DATA__ and window.__VIZ_DATA_URL__ in first <script> block."
        )
    return bootstrap_line, sidecar_line


def refresh_template_only(sample_html: Path) -> None:
    with sample_html.open("r", encoding="utf-8") as f:
        head = f.read(65536)
    title = _extract_title(head)

    boot, side = _extract_bootstrap_lines(sample_html)
    tpl = _package_template()
    out_html = (
        tpl.replace("__TITLE__", title)
        .replace("/*__BOOTSTRAP__*/", boot.strip())
        .replace("/*__SIDECAR__*/", side.strip())
    )
    sample_html.write_text(out_html, encoding="utf-8")


def full_rebuild(sample_html: Path, cache_pkl: Path, budget_mb: int) -> None:
    sys.path.insert(0, str(_REPO_ROOT))
    from excel_grapher.grapher.graph import DependencyGraph
    from excel_grapher.grapher.lightweight_viz import to_lightweight_viz, write_lightweight_viz_html

    with cache_pkl.open("rb") as f:
        blob = pickle.load(f)
    if not isinstance(blob, tuple) or len(blob) != 2:
        raise SystemExit("Pickle must be (meta, graph) tuple")
    _, graph = blob
    if not isinstance(graph, DependencyGraph):
        raise SystemExit("Pickle graph is not a DependencyGraph")

    payload = to_lightweight_viz(graph)
    write_lightweight_viz_html(
        payload,
        sample_html,
        title="LIC-DSF Template dependency graph",
        data_mode="inline",
        inline_size_budget_mb=budget_mb,
    )


def main() -> None:
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument(
        "--full",
        action="store_true",
        help="Rebuild from graph pickle (slow); default is template-only refresh.",
    )
    p.add_argument(
        "--output",
        type=Path,
        default=_DEFAULT_HTML,
        help=f"Output HTML path (default: {_DEFAULT_HTML})",
    )
    p.add_argument(
        "--cache",
        type=Path,
        default=_DEFAULT_CACHE,
        help=f"Graph pickle for --full (default: {_DEFAULT_CACHE})",
    )
    p.add_argument(
        "--inline-budget-mb",
        type=int,
        default=512,
        help="Inline JSON size budget for --full (default: 512)",
    )
    args = p.parse_args()
    out = args.output.resolve()
    if args.full:
        cache = args.cache.resolve()
        if not cache.is_file():
            raise SystemExit(f"Missing cache: {cache}")
        print("Full rebuild from pickle (this may take many minutes)...", flush=True)
        full_rebuild(out, cache, args.inline_budget_mb)
        print(f"Wrote {out} ({out.stat().st_size // 1024 // 1024} MiB)", flush=True)
        return

    if not out.is_file():
        raise SystemExit(
            f"Missing {out}; run with --full once after extract_graph_cached.py, "
            "or create the sample HTML first."
        )
    refresh_template_only(out)
    print(f"Refreshed package template into {out} (data payload unchanged).", flush=True)


if __name__ == "__main__":
    main()
