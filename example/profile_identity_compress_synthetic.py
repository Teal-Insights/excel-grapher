#!/usr/bin/env python3
"""
Build an in-memory linear identity-transit chain and profile ``compress_identity_transits()``.

No workbook I/O — useful to isolate compression algorithm cost (e.g. repeated full sorts).

    uv run python example/profile_identity_compress_synthetic.py --chain 3000 \\
        --cprofile-out /tmp/syn_compress.prof --cprofile-print 30
"""

from __future__ import annotations

import argparse
import cProfile
import pstats
import sys
from pathlib import Path

from excel_grapher.grapher.dependency_provenance import DependencyCause, EdgeProvenance
from excel_grapher.grapher.graph import DependencyGraph
from excel_grapher.grapher.node import Node


def _make_node(key: str, formula: str | None, normalized: str | None, *, is_leaf: bool = False) -> Node:
    sheet, rest = key.split("!", 1)
    if sheet.startswith("'"):
        sheet = sheet[1:-1]
    col = "".join(c for c in rest if c.isalpha())
    row = int("".join(c for c in rest if c.isdigit()))
    return Node(
        sheet=sheet,
        column=col,
        row=row,
        formula=formula,
        normalized_formula=normalized,
        value=None,
        is_leaf=is_leaf,
    )


def linear_identity_chain(length: int) -> DependencyGraph:
    """A1 -> A2 -> ... -> A(length+1) leaf; each formula is a single cell ref to the next.

    A ``Sheet1!Z1`` head cell references ``A1`` so the root transit still has a dependent and
    can be compressed away (otherwise the final root would be skipped).
    """
    g = DependencyGraph()
    dr = DependencyCause.direct_ref
    leaf_k = f"Sheet1!A{length + 1}"
    g.add_node(_make_node(leaf_k, None, None, is_leaf=True))
    for k in range(length, 0, -1):
        key = f"Sheet1!A{k}"
        dep = f"Sheet1!A{k + 1}"
        formula = f"={dep}"
        g.add_node(_make_node(key, formula, formula))
        i = formula.index(dep)
        sp = ((i, i + len(dep)),)
        g.add_edge(
            key,
            dep,
            provenance=EdgeProvenance(
                causes=frozenset({dr}),
                direct_sites_formula=sp,
                direct_sites_normalized=sp,
            ),
        )
    head = "Sheet1!Z1"
    a1 = "Sheet1!A1"
    hf = f"={a1}"
    g.add_node(_make_node(head, hf, hf))
    hi = hf.index(a1)
    hsp = ((hi, hi + len(a1)),)
    g.add_edge(
        head,
        a1,
        provenance=EdgeProvenance(
            causes=frozenset({dr}),
            direct_sites_formula=hsp,
            direct_sites_normalized=hsp,
        ),
    )
    return g


def main() -> int:
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument("--chain", type=int, default=2000, help="Number of transit formulas (default 2000).")
    p.add_argument("--cprofile-out", type=Path, default=None, metavar="FILE")
    p.add_argument("--cprofile-print", type=int, default=0, metavar="N")
    args = p.parse_args()
    if args.chain < 1:
        print("--chain must be >= 1", file=sys.stderr)
        return 1

    g = linear_identity_chain(args.chain)
    prof = cProfile.Profile()
    if args.cprofile_out is not None:
        prof.enable()
    removed = g.compress_identity_transits()
    if args.cprofile_out is not None:
        prof.disable()
        args.cprofile_out.parent.mkdir(parents=True, exist_ok=True)
        prof.dump_stats(str(args.cprofile_out))
        print(f"cProfile stats written to {args.cprofile_out.resolve()}", file=sys.stderr)
        if args.cprofile_print > 0:
            pstats.Stats(prof).sort_stats(pstats.SortKey.CUMULATIVE).print_stats(args.cprofile_print)

    print(f"Chain length: {args.chain}")
    print(f"Removed transits: {len(removed)}")
    print(f"Remaining nodes: {len(g)}")
    assert len(removed) == args.chain
    assert len(g) == 2
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
