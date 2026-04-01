#!/usr/bin/env python3
"""
Profile ``regenerate_sample_viz.py --full`` style work (pickle → to_lightweight_viz → JSON).

**Why a plain ``signal.alarm`` / ``SIGALRM`` timeout can look “stuck”:** Python only runs
signal handlers between bytecode instructions. A long ``pickle.load`` or ``json.dumps``
runs mostly in C, so your alarm may not fire until that call returns—possibly never if you
misread progress as a hang.

**``--timeout N`` (when N > 0):** arms ``faulthandler.dump_traceback_later(N, exit=True)``
(Unix real-time timer). When it fires, Python dumps **all thread tracebacks** to stderr
and then ``_exit(1)``. That can interrupt blocking C work, but **does not** run ``finally``
or write ``.prof`` files—use it to see *where* time went. For a flushable profile on
interrupt, use ``SIGTERM``/``SIGINT`` while the process is in Python code, or narrow work
with ``--stop-after``.

Shell ``timeout -k 10 300 ...`` still helps: SIGKILL ends a truly stuck process after the
grace period (no profile).

Outputs (by default next to the pickle):

  - ``lightweight-viz.prof`` — binary stats for ``snakeviz`` / ``python -m pstats``
  - ``lightweight-viz.prof.txt`` — top functions by cumulative time (also printed to stderr)
"""

from __future__ import annotations

import argparse
import cProfile
import faulthandler
import io
import pickle
import pstats
import signal
import sys
from pathlib import Path
from typing import Literal

_EXAMPLE_DIR = Path(__file__).resolve().parent
_REPO_ROOT = _EXAMPLE_DIR.parent
_DEFAULT_CACHE = _EXAMPLE_DIR / ".cache" / "lic-dsf-template-2025-08-12-dependency-graph.pkl"


def main() -> None:
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument(
        "--cache",
        type=Path,
        default=_DEFAULT_CACHE,
        help=f"Graph pickle (default: {_DEFAULT_CACHE})",
    )
    p.add_argument(
        "--timeout",
        type=int,
        default=0,
        help=(
            "Wall-clock seconds; >0 arms faulthandler.dump_traceback_later(..., exit=True) "
            "(works during long C calls like pickle; dumps stacks then exits—no .prof flush). "
            "0 disables. Default: 0"
        ),
    )
    p.add_argument(
        "--binary-profile",
        type=Path,
        default=None,
        help="Where to write binary .prof (default: <cache-dir>/lightweight-viz.prof)",
    )
    p.add_argument(
        "--stop-after",
        choices=("pickle", "viz", "json"),
        default="json",
        help="Stop after this phase (default: json = full pipeline)",
    )
    p.add_argument(
        "--top",
        type=int,
        default=60,
        help="Lines of pstats text report (default: 60)",
    )
    args = p.parse_args()

    cache = args.cache.resolve()
    if not cache.is_file():
        raise SystemExit(f"Missing cache: {cache}")

    prof_path = args.binary_profile
    if prof_path is None:
        prof_path = cache.parent / "lightweight-viz.prof"
    prof_path = prof_path.resolve()
    prof_path.parent.mkdir(parents=True, exist_ok=True)
    txt_path = prof_path.with_suffix(".prof.txt")

    prof = cProfile.Profile()
    finalized = False
    exit_reason = "finished"

    def finalize(*, reason: str) -> None:
        nonlocal finalized
        if finalized:
            return
        finalized = True
        faulthandler.cancel_dump_traceback_later()
        prof.disable()
        prof.dump_stats(str(prof_path))
        buf = io.StringIO()
        st = pstats.Stats(prof, stream=buf)
        st.sort_stats(pstats.SortKey.CUMULATIVE)
        st.print_stats(args.top)
        report = buf.getvalue()
        header = f"# profile_lightweight_viz ({reason})\n"
        txt_path.write_text(header + report, encoding="utf-8")
        print(header, file=sys.stderr, end="")
        print(report, file=sys.stderr, end="")
        print(f"[profile_lightweight_viz] wrote {prof_path} and {txt_path}", file=sys.stderr)

    _sig_names: dict[int, str] = {
        signal.SIGTERM: "SIGTERM",
        signal.SIGINT: "SIGINT",
    }

    def on_signal(signum: int, _frame: object | None) -> None:
        nonlocal exit_reason
        exit_reason = _sig_names.get(signum, f"signal_{signum}")
        print(
            f"\n[profile_lightweight_viz] caught {exit_reason}; will flush profile in finally…",
            file=sys.stderr,
        )
        raise SystemExit(128 + signum if signum != signal.SIGINT else 130)

    signal.signal(signal.SIGTERM, on_signal)
    signal.signal(signal.SIGINT, on_signal)

    if args.timeout > 0:
        faulthandler.dump_traceback_later(args.timeout, exit=True)
        print(
            f"[profile_lightweight_viz] hard wall timeout: {args.timeout}s "
            "(traceback + _exit on fire; see module docstring)",
            file=sys.stderr,
        )

    sys.path.insert(0, str(_REPO_ROOT))
    from excel_grapher.grapher.graph import DependencyGraph
    from excel_grapher.grapher.lightweight_viz import (
        serialize_lightweight_viz_json,
        to_lightweight_viz,
    )

    phase: Literal["pickle", "viz", "json"] | None = None
    try:
        prof.enable()
        with cache.open("rb") as f:
            blob = pickle.load(f)
        if not isinstance(blob, tuple) or len(blob) != 2:
            raise SystemExit("Pickle must be (meta, graph) tuple")
        _, graph = blob
        if not isinstance(graph, DependencyGraph):
            raise SystemExit("Pickle graph is not a DependencyGraph")
        phase = "pickle"
        if args.stop_after == "pickle":
            return

        payload = to_lightweight_viz(graph)
        phase = "viz"
        print(
            f"[profile_lightweight_viz] nodes={payload.stats.node_count} "
            f"local_edges={payload.stats.local_edge_count}",
            file=sys.stderr,
        )
        if args.stop_after == "viz":
            return

        s = serialize_lightweight_viz_json(payload)
        phase = "json"
        print(f"[profile_lightweight_viz] json_chars={len(s)}", file=sys.stderr)
    except SystemExit:
        raise
    except BaseException as e:
        exit_reason = f"exception_after:{phase or 'start'}:{type(e).__name__}"
        raise
    finally:
        if not finalized:
            finalize(reason=exit_reason)


if __name__ == "__main__":
    main()
