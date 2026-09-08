"""CLI smoke for `scripts/analyze_solver_mcve.py`."""

from __future__ import annotations

import json
from pathlib import Path

from scripts.analyze_solver_mcve import main


def test_analyze_solver_mcve_cli(tmp_path: Path, capsys) -> None:
    path = tmp_path / "fragment.json"
    path.write_text(
        json.dumps(
            {
                "schema_version": "1.0.0",
                "kind": "may_cycle_solver_mcve",
                "scc_index": 0,
                "scc_members": ["Sheet1!A1", "Sheet1!B1"],
                "nodes": {
                    "Sheet1!A1": {
                        "in_scc": True,
                        "in_guard_cone": False,
                        "is_leaf": False,
                        "normalized_formula": "=Sheet1!B1",
                        "value": 1,
                    },
                    "Sheet1!B1": {
                        "in_scc": True,
                        "in_guard_cone": False,
                        "is_leaf": False,
                        "normalized_formula": "=Sheet1!A1",
                        "value": 1,
                    },
                },
                "intra_scc_edges": [
                    {"from": "Sheet1!A1", "to": "Sheet1!B1", "guard": None},
                    {"from": "Sheet1!B1", "to": "Sheet1!A1", "guard": None},
                ],
                "leaf_constraints": {},
            }
        ),
        encoding="utf-8",
    )
    main([str(path)])
    out = capsys.readouterr().out
    assert "may_cycle_solver_mcve" in out
    assert "must=True" in out
