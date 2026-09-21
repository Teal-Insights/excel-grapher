"""Greenfield extract using bindings domains, not a Python constraint table.

The walk starts from placeholder shards plus an output root, lists dynamic-ref
leaves with `DynamicRefConfig.from_bindings`, and never sets
`use_cached_dynamic_refs`.
"""

from __future__ import annotations

import copy
from pathlib import Path

import pytest
import yaml

from excel_grapher.core.cell_types import GreaterThanCell
from excel_grapher.grapher import create_dependency_graph, list_dynamic_ref_constraint_candidates
from excel_grapher.grapher.dynamic_refs import DynamicRefConfig, DynamicRefError
from excel_grapher.series_bindings import (
    SeriesRelationError,
    bootstrap_binding_shards,
    load_series_bindings,
    undomained_leaves,
)
from excel_grapher.series_bindings.normalize import has_output_direction
from excel_grapher.series_bindings.versions import CURRENT_SCHEMA_VERSION
from excel_grapher.series_bindings.workflow import all_series_targets
from tests.paths import INVERTED_TREE_TINY_DSA
from tests.unit.exporter.inverted_tree.helpers import series_entry, write_workbook

_SHARDS = {
    "inputs": "inputs.bindings.yaml",
    "outputs": "outputs.bindings.yaml",
    "internals": "internals.bindings.yaml",
    "constants": "constants.bindings.yaml",
}


def _write_dynamic_workbook(path: Path, *, blank_offset_col: bool = False) -> Path:
    """OFFSET via a defined name, then INDIRECT, plus a static INDEX and a plain leaf."""
    offset_col = "Inputs!H1" if blank_offset_col else "0"
    inputs: dict[str, object] = {
        "B1": 1,
        "C1": "Lookups!Z1",
        "D1": 2,
        "E1": 5,
        "F1": 10.0,
        "G1": 2,
    }
    if blank_offset_col:
        inputs["H1"] = None
    return write_workbook(
        path,
        {
            "Inputs": inputs,
            "Lookups": {
                "A1": 100,
                "A2": "=INDIRECT(Inputs!C1)",
                "B1": 1,
                "B2": 2,
                "C1": 3,
                "C2": 4,
                "Z1": 42,
            },
            "Engine": {
                "A1": f"=OFFSET(BASE,Inputs!B1,{offset_col})",
                "B1": "=Inputs!F1*2",
                "C1": "=IF(Inputs!E1>Inputs!D1,Engine!A1,Engine!B1)",
                "D1": "=INDEX(Lookups!A1:C2,Inputs!G1,2)",
            },
            "Outputs": {"A1": "=Engine!C1+Engine!D1"},
        },
        defined_names={"BASE": "Lookups!$A$1"},
    )


def _scalar(
    series_id: str,
    data_range: str,
    *,
    dtype: str = "int",
    domain: dict[str, object] | None = None,
    relations: list[dict[str, str]] | None = None,
    constant: bool = False,
) -> dict[str, object]:
    entry = series_entry(
        series_id,
        data_range,
        layout="scalar",
        direction="constant" if constant else "input",
        dtype=dtype,
        domain=domain,
    )
    if domain is not None and "input" in entry:
        entry["domain"] = entry["input"].pop("domain")
    if relations is not None:
        entry["relations"] = relations
    return entry


def _write_shards(directory: Path, **series_by_shard: list[dict[str, object]]) -> None:
    bootstrap_binding_shards(directory, schema_version=CURRENT_SCHEMA_VERSION, workbook="dyn.xlsx")
    for shard, filename in _SHARDS.items():
        series = series_by_shard.get(shard, [])
        document = {"schema_version": CURRENT_SCHEMA_VERSION, "series": series}
        (directory / filename).write_text(
            yaml.safe_dump(document, sort_keys=False), encoding="utf-8"
        )


def _candidates(workbook: Path, bindings_dir: Path) -> list[str]:
    bindings = load_series_bindings(bindings_dir)
    config = DynamicRefConfig.from_bindings(bindings, workbook, bindings_path=bindings_dir)
    targets = all_series_targets(bindings, workbook=workbook)
    return list_dynamic_ref_constraint_candidates(workbook, targets, dynamic_refs=config)


def _output_series() -> dict[str, object]:
    return series_entry(
        "result",
        "Outputs!A1",
        layout="scalar",
        direction="output",
        compute_name="compute_result",
    )


def test_placeholder_shards_have_no_extraction_root(tmp_path: Path) -> None:
    workbook = _write_dynamic_workbook(tmp_path / "dyn.xlsx")
    bindings_dir = tmp_path / "bindings"
    _write_shards(bindings_dir)
    bindings = load_series_bindings(bindings_dir)
    assert all_series_targets(bindings, workbook=workbook) == []


def test_iterative_domains_extract_without_cached_refs_or_constraint_table(
    tmp_path: Path,
) -> None:
    workbook = _write_dynamic_workbook(tmp_path / "dyn.xlsx")
    bindings_dir = tmp_path / "bindings"
    _write_shards(bindings_dir, outputs=[_output_series()])

    assert _candidates(workbook, bindings_dir) == ["Inputs!B1", "Inputs!G1"]
    bindings = load_series_bindings(bindings_dir)
    config = DynamicRefConfig.from_bindings(bindings, workbook, bindings_path=bindings_dir)
    targets = all_series_targets(bindings, workbook=workbook)
    with pytest.raises(DynamicRefError, match="Inputs!G1"):
        create_dependency_graph(
            workbook,
            targets,
            load_values=True,
            dynamic_refs=config,
            use_cached_dynamic_refs=False,
        )

    # A domain that excludes the cached row hides the INDIRECT leaf and still extracts.
    _write_shards(
        bindings_dir,
        outputs=[_output_series()],
        inputs=[
            _scalar("row_offset", "Inputs!B1", domain={"enum": [0]}),
            _scalar("index_row", "Inputs!G1", domain={"between": {"min": 1, "max": 2}}),
        ],
    )
    assert _candidates(workbook, bindings_dir) == []
    narrow = load_series_bindings(bindings_dir)
    narrow_config = DynamicRefConfig.from_bindings(narrow, workbook, bindings_path=bindings_dir)
    narrow_graph = create_dependency_graph(
        workbook,
        all_series_targets(narrow, workbook=workbook),
        load_values=True,
        dynamic_refs=narrow_config,
        use_cached_dynamic_refs=False,
    )
    assert "Lookups!A2" not in narrow_graph
    assert "Inputs!C1" not in narrow_graph

    _write_shards(
        bindings_dir,
        outputs=[_output_series()],
        inputs=[
            _scalar("row_offset", "Inputs!B1", domain={"between": {"min": 0, "max": 1}}),
            _scalar("index_row", "Inputs!G1", domain={"between": {"min": 1, "max": 2}}),
        ],
    )
    assert _candidates(workbook, bindings_dir) == ["Inputs!C1"]

    _write_shards(
        bindings_dir,
        outputs=[_output_series()],
        inputs=[
            _scalar("row_offset", "Inputs!B1", domain={"between": {"min": 0, "max": 1}}),
            _scalar("index_row", "Inputs!G1", domain={"between": {"min": 1, "max": 2}}),
            _scalar(
                "indirect_address",
                "Inputs!C1",
                dtype="string",
                domain={"enum": ["Lookups!Z1"]},
            ),
        ],
    )
    assert _candidates(workbook, bindings_dir) == []
    dynamic_only = load_series_bindings(bindings_dir)
    dynamic_config = DynamicRefConfig.from_bindings(
        dynamic_only, workbook, bindings_path=bindings_dir
    )
    graph = create_dependency_graph(
        workbook,
        all_series_targets(dynamic_only, workbook=workbook),
        load_values=True,
        dynamic_refs=dynamic_config,
        use_cached_dynamic_refs=False,
    )
    assert "Lookups!A2" in graph
    assert "Inputs!C1" in graph
    assert "BASE" not in _candidates(workbook, bindings_dir)
    undomained = undomained_leaves(graph, dynamic_only, workbook=workbook)
    assert undomained == [
        "Inputs!D1",
        "Inputs!E1",
        "Inputs!F1",
        "Lookups!A1",
        "Lookups!B1",
        "Lookups!Z1",
        "Lookups!B2",
    ]

    with pytest.raises(SeriesRelationError, match="grace"):
        _write_shards(
            bindings_dir,
            outputs=[_output_series()],
            inputs=[
                _scalar("row_offset", "Inputs!B1", domain={"between": {"min": 0, "max": 1}}),
                _scalar("index_row", "Inputs!G1", domain={"between": {"min": 1, "max": 2}}),
                _scalar(
                    "indirect_address",
                    "Inputs!C1",
                    dtype="string",
                    domain={"enum": ["Lookups!Z1"]},
                ),
                _scalar(
                    "maturity",
                    "Inputs!E1",
                    domain={"between": {"min": 0, "max": 80}},
                    relations=[{"greater_than": "grace"}],
                ),
            ],
        )
        DynamicRefConfig.from_bindings(
            load_series_bindings(bindings_dir), workbook, bindings_path=bindings_dir
        )

    _write_shards(
        bindings_dir,
        outputs=[_output_series()],
        inputs=[
            _scalar("row_offset", "Inputs!B1", domain={"between": {"min": 0, "max": 1}}),
            _scalar("index_row", "Inputs!G1", domain={"between": {"min": 1, "max": 2}}),
            _scalar(
                "indirect_address",
                "Inputs!C1",
                dtype="string",
                domain={"enum": ["Lookups!Z1"]},
            ),
            _scalar("grace", "Inputs!D1", domain={"between": {"min": 0, "max": 50}}),
            _scalar(
                "maturity",
                "Inputs!E1",
                domain={"between": {"min": 0, "max": 80}},
                relations=[{"greater_than": "grace"}],
            ),
            _scalar(
                "plain_scale",
                "Inputs!F1",
                dtype="float",
                domain={"real_between": {"min": 0, "max": 100}},
            ),
        ],
        constants=[
            _scalar("lookup_base", "Lookups!A1", dtype="float", constant=True),
            _scalar("lookup_b1", "Lookups!B1", dtype="float", constant=True),
            _scalar("lookup_b2", "Lookups!B2", dtype="float", constant=True),
            _scalar("lookup_z", "Lookups!Z1", dtype="float", constant=True),
        ],
    )
    enriched = load_series_bindings(bindings_dir)
    node_count = len(list(graph.keys()))
    graph.attach_domains(enriched, workbook=workbook, bindings_path=bindings_dir)
    assert len(list(graph.keys())) == node_count
    assert "Lookups!A2" in graph
    assert undomained_leaves(graph, enriched, workbook=workbook) == []
    maturity = graph.domain_for("Inputs!E1")
    assert maturity is not None
    assert maturity.relations == (GreaterThanCell("Inputs!D1"),)
    assert (
        list_dynamic_ref_constraint_candidates(
            workbook,
            all_series_targets(enriched, workbook=workbook),
            dynamic_refs=DynamicRefConfig.from_bindings(
                enriched, workbook, bindings_path=bindings_dir
            ),
        )
        == []
    )


def test_blank_dynamic_ref_leaf_is_not_pinned_by_from_workbook(tmp_path: Path) -> None:
    workbook = write_workbook(
        tmp_path / "blank.xlsx",
        {
            "Inputs": {"H1": None},
            "Lookups": {"A1": 100},
            "Engine": {"A1": "=OFFSET(Lookups!A1,0,Inputs!H1)"},
            "Outputs": {"A1": "=Engine!A1"},
        },
    )
    bindings_dir = tmp_path / "bindings"
    _write_shards(
        bindings_dir,
        outputs=[_output_series()],
        constants=[_scalar("blank_col", "Inputs!H1", constant=True)],
    )
    assert _candidates(workbook, bindings_dir) == ["Inputs!H1"]

    _write_shards(
        bindings_dir,
        outputs=[_output_series()],
        inputs=[_scalar("blank_col", "Inputs!H1", domain={"enum": [0]})],
    )
    assert _candidates(workbook, bindings_dir) == []
    bindings = load_series_bindings(bindings_dir)
    graph = create_dependency_graph(
        workbook,
        all_series_targets(bindings, workbook=workbook),
        load_values=True,
        dynamic_refs=DynamicRefConfig.from_bindings(bindings, workbook, bindings_path=bindings_dir),
        use_cached_dynamic_refs=False,
    )
    assert "Lookups!A1" in graph
    assert "Inputs!H1" in graph


def test_tiny_dsa_output_roots_extract_after_one_domain() -> None:
    """Pipeline canary: outputs only, then the one dynamic-ref leaf, no cache."""
    workbook = INVERTED_TREE_TINY_DSA / "tiny-dsa.xlsx"
    bindings_dir = INVERTED_TREE_TINY_DSA / "bindings"
    full = load_series_bindings(bindings_dir)
    outputs_only = copy.deepcopy(full)
    outputs_only["series"] = [series for series in full["series"] if has_output_direction(series)]
    output_config = DynamicRefConfig.from_bindings(outputs_only, workbook)
    output_targets = all_series_targets(outputs_only, workbook=workbook)
    assert list_dynamic_ref_constraint_candidates(
        workbook, output_targets, dynamic_refs=output_config
    ) == ["Inputs!B22"]
    with pytest.raises(DynamicRefError, match="Inputs!B22"):
        create_dependency_graph(
            workbook,
            output_targets,
            load_values=True,
            dynamic_refs=output_config,
            use_cached_dynamic_refs=False,
        )

    shock_type = next(series for series in full["series"] if series["id"] == "shock_type")
    with_domain = copy.deepcopy(outputs_only)
    with_domain["series"] = [*outputs_only["series"], copy.deepcopy(shock_type)]
    config = DynamicRefConfig.from_bindings(with_domain, workbook)
    targets = all_series_targets(with_domain, workbook=workbook)
    assert list_dynamic_ref_constraint_candidates(workbook, targets, dynamic_refs=config) == []
    graph = create_dependency_graph(
        workbook,
        targets,
        load_values=True,
        dynamic_refs=config,
        use_cached_dynamic_refs=False,
    )
    assert "Inputs!B22" in graph
    undomained = undomained_leaves(graph, with_domain, workbook=workbook)
    assert "Inputs!B22" not in undomained
    assert "Inputs!B21" in undomained

    graph.attach_domains(full, workbook=workbook, bindings_path=bindings_dir)
    assert undomained_leaves(graph, full, workbook=workbook) == []
