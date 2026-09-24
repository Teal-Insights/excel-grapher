"""Issue 666 — inverted-tree `compute_*` arguments honor `input.domain`."""

from __future__ import annotations

from importlib import import_module
from pathlib import Path

import pytest

from excel_grapher.exporter.inverted_tree.errors import InvertedTreeExportError
from excel_grapher.grapher import create_dependency_graph
from excel_grapher.series_bindings import validate_series_bindings
from excel_grapher.series_bindings.schema import validate_bindings_document
from excel_grapher.series_bindings.workflow import all_series_targets
from tests.unit.exporter.inverted_tree.helpers import (
    bindings_document,
    generate_inverted,
    load_package,
    series_entry,
    write_workbook,
)
from tests.unit.exporter.inverted_tree.test_shape_a1_leaf_closure import (
    _a1_bindings,
    _a1_workbook,
)


def _enum_flag_workbook(tmp_path: Path) -> Path:
    return write_workbook(
        tmp_path / "flag.xlsx",
        {
            "Inputs": {"A1": 0},
            "Outputs": {"A1": "=Inputs!A1"},
        },
    )


def _enum_flag_bindings() -> dict:
    return bindings_document(
        series_entry(
            "flag",
            "Inputs!A1",
            layout="scalar",
            direction="input",
            dtype="int",
            domain={"enum": [0, 1]},
        ),
        series_entry("out", "Outputs!A1", layout="scalar", direction="output", dtype="int"),
    )


def _rate_workbook(tmp_path: Path) -> Path:
    return write_workbook(
        tmp_path / "rate.xlsx",
        {
            "Inputs": {"B1": 0.25, "C1": 0.5, "B10": 1, "C10": 2},
            "Outputs": {"A1": "=Inputs!B1", "B1": "=Inputs!C1", "A10": 1, "B10": 2},
        },
    )


def _rate_bindings() -> dict:
    return bindings_document(
        series_entry(
            "rate",
            "Inputs!B1:C1",
            layout="series",
            direction="input",
            header_row=10,
            domain={"real_between": {"min": 0, "max": 1}},
        ),
        series_entry(
            "out",
            "Outputs!A1:B1",
            layout="series",
            direction="output",
            header_row=10,
        ),
    )


def test_emit_refuses_float_dtype_with_between_domain(tmp_path: Path) -> None:
    workbook = write_workbook(
        tmp_path / "share.xlsx",
        {
            "Inputs": {"A1": 0.0},
            "Outputs": {"B1": "=Inputs!A1"},
        },
    )
    document = bindings_document(
        series_entry(
            "share",
            "Inputs!A1",
            layout="scalar",
            direction="input",
            dtype="float",
            domain={"between": {"min": 0, "max": 1}},
        ),
        series_entry("result", "Outputs!B1", layout="scalar", direction="output"),
    )
    bindings = validate_bindings_document(document)
    graph = create_dependency_graph(
        workbook,
        all_series_targets(bindings, workbook=workbook),
        load_values=True,
    )
    report = validate_series_bindings(graph, bindings, workbook=workbook)
    assert report["ok"] is False
    assert any(issue["code"] == "domain_dtype_mismatch" for issue in report["issues"])
    with pytest.raises(InvertedTreeExportError, match="domain_dtype_mismatch"):
        generate_inverted(workbook, document)


def test_scalar_enum_domain_accepts_and_rejects(tmp_path: Path) -> None:
    pkg = load_package(
        generate_inverted(_enum_flag_workbook(tmp_path), _enum_flag_bindings()),
        tmp_path,
        name="domain_enum",
    )
    assert pkg.compute_out(pkg.OutInputs(flag=0)) == 0
    assert pkg.compute_out(pkg.OutInputs(flag=1)) == 1
    with pytest.raises(ValueError, match=r"flag out of domain"):
        pkg.OutInputs(flag=2)


def test_series_real_between_names_series_on_out_of_range_member(tmp_path: Path) -> None:
    pkg = load_package(
        generate_inverted(_rate_workbook(tmp_path), _rate_bindings()),
        tmp_path,
        name="domain_rate",
    )
    rate = pkg.data.RATE.with_nested((0.0, 1.0))
    result = pkg.compute_out(pkg.OutInputs(rate=rate))
    assert (result[1], result[2]) == pytest.approx((0.0, 1.0))
    with pytest.raises(ValueError, match=r"rate\(2,\) out of domain"):
        pkg.OutInputs(rate=pkg.data.RATE.with_nested((0.0, 1.1)))


def test_no_input_domain_does_not_emit_domain_guard(tmp_path: Path) -> None:
    modules = generate_inverted(_a1_workbook(tmp_path), _a1_bindings())
    assert "require_input_domain" not in modules["api.py"]
    assert "require_input_domain" not in modules["internals.py"]
    assert "require_input_domain" not in modules["validation.py"]


def test_shared_runner_checks_domain_before_evaluation(tmp_path: Path) -> None:
    workbook = write_workbook(
        tmp_path / "shared.xlsx",
        {
            "Inputs": {"A1": 0},
            "Engine": {"A1": "=Inputs!A1", "B1": "=Inputs!A1+1"},
            "Outputs": {"A1": "=Engine!A1", "B1": "=Engine!B1"},
        },
    )
    document = bindings_document(
        series_entry(
            "flag",
            "Inputs!A1",
            layout="scalar",
            direction="input",
            dtype="int",
            domain={"enum": [0, 1]},
        ),
        series_entry("engine_a", "Engine!A1", layout="scalar", direction="internal", dtype="int"),
        series_entry("engine_b", "Engine!B1", layout="scalar", direction="internal", dtype="int"),
        series_entry("out_a", "Outputs!A1", layout="scalar", direction="output", dtype="int"),
        series_entry("out_b", "Outputs!B1", layout="scalar", direction="output", dtype="int"),
    )
    modules = generate_inverted(workbook, document)
    validation = modules["validation.py"]
    assert validation.count("require_input_domain(flag") == 1
    assert "require_input_domain" not in modules["api.py"]
    assert "require_input_domain" not in modules["internals.py"]
    pkg = load_package(modules, tmp_path, name="domain_shared")
    assert pkg.compute_out_a(pkg.OutAInputs(flag=0)) == 0
    with pytest.raises(ValueError, match=r"flag out of domain"):
        pkg.OutBInputs(flag=2)


def test_compute_float_real_between_coerces_int_like_setters(tmp_path: Path) -> None:
    workbook = write_workbook(
        tmp_path / "share.xlsx",
        {
            "Inputs": {"A1": 0.0},
            "Outputs": {"B1": "=Inputs!A1"},
        },
    )
    document = bindings_document(
        series_entry(
            "share",
            "Inputs!A1",
            layout="scalar",
            direction="input",
            dtype="float",
            domain={"real_between": {"min": 0, "max": 1}},
        ),
        series_entry("result", "Outputs!B1", layout="scalar", direction="output"),
    )
    modules = generate_inverted(workbook, document)
    assert "coerce_input_measure(share" in modules["validation.py"]
    assert modules["validation.py"].index("coerce_input_measure(share") < modules[
        "validation.py"
    ].index("require_input_domain(share")
    assert "coerce_input_measure" not in modules["api.py"]
    assert "coerce_input_measure" not in modules["internals.py"]
    pkg = load_package(modules, tmp_path, name="share_coerce")
    zero = pkg.compute_result(pkg.ResultInputs(share=0))
    assert zero == 0.0
    assert type(zero) is float
    assert pkg.compute_result(pkg.ResultInputs(share=0.0)) == 0.0
    assert pkg.compute_result(pkg.ResultInputs(share=pkg.data.SHARE_DEFAULT)) == 0.0
    with pytest.raises(ValueError, match=r"share out of domain"):
        pkg.ResultInputs(share=1.1)


def test_union_domain_accepts_workbook_sentinel_and_numbers(tmp_path: Path) -> None:
    """A text sentinel and a real interval are one caller-facing input domain."""
    bounds = {"min": -1.0e15, "max": 1.0e15}
    workbook = write_workbook(
        tmp_path / "sentinel.xlsx",
        {
            "Inputs": {"A1": "n.a."},
            "Outputs": {"B1": "=IF(ISNUMBER(Inputs!A1), Inputs!A1, 0)"},
        },
    )
    document = bindings_document(
        series_entry(
            "sentinel",
            "Inputs!A1",
            layout="scalar",
            direction="input",
            dtype="float",
            domain={"enum": ["n.a."], "real_between": bounds},
        ),
        series_entry(
            "result",
            "Outputs!B1",
            layout="scalar",
            direction="output",
            dtype="float",
            compute_name="compute_result",
        ),
        schema_version="1.22.0",
    )
    modules = generate_inverted(workbook, document)
    validation = modules["validation.py"]
    assert "require_input_domain(sentinel" in validation
    assert '"n.a."' in validation or "'n.a.'" in validation
    assert "real_between" in validation
    pkg = load_package(modules, tmp_path, name="sentinel_union")
    assert pkg.compute_result(pkg.ResultInputs.from_defaults()) == 0
    assert pkg.compute_result(pkg.ResultInputs(sentinel=3.5)) == 3.5
    with pytest.raises(ValueError, match=r"sentinel out of domain"):
        pkg.ResultInputs(sentinel="nope")
    with pytest.raises(ValueError, match=r"sentinel out of domain"):
        pkg.ResultInputs(sentinel=1.0e16)


def test_union_domain_annotation_when_bindings_supply_cell_types(tmp_path: Path) -> None:
    from excel_grapher.grapher.dynamic_refs import DynamicRefConfig

    bounds = {"min": -1.0, "max": 1.0}
    workbook = write_workbook(
        tmp_path / "sentinel.xlsx",
        {
            "Inputs": {"A1": "n.a."},
            "Outputs": {"B1": "=IF(ISNUMBER(Inputs!A1), Inputs!A1, 0)"},
        },
    )
    document = bindings_document(
        series_entry(
            "sentinel",
            "Inputs!A1",
            layout="scalar",
            direction="input",
            dtype="float",
            domain={"enum": ["n.a."], "real_between": bounds},
        ),
        series_entry(
            "result",
            "Outputs!B1",
            layout="scalar",
            direction="output",
            dtype="float",
            compute_name="compute_result",
        ),
        schema_version="1.22.0",
    )
    bindings = validate_bindings_document(document)
    modules = generate_inverted(
        workbook,
        document,
        dynamic_refs=DynamicRefConfig.from_bindings(bindings, workbook),
    )
    assert 'Literal["n.a."]' in modules["model.py"]
    assert "Annotated[float, RealBetween(" in modules["model.py"]
    assert " | " in modules["model.py"]
    pkg = load_package(modules, tmp_path, name="sentinel_annotated")
    assert pkg.compute_result(pkg.ResultInputs.from_defaults()) == 0
    assert pkg.compute_result(pkg.ResultInputs(sentinel=-0.25)) == -0.25
    with pytest.raises(ValueError, match=r"sentinel out of domain"):
        pkg.ResultInputs(sentinel="nope")


def test_compute_series_float_coerces_int_members(tmp_path: Path) -> None:
    workbook = _rate_workbook(tmp_path)
    modules = generate_inverted(workbook, _rate_bindings())
    assert "coerce_input_measure(rate" in modules["validation.py"]
    pkg = load_package(modules, tmp_path, name="rate_coerce")
    rate = pkg.data.RATE.with_nested((0, 1))
    assert type(rate[1]) is int
    checked = import_module(f"{pkg.__name__}.validation").CHECKS["rate"](rate)
    assert type(checked[1]) is float
    assert type(checked[2]) is float
    result = pkg.compute_out(pkg.OutInputs(rate=rate))
    assert (result[1], result[2]) == pytest.approx((0.0, 1.0))
