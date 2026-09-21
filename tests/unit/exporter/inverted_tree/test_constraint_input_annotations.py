"""Issue 944 — public inputs take their type from `graph.cell_type_env`."""

from __future__ import annotations

import subprocess
from pathlib import Path
from typing import Annotated, Literal, get_args, get_origin, get_type_hints

import pytest

from excel_grapher.core.cell_types import Between, RealBetween
from excel_grapher.exporter.inverted_tree.errors import InvertedTreeExportError
from excel_grapher.grapher import DynamicRefConfig
from excel_grapher.series_bindings.schema import validate_bindings_document
from tests.unit.exporter.inverted_tree.helpers import (
    bindings_document,
    generate_inverted,
    load_package,
    series_entry,
    write_workbook,
)

_REPO_ROOT = Path(__file__).resolve().parents[4]


def _flag_workbook(tmp_path: Path) -> Path:
    return write_workbook(
        tmp_path / "flag.xlsx",
        {
            "Inputs": {"A1": "On"},
            "Outputs": {"B1": '=IF(Inputs!A1="On", 1, 0)'},
        },
    )


def _flag_bindings(*, domain: dict | None = None) -> dict:
    return bindings_document(
        series_entry(
            "flag",
            "Inputs!A1",
            layout="scalar",
            direction="input",
            dtype="string",
            domain=domain,
        ),
        series_entry(
            "result",
            "Outputs!B1",
            layout="scalar",
            direction="output",
            dtype="int",
            compute_name="compute_result",
        ),
        schema_version="1.19.0",
    )


def _enum_config() -> DynamicRefConfig:
    return DynamicRefConfig.from_constraints({"Inputs!A1": Literal["On", "Off"]})


def _hints(cls: type) -> dict[str, object]:
    return get_type_hints(cls, include_extras=True)


def test_constraint_enum_is_literal_on_model_and_inputs(tmp_path: Path) -> None:
    modules = generate_inverted(
        _flag_workbook(tmp_path),
        _flag_bindings(),
        dynamic_refs=_enum_config(),
    )
    assert "require_input_domain" not in modules["validation.py"]
    assert "require_annotated_domain" in modules["validation.py"]
    assert 'Literal["Off", "On"]' in modules["model.py"]
    assert 'Literal["Off", "On"]' in modules["validation.py"]

    pkg = load_package(modules, tmp_path, name="constraint_enum")
    model_flag = _hints(pkg.Model)["flag"]
    inputs_flag = _hints(pkg.ResultInputs)["flag"]
    assert model_flag == inputs_flag
    assert get_origin(model_flag) is Literal
    assert set(get_args(model_flag)) == {"On", "Off"}
    assert pkg.compute_result(pkg.ResultInputs(flag="On")) == 1
    assert pkg.compute_result(pkg.ResultInputs(flag="Off")) == 0
    with pytest.raises(ValueError, match=r"flag out of domain"):
        pkg.Model(flag="Maybe")
    with pytest.raises(ValueError, match=r"flag out of domain"):
        pkg.ResultInputs(flag="Maybe")


def test_constraint_enum_annotation_passes_ty(tmp_path: Path) -> None:
    modules = generate_inverted(
        _flag_workbook(tmp_path),
        _flag_bindings(),
        dynamic_refs=_enum_config(),
    )
    assert 'Literal["Off", "On"]' in modules["model.py"]
    package = tmp_path / "constraint_ty"
    package.mkdir()
    for filename, content in modules.items():
        (package / filename).write_text(content, encoding="utf-8")
    ty = subprocess.run(
        [
            "uv",
            "run",
            "--no-sync",
            "ty",
            "check",
            "--extra-search-path",
            str(package.parent),
            "--project",
            str(_REPO_ROOT),
            "--ignore",
            "unresolved-attribute",
            str(package / "model.py"),
        ],
        cwd=str(_REPO_ROOT),
        capture_output=True,
        text=True,
        check=False,
    )
    assert ty.returncode == 0, f"ty failed:\n{ty.stdout}\n{ty.stderr}"


def test_binding_domain_on_the_graph_is_not_a_second_runtime_dict(tmp_path: Path) -> None:
    workbook = _flag_workbook(tmp_path)
    document = _flag_bindings(domain={"enum": ["On", "Off"]})
    bindings = validate_bindings_document(document)
    modules = generate_inverted(
        workbook,
        document,
        dynamic_refs=DynamicRefConfig.from_bindings(bindings, workbook),
    )
    assert "require_input_domain" not in modules["validation.py"]
    assert get_origin(_hints(load_package(modules, tmp_path, name="bound_enum").Model)["flag"]) is (
        Literal
    )


def test_unconstrained_input_keeps_measure_dtype(tmp_path: Path) -> None:
    modules = generate_inverted(_flag_workbook(tmp_path), _flag_bindings())
    assert "Literal[" not in modules["model.py"]
    assert "require_annotated_domain" not in modules["validation.py"]
    pkg = load_package(modules, tmp_path, name="unconstrained_flag")
    assert _hints(pkg.Model)["flag"] is str


def test_between_and_real_between_annotations(tmp_path: Path) -> None:
    workbook = write_workbook(
        tmp_path / "bounds.xlsx",
        {
            "Inputs": {"A1": 2000, "B1": 0.25},
            "Outputs": {"A1": "=Inputs!A1", "B1": "=Inputs!B1"},
        },
    )
    document = bindings_document(
        series_entry("year", "Inputs!A1", layout="scalar", direction="input", dtype="int"),
        series_entry("share", "Inputs!B1", layout="scalar", direction="input", dtype="float"),
        series_entry("year_out", "Outputs!A1", layout="scalar", direction="output", dtype="int"),
        series_entry("share_out", "Outputs!B1", layout="scalar", direction="output", dtype="float"),
        schema_version="1.19.0",
    )
    modules = generate_inverted(
        workbook,
        document,
        dynamic_refs=DynamicRefConfig.from_constraints(
            {
                "Inputs!A1": Annotated[int, Between(1990, 2100)],
                "Inputs!B1": Annotated[float, RealBetween(0, 1)],
            }
        ),
    )
    pkg = load_package(modules, tmp_path, name="constraint_bounds")
    year = _hints(pkg.Model)["year"]
    share = _hints(pkg.Model)["share"]
    assert get_origin(year) is Annotated
    assert get_args(year)[0] is int
    year_meta = get_args(year)[1]
    assert type(year_meta).__name__ == "Between"
    assert (year_meta.min, year_meta.max) == (1990, 2100)
    assert get_args(share)[0] is float
    share_meta = get_args(share)[1]
    assert type(share_meta).__name__ == "RealBetween"
    assert (share_meta.min, share_meta.max) == (0.0, 1.0)
    assert pkg.compute_year_out(pkg.YearOutInputs(year=2000)) == 2000
    with pytest.raises(ValueError, match=r"year out of domain"):
        pkg.YearOutInputs(year=1989)
    with pytest.raises(ValueError, match=r"share out of domain"):
        pkg.ShareOutInputs(share=1.1)
    assert pkg.compute_share_out(pkg.ShareOutInputs(share=0)) == 0.0


def test_incompatible_domains_on_one_series_fail_closed(tmp_path: Path) -> None:
    workbook = write_workbook(
        tmp_path / "split.xlsx",
        {
            "Inputs": {"B1": "On", "C1": "Yes", "B10": 1, "C10": 2},
            "Outputs": {"A1": "=Inputs!B1", "B1": "=Inputs!C1", "A10": 1, "B10": 2},
        },
    )
    document = bindings_document(
        series_entry(
            "choice",
            "Inputs!B1:C1",
            layout="series",
            direction="input",
            dtype="string",
            header_row=10,
        ),
        series_entry(
            "out",
            "Outputs!A1:B1",
            layout="series",
            direction="output",
            dtype="string",
            header_row=10,
        ),
        schema_version="1.19.0",
    )
    with pytest.raises(InvertedTreeExportError, match="incompatible constraint domains"):
        generate_inverted(
            workbook,
            document,
            dynamic_refs=DynamicRefConfig.from_constraints(
                {
                    "Inputs!B1": Literal["On", "Off"],
                    "Inputs!C1": Literal["Yes", "No"],
                }
            ),
        )


def test_partial_constraint_on_a_series_fails_closed(tmp_path: Path) -> None:
    workbook = write_workbook(
        tmp_path / "partial.xlsx",
        {
            "Inputs": {"B1": "On", "C1": "On", "B10": 1, "C10": 2},
            "Outputs": {"A1": "=Inputs!B1", "B1": "=Inputs!C1", "A10": 1, "B10": 2},
        },
    )
    document = bindings_document(
        series_entry(
            "choice",
            "Inputs!B1:C1",
            layout="series",
            direction="input",
            dtype="string",
            header_row=10,
        ),
        series_entry(
            "out",
            "Outputs!A1:B1",
            layout="series",
            direction="output",
            dtype="string",
            header_row=10,
        ),
        schema_version="1.19.0",
    )
    with pytest.raises(InvertedTreeExportError, match="incompatible constraint domains"):
        generate_inverted(
            workbook,
            document,
            dynamic_refs=DynamicRefConfig.from_constraints({"Inputs!B1": Literal["On", "Off"]}),
        )


def test_extract_only_pins_stay_off_the_public_type(tmp_path: Path) -> None:
    workbook = write_workbook(
        tmp_path / "pins.xlsx",
        {
            "Inputs": {"A1": "On", "B1": "frozen"},
            "Outputs": {"A1": '=IF(Inputs!A1="On", 1, 0)'},
        },
    )
    document = bindings_document(
        series_entry("flag", "Inputs!A1", layout="scalar", direction="input", dtype="string"),
        series_entry("pinned", "Inputs!B1", layout="scalar", direction="constant", dtype="string"),
        series_entry("result", "Outputs!A1", layout="scalar", direction="output", dtype="int"),
        schema_version="1.19.0",
    )
    modules = generate_inverted(
        workbook,
        document,
        dynamic_refs=DynamicRefConfig.from_constraints(
            {
                "Inputs!A1": Literal["On"],
                "Inputs!B1": Literal["frozen", "other"],
            }
        ),
    )
    pkg = load_package(modules, tmp_path, name="extract_pins")
    assert "pinned" not in _hints(pkg.Model)
    assert "pinned" not in _hints(pkg.ResultInputs)
    assert _hints(pkg.Model)["flag"] is str
    assert "Literal[" not in modules["model.py"]

    blank = generate_inverted(
        _flag_workbook(tmp_path),
        _flag_bindings(),
        dynamic_refs=DynamicRefConfig.from_constraints({"Inputs!A1": Literal[None]}),
    )
    assert "Literal[" not in blank["model.py"]


def test_value_map_needles_are_not_the_public_literal(tmp_path: Path) -> None:
    workbook = write_workbook(
        tmp_path / "selector.xlsx",
        {
            "Inputs": {"A1": "High "},
            "Outputs": {"A1": '=IF(Inputs!A1="High ", 1, 0)'},
        },
    )
    document = bindings_document(
        series_entry(
            "selector",
            "Inputs!A1",
            layout="scalar",
            direction="input",
            dtype="string",
            value_map={"High": "High ", "Low": "Low "},
        ),
        series_entry("out", "Outputs!A1", layout="scalar", direction="output", dtype="int"),
        schema_version="1.19.0",
    )
    bindings = validate_bindings_document(document)
    modules = generate_inverted(
        workbook,
        document,
        dynamic_refs=DynamicRefConfig.from_bindings(bindings, workbook),
    )
    assert "Literal[" not in modules["model.py"]
    assert "require_input_domain(selector" in modules["validation.py"]
    pkg = load_package(modules, tmp_path, name="needle_map")
    assert pkg.compute_out(pkg.OutInputs(selector="High")) == 1
    with pytest.raises(ValueError, match=r"selector out of domain"):
        pkg.OutInputs(selector="Nope")


def test_between_on_float_input_fails_closed(tmp_path: Path) -> None:
    workbook = write_workbook(
        tmp_path / "share.xlsx",
        {"Inputs": {"A1": 0.0}, "Outputs": {"B1": "=Inputs!A1"}},
    )
    document = bindings_document(
        series_entry("share", "Inputs!A1", layout="scalar", direction="input", dtype="float"),
        series_entry("result", "Outputs!B1", layout="scalar", direction="output"),
        schema_version="1.19.0",
    )
    with pytest.raises(InvertedTreeExportError, match="Between requires integer"):
        generate_inverted(
            workbook,
            document,
            dynamic_refs=DynamicRefConfig.from_constraints(
                {"Inputs!A1": Annotated[int, Between(0, 1)]}
            ),
        )


def test_bool_literal_rejects_int_one_and_zero(tmp_path: Path) -> None:
    workbook = write_workbook(
        tmp_path / "bool.xlsx",
        {"Inputs": {"A1": True}, "Outputs": {"A1": "=Inputs!A1"}},
    )
    document = bindings_document(
        series_entry("flag", "Inputs!A1", layout="scalar", direction="input", dtype="bool"),
        series_entry("out", "Outputs!A1", layout="scalar", direction="output", dtype="bool"),
        schema_version="1.19.0",
    )
    modules = generate_inverted(
        workbook,
        document,
        dynamic_refs=DynamicRefConfig.from_constraints({"Inputs!A1": Literal[True, False]}),
    )
    pkg = load_package(modules, tmp_path, name="bool_literal")
    assert get_origin(_hints(pkg.Model)["flag"]) is Literal
    assert pkg.Model(flag=True).flag is True
    with pytest.raises(ValueError, match=r"flag out of domain"):
        pkg.Model(flag=1)
    with pytest.raises(ValueError, match=r"flag out of domain"):
        pkg.Model(flag=0)


def test_shared_series_enum_is_literal_element_type(tmp_path: Path) -> None:
    workbook = write_workbook(
        tmp_path / "choices.xlsx",
        {
            "Inputs": {"B1": "On", "C1": "Off", "B10": 1, "C10": 2},
            "Outputs": {"A1": "=Inputs!B1", "B1": "=Inputs!C1", "A10": 1, "B10": 2},
        },
    )
    document = bindings_document(
        series_entry(
            "choice",
            "Inputs!B1:C1",
            layout="series",
            direction="input",
            dtype="string",
            header_row=10,
        ),
        series_entry(
            "out",
            "Outputs!A1:B1",
            layout="series",
            direction="output",
            dtype="string",
            header_row=10,
        ),
        schema_version="1.19.0",
    )
    modules = generate_inverted(
        workbook,
        document,
        dynamic_refs=DynamicRefConfig.from_constraints(
            {
                "Inputs!B1": Literal["On", "Off"],
                "Inputs!C1": Literal["On", "Off"],
            }
        ),
    )
    assert 'Series[Literal["Off", "On"] | None]' in modules["data.py"]
    assert "require_annotated_domain" in modules["validation.py"]
    pkg = load_package(modules, tmp_path, name="series_enum")
    good = pkg.data.CHOICE.with_nested(("On", "Off"))
    result = pkg.compute_out(pkg.OutInputs(choice=good))
    assert (result[1], result[2]) == ("On", "Off")
    with pytest.raises(ValueError, match=r"choice.*out of domain"):
        pkg.OutInputs(choice=pkg.data.CHOICE.with_nested(("On", "Maybe")))
