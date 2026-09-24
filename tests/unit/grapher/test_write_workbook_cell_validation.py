"""Excel data validation on write-back for constrained inputs."""

from __future__ import annotations

from pathlib import Path
from typing import Annotated, Any, Literal

import fastpyxl
import pytest
from fastpyxl.worksheet.datavalidation import expand_cell_ranges

from excel_grapher.core.cell_types import (
    Between,
    CellKind,
    CellType,
    EnumDomain,
    GreaterThanCell,
    NotEqualCell,
    RealBetween,
    RealIntervalDomain,
    constraints_to_cell_type_env,
)
from excel_grapher.grapher import create_dependency_graph, write_workbook
from excel_grapher.grapher.graph import DependencyGraph
from excel_grapher.grapher.node import make_cell_node
from excel_grapher.series_bindings import validate_bindings_document
from excel_grapher.series_bindings.types import WorkbookSeriesBindings


def _validation_rows(path: Path, sheet: str = "Inputs") -> list[dict[str, Any]]:
    wb = fastpyxl.load_workbook(path)
    try:
        rows: list[dict[str, Any]] = []
        for rule in wb[sheet].data_validations.dataValidation:
            sqref = str(rule.sqref)
            rows.append(
                {
                    "type": rule.type,
                    "operator": rule.operator,
                    "formula1": rule.formula1,
                    "formula2": rule.formula2,
                    "allow_blank": bool(rule.allowBlank),
                    "show_error": bool(rule.showErrorMessage),
                    "error_style": rule.errorStyle,
                    "sqref": sqref,
                }
            )
        return rows
    finally:
        wb.close()


def _cells(row: dict[str, Any]) -> set[str]:
    return expand_cell_ranges(row["sqref"]) if row["sqref"] else set()


def _covers(rows: list[dict[str, Any]], coord: str) -> list[dict[str, Any]]:
    return [row for row in rows if coord in _cells(row)]


def _write_flag_workbook(path: Path) -> None:
    wb = fastpyxl.Workbook()
    ws = wb.active
    assert ws is not None
    ws.title = "Inputs"
    ws["A1"] = "On"
    ws["B1"] = '=IF(A1="On",1,0)'
    ws["C1"] = 3
    ws["A2"] = "Off"
    wb.save(path)
    wb.close()


def _measure(dtype: str, read: str) -> dict[str, Any]:
    return {
        "measure": {
            "concept": "OBS_VALUE",
            "dtype": dtype,
            "bind": {"kind": "data_cell", "read": read},
        },
        "dimensions": [],
    }


def _series(
    *,
    series_id: str,
    data_range: str,
    direction: str = "input",
    dtype: str = "string",
    read: str | None = None,
    domain: dict[str, Any] | None = None,
    relations: list[dict[str, str]] | None = None,
    value_map: dict[str, str] | None = None,
    layout: str = "scalar",
) -> dict[str, Any]:
    block: dict[str, Any] = {}
    if value_map is not None:
        block["value_map"] = value_map
    series: dict[str, Any] = {
        "id": series_id,
        "sheet": data_range.split("!", 1)[0].strip("'"),
        "data_range": data_range,
        "layout": layout,
        direction: block,
        "structure": _measure(dtype, read or dtype),
        "key": [],
    }
    if domain is not None:
        series["domain"] = domain
    if relations is not None:
        series["relations"] = relations
    return series


def _bindings(*series: dict[str, Any]) -> WorkbookSeriesBindings:
    return validate_bindings_document({"schema_version": "1.21.0", "series": list(series)})


def _graph_from(path: Path, *targets: str) -> DependencyGraph:
    return create_dependency_graph(path, list(targets), load_values=True)


def test_input_enum_and_between_become_excel_validations(tmp_path: Path) -> None:
    """Dropdown and whole-number bounds are written onto input cells only."""
    source = tmp_path / "source.xlsx"
    _write_flag_workbook(source)
    graph = _graph_from(source, "Inputs!B1", "Inputs!C1")
    bindings = _bindings(
        _series(
            series_id="flag",
            data_range="Inputs!A1",
            domain={"enum": ["On", "Off"]},
        ),
        _series(
            series_id="year",
            data_range="Inputs!C1",
            dtype="int",
            read="int",
            domain={"between": {"min": 1, "max": 5}},
        ),
        _series(
            series_id="result",
            data_range="Inputs!B1",
            direction="output",
            dtype="float",
            read="float",
            domain={"enum": ["nope"]},
        )
        | {"output": {"compute": {"name": "compute_result"}}},
    )
    dest = tmp_path / "out.xlsx"
    write_workbook(
        graph,
        dest,
        series_bindings=bindings,
        bindings_workbook=source,
    )

    rows = _validation_rows(dest)
    flag = _covers(rows, "A1")
    assert len(flag) == 1
    assert flag[0]["type"] == "list"
    assert flag[0]["formula1"] == '"Off,On"'
    assert flag[0]["show_error"] is True
    assert flag[0]["error_style"] == "stop"
    assert flag[0]["allow_blank"] is False
    year = _covers(rows, "C1")
    assert len(year) == 1
    assert year[0]["type"] == "whole"
    assert year[0]["operator"] == "between"
    assert year[0]["formula1"] == "1"
    assert year[0]["formula2"] == "5"
    assert _covers(rows, "B1") == []


def test_same_enum_is_one_validation_rule(tmp_path: Path) -> None:
    source = tmp_path / "source.xlsx"
    wb = fastpyxl.Workbook()
    ws = wb.active
    assert ws is not None
    ws.title = "Inputs"
    ws["B2"] = "On"
    ws["B3"] = "Off"
    wb.save(source)
    wb.close()
    graph = _graph_from(source, "Inputs!B2", "Inputs!B3")
    bindings = _bindings(
        _series(series_id="flag_b2", data_range="Inputs!B2", domain={"enum": ["On", "Off"]}),
        _series(series_id="flag_b3", data_range="Inputs!B3", domain={"enum": ["On", "Off"]}),
    )
    dest = tmp_path / "out.xlsx"
    write_workbook(graph, dest, series_bindings=bindings, bindings_workbook=source)
    rows = [row for row in _validation_rows(dest) if row["type"] == "list"]
    assert len(rows) == 1
    assert _cells(rows[0]) == {"B2", "B3"}


def test_real_between_and_one_sided_bounds(tmp_path: Path) -> None:
    source = tmp_path / "source.xlsx"
    wb = fastpyxl.Workbook()
    ws = wb.active
    assert ws is not None
    ws.title = "Inputs"
    ws["A1"] = 0.2
    ws["B1"] = 4
    wb.save(source)
    wb.close()
    graph = _graph_from(source, "Inputs!A1", "Inputs!B1")
    bindings = _bindings(
        _series(
            series_id="share",
            data_range="Inputs!A1",
            dtype="float",
            read="float",
            domain={"real_between": {"min": 0, "max": 1}},
        ),
        _series(
            series_id="floor",
            data_range="Inputs!B1",
            dtype="int",
            read="int",
            domain={"between": {"min": 0}},
        ),
    )
    dest = tmp_path / "out.xlsx"
    write_workbook(graph, dest, series_bindings=bindings, bindings_workbook=source)
    rows = _validation_rows(dest)
    share = _covers(rows, "A1")[0]
    assert share["type"] == "decimal"
    assert share["operator"] == "between"
    assert share["formula1"] == "0"
    assert share["formula2"] == "1"
    floor = _covers(rows, "B1")[0]
    assert floor["type"] == "whole"
    assert floor["operator"] == "greaterThanOrEqual"
    assert floor["formula1"] == "0"
    assert floor["formula2"] is None


def test_relations_become_custom_formulas(tmp_path: Path) -> None:
    source = tmp_path / "source.xlsx"
    wb = fastpyxl.Workbook()
    ws = wb.active
    assert ws is not None
    ws.title = "Inputs"
    ws["G2"] = 1
    ws["H2"] = 10
    ws["I2"] = 3
    wb.save(source)
    wb.close()
    graph = _graph_from(source, "Inputs!G2", "Inputs!H2", "Inputs!I2")
    bindings = _bindings(
        _series(
            series_id="grace",
            data_range="Inputs!G2",
            dtype="int",
            read="int",
            domain={"between": {"min": 0, "max": 50}},
        ),
        _series(
            series_id="maturity",
            data_range="Inputs!H2",
            dtype="int",
            read="int",
            domain={"between": {"min": 0, "max": 80}},
            relations=[{"greater_than": "grace"}],
        ),
        _series(
            series_id="tenor",
            data_range="Inputs!I2",
            dtype="int",
            read="int",
            relations=[{"not_equal": "grace"}],
        ),
    )
    dest = tmp_path / "out.xlsx"
    write_workbook(graph, dest, series_bindings=bindings, bindings_workbook=source)
    rows = _validation_rows(dest)
    maturity = _covers(rows, "H2")[0]
    assert maturity["type"] == "custom"
    assert maturity["formula1"] == "AND(H2>=0,H2<=80,INT(H2)=H2,H2>G2)"
    tenor = _covers(rows, "I2")[0]
    assert tenor["type"] == "custom"
    assert tenor["formula1"] == "I2<>G2"


def test_constant_and_value_map_policy(tmp_path: Path) -> None:
    """Constants are not inputs; value_map needles are the input domain."""
    source = tmp_path / "source.xlsx"
    wb = fastpyxl.Workbook()
    ws = wb.active
    assert ws is not None
    ws.title = "Inputs"
    ws["A1"] = "Borvelia"
    ws["B1"] = "High "
    wb.save(source)
    wb.close()
    graph = _graph_from(source, "Inputs!A1", "Inputs!B1")
    bindings = _bindings(
        _series(
            series_id="country",
            data_range="Inputs!A1",
            direction="constant",
            domain={"from_workbook": True},
        ),
        _series(
            series_id="selector",
            data_range="Inputs!B1",
            value_map={"High": "High ", "Low": "Low"},
        ),
    )
    dest = tmp_path / "out.xlsx"
    write_workbook(graph, dest, series_bindings=bindings, bindings_workbook=source)
    rows = _validation_rows(dest)
    assert _covers(rows, "A1") == []
    selector = _covers(rows, "B1")[0]
    assert selector["type"] == "list"
    assert selector["formula1"] == '"High ,Low"'


def test_include_cell_validation_false_omits_rules(tmp_path: Path) -> None:
    source = tmp_path / "source.xlsx"
    _write_flag_workbook(source)
    graph = _graph_from(source, "Inputs!A1")
    bindings = _bindings(
        _series(series_id="flag", data_range="Inputs!A1", domain={"enum": ["On", "Off"]})
    )
    dest = tmp_path / "out.xlsx"
    write_workbook(
        graph,
        dest,
        series_bindings=bindings,
        bindings_workbook=source,
        include_cell_validation=False,
    )
    assert _validation_rows(dest) == []


def test_union_domain_writes_literal_or_interval_formula(tmp_path: Path) -> None:
    graph = DependencyGraph()
    graph.add_node(make_cell_node("Inputs", "A", 1, value="n.a."))
    graph.sheet_order = ["Inputs"]
    graph.cell_type_env = {
        "Inputs!A1": CellType(
            kind=CellKind.ANY,
            enum=EnumDomain(values=frozenset({"n.a."})),
            real_interval=RealIntervalDomain(min=-1.0, max=1.0),
        )
    }
    dest = tmp_path / "out.xlsx"
    write_workbook(graph, dest)
    row = _covers(_validation_rows(dest), "A1")[0]
    assert row["type"] == "custom"
    assert row["formula1"] == 'OR(A1="n.a.",AND(ISNUMBER(A1),A1>=-1,A1<=1))'


def test_cell_type_env_validates_value_leaves_only(tmp_path: Path) -> None:
    graph = DependencyGraph()
    graph.add_node(make_cell_node("Inputs", "A", 1, value="On"))
    graph.add_node(
        make_cell_node(
            "Inputs",
            "B",
            1,
            normalized_formula='=IF(A1="On",1,0)',
            is_leaf=False,
        )
    )
    graph.sheet_order = ["Inputs"]
    graph.cell_type_env = constraints_to_cell_type_env(
        {
            "Inputs!A1": Literal["On", "Off"],
            "Inputs!B1": Literal["On", "Off"],
        },
        {},
    )
    dest = tmp_path / "out.xlsx"
    write_workbook(graph, dest)
    rows = _validation_rows(dest)
    assert len(rows) == 1
    assert _covers(rows, "A1")[0]["formula1"] == '"Off,On"'
    assert _covers(rows, "B1") == []


def test_projection_uses_projected_cell_type_env(tmp_path: Path) -> None:
    from excel_grapher.exporter.projection import IdentityTransitCompression

    graph = DependencyGraph()
    graph.add_node(make_cell_node("Inputs", "A", 1, value=2))
    graph.sheet_order = ["Inputs"]
    graph.cell_type_env = constraints_to_cell_type_env(
        {"Inputs!A1": Annotated[int, Between(1, 5)]},
        {},
    )
    projection = IdentityTransitCompression().project(graph)
    dest = tmp_path / "out.xlsx"
    write_workbook(projection, dest)
    rows = _validation_rows(dest)
    assert _covers(rows, "A1")[0]["operator"] == "between"


def test_cross_sheet_relation_quotes_the_partner_sheet(tmp_path: Path) -> None:
    graph = DependencyGraph()
    graph.add_node(make_cell_node("Inputs", "A", 1, value=3))
    graph.add_node(make_cell_node("Other Sheet", "B", 1, value=1))
    graph.sheet_order = ["Inputs", "Other Sheet"]
    graph.cell_type_env = constraints_to_cell_type_env(
        {"Inputs!A1": Annotated[int, GreaterThanCell("'Other Sheet'!B1")]},
        {},
    )
    dest = tmp_path / "out.xlsx"
    write_workbook(graph, dest)
    rows = _validation_rows(dest)
    assert _covers(rows, "A1")[0]["formula1"] == "A1>'Other Sheet'!B1"


def test_comma_in_enum_fails_closed_for_inline_list(tmp_path: Path) -> None:
    """Keep refuse-closed when a pure list cannot encode the delimiter."""
    source = tmp_path / "source.xlsx"
    _write_flag_workbook(source)
    graph = _graph_from(source, "Inputs!A1")
    # Force the list path by avoiding any relation; the writer must still
    # refuse an inline list that Excel would split on commas.
    bindings = _bindings(
        _series(
            series_id="flag",
            data_range="Inputs!A1",
            domain={"enum": ["North, South", "East"]},
        )
    )
    dest = tmp_path / "out.xlsx"
    write_workbook(
        graph,
        dest,
        series_bindings=bindings,
        bindings_workbook=source,
    )
    row = _covers(_validation_rows(dest), "A1")[0]
    assert row["type"] == "custom"
    assert "North, South" in (row["formula1"] or "")


def test_relation_partner_missing_from_view_fails_closed(tmp_path: Path) -> None:
    source = tmp_path / "source.xlsx"
    wb = fastpyxl.Workbook()
    ws = wb.active
    assert ws is not None
    ws.title = "Inputs"
    ws["G2"] = 1
    ws["H2"] = 10
    wb.save(source)
    wb.close()
    graph = DependencyGraph()
    graph.add_node(make_cell_node("Inputs", "H", 2, value=10))
    graph.sheet_order = ["Inputs"]
    bindings = _bindings(
        _series(
            series_id="grace",
            data_range="Inputs!G2",
            dtype="int",
            read="int",
        ),
        _series(
            series_id="maturity",
            data_range="Inputs!H2",
            dtype="int",
            read="int",
            relations=[{"greater_than": "grace"}],
        ),
    )
    with pytest.raises(ValueError, match="partner"):
        write_workbook(
            graph,
            tmp_path / "out.xlsx",
            series_bindings=bindings,
            bindings_workbook=source,
        )


def test_from_workbook_input_without_value_fails_closed(tmp_path: Path) -> None:
    graph = DependencyGraph()
    graph.add_node(make_cell_node("Inputs", "A", 1, value=None))
    graph.sheet_order = ["Inputs"]
    bindings = _bindings(
        _series(
            series_id="pinned",
            data_range="Inputs!A1",
            domain={"from_workbook": True},
        )
    )
    with pytest.raises(ValueError, match="cached value"):
        write_workbook(
            graph,
            tmp_path / "out.xlsx",
            series_bindings=bindings,
            include_bound_labels=False,
        )


def test_not_equal_and_blank_enum_member(tmp_path: Path) -> None:
    graph = DependencyGraph()
    graph.add_node(make_cell_node("Inputs", "A", 1, value="On"))
    graph.sheet_order = ["Inputs"]
    graph.cell_type_env = constraints_to_cell_type_env(
        {"Inputs!A1": Literal["On", None]},
        {},
    )
    dest = tmp_path / "out.xlsx"
    write_workbook(graph, dest)
    row = _covers(_validation_rows(dest), "A1")[0]
    assert row["type"] == "list"
    assert row["formula1"] == '"On"'
    assert row["allow_blank"] is True


def test_bool_enum_uses_excel_boolean_literals(tmp_path: Path) -> None:
    graph = DependencyGraph()
    graph.add_node(make_cell_node("Inputs", "A", 1, value=True))
    graph.sheet_order = ["Inputs"]
    graph.cell_type_env = constraints_to_cell_type_env(
        {"Inputs!A1": Literal[True, False]},
        {},
    )
    dest = tmp_path / "out.xlsx"
    write_workbook(graph, dest)
    row = _covers(_validation_rows(dest), "A1")[0]
    assert row["type"] == "custom"
    assert row["formula1"] == "OR(A1=FALSE,A1=TRUE)"


def test_formula_longer_than_excel_limit_fails_closed(tmp_path: Path) -> None:
    graph = DependencyGraph()
    graph.add_node(make_cell_node("Inputs", "A", 1, value="x"))
    graph.sheet_order = ["Inputs"]
    values = tuple(f"item-{index:03d}-" + ("x" * 40) for index in range(8))
    graph.cell_type_env = {
        "Inputs!A1": CellType(
            kind=CellKind.STRING,
            enum=EnumDomain(values=frozenset(values)),
        )
    }
    with pytest.raises(ValueError, match="255"):
        write_workbook(graph, tmp_path / "out.xlsx")


def test_not_equal_cell_metadata_on_env(tmp_path: Path) -> None:
    graph = DependencyGraph()
    graph.add_node(make_cell_node("Inputs", "A", 1, value=1))
    graph.add_node(make_cell_node("Inputs", "B", 1, value=2))
    graph.sheet_order = ["Inputs"]
    graph.cell_type_env = constraints_to_cell_type_env(
        {"Inputs!A1": Annotated[int, NotEqualCell("Inputs!B1"), RealBetween(0.0, 1.0)]},
        {},
    )
    dest = tmp_path / "out.xlsx"
    write_workbook(graph, dest)
    row = _covers(_validation_rows(dest), "A1")[0]
    assert row["type"] == "custom"
    assert row["formula1"] == "AND(A1>=0,A1<=1,A1<>B1)"


def test_bindings_and_env_union_without_conflict(tmp_path: Path) -> None:
    """Label bindings do not drop extract-time env domains on other cells."""
    source = tmp_path / "source.xlsx"
    wb = fastpyxl.Workbook()
    ws = wb.active
    assert ws is not None
    ws.title = "Inputs"
    ws["A1"] = "label"
    ws["B1"] = "On"
    ws["C1"] = 2
    wb.save(source)
    wb.close()
    graph = _graph_from(source, "Inputs!B1", "Inputs!C1")
    graph.cell_type_env = constraints_to_cell_type_env(
        {"Inputs!C1": Annotated[int, Between(1, 5)]},
        {},
    )
    bindings = _bindings(
        _series(
            series_id="flag",
            data_range="Inputs!B1",
            domain={"enum": ["On", "Off"]},
        )
    )
    dest = tmp_path / "out.xlsx"
    write_workbook(graph, dest, series_bindings=bindings, bindings_workbook=source)
    rows = _validation_rows(dest)
    assert _covers(rows, "B1")[0]["type"] == "list"
    year = _covers(rows, "C1")[0]
    assert year["type"] == "whole"
    assert year["formula1"] == "1"
    assert year["formula2"] == "5"


def test_bindings_and_env_conflict_fails_closed(tmp_path: Path) -> None:
    source = tmp_path / "source.xlsx"
    _write_flag_workbook(source)
    graph = _graph_from(source, "Inputs!A1")
    graph.cell_type_env = constraints_to_cell_type_env(
        {"Inputs!A1": Literal["On", "Maybe"]},
        {},
    )
    bindings = _bindings(
        _series(series_id="flag", data_range="Inputs!A1", domain={"enum": ["On", "Off"]})
    )
    with pytest.raises(ValueError, match="conflicting"):
        write_workbook(
            graph,
            tmp_path / "out.xlsx",
            series_bindings=bindings,
            bindings_workbook=source,
        )


def test_comma_in_enum_uses_custom_formula_when_needed(tmp_path: Path) -> None:
    """Commas break inline lists, but custom formulas can quote them."""
    graph = DependencyGraph()
    graph.add_node(make_cell_node("Inputs", "A", 1, value="North, South"))
    graph.sheet_order = ["Inputs"]
    graph.cell_type_env = constraints_to_cell_type_env(
        {"Inputs!A1": Literal["North, South", "East"]},
        {},
    )
    dest = tmp_path / "out.xlsx"
    write_workbook(graph, dest)
    row = _covers(_validation_rows(dest), "A1")[0]
    assert row["type"] == "custom"
    assert row["formula1"] == 'OR(A1="East",A1="North, South")'
