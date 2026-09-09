import inspect
from pathlib import Path

import pytest

from excel_grapher.exporter import CodeGenerator
from excel_grapher.grapher import DependencyGraph
from excel_grapher.series_bindings import validate_bindings_document
from tests.unit.exporter.inverted_tree.helpers import (
    bindings_document,
    generate_inverted,
    inverted_graph_parts,
    load_package,
    series_entry,
    write_workbook,
)


def test_unused_scalar_lookup_error_does_not_escape_lazy_branch(tmp_path: Path) -> None:
    workbook = write_workbook(
        tmp_path / "unused_lookup.xlsx",
        {
            "Sheet1": {
                "A1": "missing",
                "B1": "=VLOOKUP(A1,D1:E2,2,FALSE)",
                "C1": '=IF(A1="use",B1,42)',
            }
        },
    )
    document = bindings_document(
        series_entry("selector", "Sheet1!A1", layout="scalar", direction="input", dtype="string"),
        series_entry("lookup", "Sheet1!B1", layout="scalar", direction="internal", dtype="string"),
        series_entry("result", "Sheet1!C1", layout="scalar", direction="output"),
    )
    modules = generate_inverted(workbook, document, blank_ranges=("Sheet1!D1:E2",))
    package = load_package(modules, tmp_path, name="unused_lookup")
    assert package.compute_result(selector="missing") == 42
    assert package.compute_result(selector="use") == "#N/A"


def test_public_codegen_uses_explicit_inputs(tmp_path: Path) -> None:
    workbook = write_workbook(tmp_path / "public.xlsx", {"Sheet1": {"A1": 3, "B1": "=A1*2"}})
    document = bindings_document(
        series_entry("src", "Sheet1!A1", layout="scalar", direction="input"),
        series_entry("out", "Sheet1!B1", layout="scalar", direction="output"),
    )
    _, _, graph = inverted_graph_parts(workbook, document)
    with CodeGenerator(graph) as generator:
        modules = generator.generate_modules(
            series_bindings=validate_bindings_document(document), bindings_workbook=workbook
        )
    package = load_package(modules, tmp_path, name="public_codegen")
    assert "tensor.py" in modules
    assert package.compute_out(src=7) == 14
    assert package.compute_out.__cells__ == {(): "Sheet1!B1"}
    assert not hasattr(package, "make_context")
    assert not hasattr(package, "set_src")


def test_public_codegen_requires_labels_and_returns_tensor(tmp_path: Path) -> None:
    workbook = write_workbook(
        tmp_path / "years.xlsx",
        {
            "Sheet1": {
                "B1": 2025,
                "C1": 2026,
                "B2": 3,
                "C2": 4,
                "B3": "=B2*2",
                "C3": "=C2*2",
            }
        },
    )
    document = bindings_document(
        series_entry("src", "Sheet1!B2:C2", layout="series", direction="input", header_row=1),
        series_entry("out", "Sheet1!B3:C3", layout="series", direction="output", header_row=1),
    )
    _, _, graph = inverted_graph_parts(workbook, document)
    with CodeGenerator(graph) as generator:
        modules = generator.generate_modules(
            series_bindings=validate_bindings_document(document), bindings_workbook=workbook
        )
    package = load_package(modules, tmp_path, name="named_public")
    assert package.data.SRC_DOMAIN is package.data.OUT_DOMAIN
    assert CodeGenerator.representation_version == package.data.CODEGEN_SCHEMA_VERSION
    tensor = package.data.Src.from_records(
        domain=package.data.SRC_DOMAIN, records=(((2026,), 7.0), ((2025,), 5.0))
    )
    result = package.compute_out(src=tensor)
    assert package.compute_out.__cells__ == {(2025,): "Sheet1!B3", (2026,): "Sheet1!C3"}
    with pytest.raises(TypeError):
        package.compute_out.__cells__[(2025,)] = "Sheet1!A1"
    assert result[2025] == 10
    assert result[2026] == 14
    assert isinstance(result, package.data.Out)
    with pytest.raises(ValueError, match="src.*Tensor"):
        package.compute_out(src=(5, 7))
    wrong = package.Tensor.from_nested(
        domain=package.Domain.product(package.Axis("TIME_PERIOD", (2026, 2027), int)), values=(5, 7)
    )
    with pytest.raises(ValueError, match="src.*2025"):
        package.compute_out(src=wrong)
    assert "Tensor" in modules["data.py"]
    assert "src[time_period]" in modules["internals.py"]
    assert package.internals.out(src=tensor)[2026] == 14
    assert package.as_records(package.compute_out, result) == [
        {"TIME_PERIOD": 2025, "OBS_VALUE": 10},
        {"TIME_PERIOD": 2026, "OBS_VALUE": 14},
    ]


def test_generate_modules_signature_is_bindings_only() -> None:
    params = inspect.signature(CodeGenerator.generate_modules).parameters
    assert list(params) == ["self", "series_bindings", "bindings_workbook", "blank_ranges"]
    for removed in ("targets", "paradigm", "address_helpers", "include_compute_all"):
        assert removed not in params


def test_input_projection_does_not_require_off_graph_coordinates(tmp_path: Path) -> None:
    workbook = write_workbook(
        tmp_path / "projection.xlsx",
        {
            "Sheet1": {
                "B1": 2025,
                "C1": 2026,
                "D1": 2027,
                "B2": 3,
                "C2": 4,
                "D2": 5,
                "C3": "=C2*2",
            }
        },
    )
    document = bindings_document(
        series_entry("src", "Sheet1!B2:D2", layout="series", direction="input", header_row=1),
        series_entry("unused", "Sheet1!B4:C4", layout="series", direction="input", header_row=1),
        series_entry("out", "Sheet1!C3", layout="scalar", direction="output"),
    )
    from excel_grapher.grapher import create_dependency_graph

    graph = create_dependency_graph(workbook, ["Sheet1!C3"], load_values=True)
    with CodeGenerator(graph) as generator:
        modules = generator.generate_modules(
            series_bindings=validate_bindings_document(document), bindings_workbook=workbook
        )
    package = load_package(modules, tmp_path, name="named_projection")
    assert tuple(package.data.SRC_DOMAIN) == ((2025,), (2026,), (2027,))
    assert not hasattr(package.data, "UNUSED_DEFAULT")
    assert package.data.SRC_DEFAULT[2025] == 3
    assert package.data.SRC_DEFAULT[2027] == 5
    src = package.Tensor.from_nested(
        domain=package.Domain.product(package.Axis("TIME_PERIOD", (2026,), int)), values=(9.0,)
    )
    assert package.compute_out(src=src) == 18


def test_four_axis_sparse_generated_function(tmp_path: Path) -> None:
    cells = {"D1": 2026, "E1": 2027, "G1": 2026, "H1": 2027}
    paths = (("loan", "resident", 2025), ("loan", "foreign", 2026), ("bond", "foreign", 2025))
    for row, (instrument, holder, vintage) in enumerate(paths, 2):
        cells.update(
            {
                f"A{row}": instrument,
                f"B{row}": holder,
                f"C{row}": vintage,
                f"D{row}": row * 10,
                f"E{row}": row * 20,
                f"G{row}": f"=D{row}*2",
                f"H{row}": f"=E{row}*2",
            }
        )
    workbook = write_workbook(tmp_path / "four_axis.xlsx", {"Sheet1": cells})
    entries = []
    fields = ("INSTRUMENT", "HOLDER", "ISSUANCE_YEAR", "TIME_PERIOD")
    for sid, region, direction in (("src", "D2:E4", "input"), ("out", "G2:H4", "output")):
        entry = series_entry(
            sid, f"Sheet1!{region}", layout="matrix", direction=direction, key=fields
        )
        entry["structure"]["dimensions"] = [
            {
                "id": field,
                "concept": field,
                "role": "key",
                "scope": "cell",
                "bind": (
                    {"kind": "column_header", "header_row": 1, "read": "int"}
                    if field == "TIME_PERIOD"
                    else {
                        "kind": "row_label",
                        "label_column": column,
                        "read": "int" if field == "ISSUANCE_YEAR" else "string",
                    }
                ),
            }
            for field, column in zip(fields, ("A", "B", "C", "D"), strict=True)
        ]
        entries.append(entry)
    document = bindings_document(*entries)
    _, _, graph = inverted_graph_parts(workbook, document)
    with CodeGenerator(graph) as generator:
        modules = generator.generate_modules(
            series_bindings=validate_bindings_document(document), bindings_workbook=workbook
        )
    package = load_package(modules, tmp_path, name="named_four_axis")
    result = package.compute_out(src=package.data.SRC_DEFAULT)
    assert len(result.domain.axes) == 4
    assert len(result.domain) == 6
    assert result["bond", "foreign", 2025, 2027] == 160
    with pytest.raises(KeyError):
        result["bond", "resident", 2025, 2027]
    assert "src[instrument, holder, issuance_year, time_period]" in modules["internals.py"]
    assert package.Tensor.from_json(result.to_json()).domain == result.domain


def test_single_observation_series_keeps_its_axis(tmp_path: Path) -> None:
    workbook = write_workbook(
        tmp_path / "singleton.xlsx", {"Sheet1": {"B1": 2025, "B2": 3, "B3": "=B2*2"}}
    )
    document = bindings_document(
        series_entry("src", "Sheet1!B2", layout="series", direction="input", header_row=1),
        series_entry("out", "Sheet1!B3", layout="series", direction="output", header_row=1),
    )
    _, _, graph = inverted_graph_parts(workbook, document)
    with CodeGenerator(graph) as generator:
        modules = generator.generate_modules(
            series_bindings=validate_bindings_document(document), bindings_workbook=workbook
        )
    package = load_package(modules, tmp_path, name="named_singleton")
    result = package.compute_out(src=package.data.SRC_DEFAULT)
    assert result[2025] == 6
    assert package.internals.out(src=package.data.SRC_DEFAULT)[2025] == 6


def test_codegen_identity_includes_graph_projection(tmp_path: Path) -> None:
    from excel_grapher.exporter.inverted_tree.catalog import build_catalog
    from excel_grapher.exporter.inverted_tree.named_emit import named_codegen_fingerprint
    from excel_grapher.grapher import create_dependency_graph

    workbook = write_workbook(
        tmp_path / "fingerprint.xlsx",
        {"Sheet1": {"B1": 2025, "C1": 2026, "B2": 2, "C2": 3, "B3": "=B2"}},
    )
    bindings = validate_bindings_document(
        bindings_document(
            series_entry("src", "Sheet1!B2:C2", layout="series", direction="input", header_row=1),
            series_entry("out", "Sheet1!B3", layout="scalar", direction="output"),
        )
    )
    graph = create_dependency_graph(workbook, ["Sheet1!B3"], load_values=True)
    projected = build_catalog(bindings, workbook=workbook, graph=graph)
    authored = build_catalog(bindings, workbook=workbook)
    assert named_codegen_fingerprint(projected) != named_codegen_fingerprint(authored)
    assert named_codegen_fingerprint(projected) == named_codegen_fingerprint(
        build_catalog(bindings, workbook=workbook, graph=graph)
    )


def test_generated_default_distinguishes_blank_and_zero(tmp_path: Path) -> None:
    workbook = write_workbook(
        tmp_path / "blanks.xlsx", {"Sheet1": {"B1": 2025, "C1": 2026, "C2": 0, "B3": "=SUM(B2:C2)"}}
    )
    document = bindings_document(
        series_entry("src", "Sheet1!B2:C2", layout="series", direction="input", header_row=1),
        series_entry("out", "Sheet1!B3", layout="scalar", direction="output"),
    )
    _, _, graph = inverted_graph_parts(workbook, document)
    with CodeGenerator(graph) as generator:
        modules = generator.generate_modules(
            series_bindings=validate_bindings_document(document), bindings_workbook=workbook
        )
    package = load_package(modules, tmp_path, name="named_blank_default")
    assert package.data.SRC_DEFAULT[2025] is None
    assert package.data.SRC_DEFAULT[2026] == 0
    assert package.compute_out(src=package.data.SRC_DEFAULT) == 0
    assert "data.Src[float | str | None]" in modules["api.py"]


def test_numeric_input_retains_supplied_boolean_comparison_semantics(tmp_path: Path) -> None:
    workbook = write_workbook(
        tmp_path / "boolean_measure.xlsx",
        {"Sheet1": {"B1": 2025, "B2": True, "B3": "=IF(B2>100,1,0)"}},
    )
    document = bindings_document(
        series_entry("src", "Sheet1!B2", layout="series", direction="input", header_row=1),
        series_entry("out", "Sheet1!B3", direction="output"),
    )
    from tests.unit.exporter.inverted_tree.helpers import generate_inverted

    package = load_package(
        generate_inverted(workbook, document), tmp_path, name="named_bool_measure"
    )
    supplied = package.Tensor.from_records(
        domain=package.data.SRC_DOMAIN, records=(((2025,), True),)
    )
    assert package.compute_out(src=supplied) == 1.0


@pytest.mark.parametrize("dtype", ["float", "int", "bool", "string", "datetime"])
def test_generated_schema_preserves_excel_error_values(tmp_path: Path, dtype: str) -> None:
    workbook = write_workbook(tmp_path / "error.xlsx", {"Sheet1": {"B1": 2025, "B2": "=1/0"}})
    document = bindings_document(
        series_entry(
            "out", "Sheet1!B2", layout="series", direction="output", header_row=1, dtype=dtype
        )
    )
    _, _, graph = inverted_graph_parts(workbook, document)
    with CodeGenerator(graph) as generator:
        modules = generator.generate_modules(
            series_bindings=validate_bindings_document(document), bindings_workbook=workbook
        )
    package = load_package(modules, tmp_path, name=f"named_error_{dtype}")
    assert package.compute_out()[2025] == "#DIV/0!"
    assert "str" in package.compute_out.__annotations__["return"]


def test_package_generation_requires_bindings() -> None:
    with pytest.raises(TypeError, match="series_bindings.*bindings_workbook"):
        CodeGenerator(DependencyGraph()).generate_modules()  # ty: ignore[missing-argument]


def test_address_keyed_generate_is_removed() -> None:
    assert not hasattr(CodeGenerator, "generate")


def test_generate_modules_rejects_legacy_kwargs() -> None:
    gen = CodeGenerator(DependencyGraph())
    required = {"series_bindings": {"series": []}, "bindings_workbook": "x.xlsx"}
    with pytest.raises(TypeError, match="positional"):
        gen.generate_modules(["Sheet1!A1"], **required)  # ty: ignore[too-many-positional-arguments]
    for extra in (
        {"paradigm": "ctx"},
        {"address_helpers": {}},
        {"include_compute_all": True},
    ):
        with pytest.raises(TypeError, match="unexpected keyword argument"):
            gen.generate_modules(**required, **extra)


@pytest.mark.parametrize(
    "axis_name, input_name",
    [("CLASS", "source"), ("TIME_PERIOD", "time_period"), ("DATA", "source"), ("XL_MUL", "source")],
)
def test_semantic_axis_variables_do_not_shadow_inputs_or_runtime(
    tmp_path: Path, axis_name: str, input_name: str
) -> None:
    workbook = write_workbook(
        tmp_path / "axis_names.xlsx",
        {"M": {"A1": 1, "B1": 2, "A2": 3.0, "B2": 4.0, "A3": "=A2*2", "B3": "=B2*2"}},
    )
    document = bindings_document(
        series_entry(
            input_name,
            "M!A2:B2",
            layout="series",
            direction="input",
            header_row=1,
            key_concept=axis_name,
        ),
        series_entry(
            "result",
            "M!A3:B3",
            layout="series",
            direction="output",
            header_row=1,
            key_concept=axis_name,
        ),
    )
    pkg = load_package(generate_inverted(workbook, document), tmp_path, name="axis_names")
    source = getattr(pkg.data, input_name.upper() + "_DEFAULT")
    result = pkg.internals.result(**{input_name: source})
    assert dict(result.items()) == {(1,): 6.0, (2,): 8.0}


def test_fixed_lookup_table_is_shared_across_semantic_years(tmp_path: Path) -> None:
    from fastpyxl.utils.cell import get_column_letter

    cells = {}
    for i in range(1, 41):
        column = get_column_letter(i)
        cells[f"{column}1"] = 2000 + i
        cells[f"{column}2"] = i * 2.0
        cells[f"{column}3"] = f"=HLOOKUP({column}1,$A$1:$AN$2,2,FALSE)"
    workbook = write_workbook(tmp_path / "shared_lookup.xlsx", {"M": cells})
    document = bindings_document(
        series_entry("years", "M!A1:AN1", layout="series", direction="constant", header_row=1),
        series_entry("lookup_values", "M!A2:AN2", layout="series", direction="input", header_row=1),
        series_entry("result", "M!A3:AN3", layout="series", direction="output", header_row=1),
    )
    modules = generate_inverted(workbook, document)
    pkg = load_package(modules, tmp_path, name="shared_lookup")
    result = pkg.internals.result(
        years=pkg.data.YEARS, lookup_values=pkg.data.LOOKUP_VALUES_DEFAULT
    )
    assert dict(result.items()) == {(2000 + i,): i * 2.0 for i in range(1, 41)}
    assert len(modules["internals.py"]) < 12_000


def test_generated_orders_reuse_authored_coordinate_metadata(tmp_path: Path) -> None:
    cells = {"D1": "=B2"}
    for i in range(200):
        cells[f"A{i + 2}"] = f"instrument with an authored descriptive identifier {i:04d}"
        cells[f"B{i + 2}"] = i + 1.0
    workbook = write_workbook(tmp_path / "coordinate_metadata.xlsx", {"M": cells})
    document = bindings_document(
        series_entry(
            "values",
            "M!B2:B201",
            layout="series",
            direction="input",
            label_column="A",
            key_concept="INSTRUMENT",
            key_read="string",
        ),
        series_entry("result", "M!D1", direction="output"),
    )
    modules = generate_inverted(workbook, document)
    pkg = load_package(modules, tmp_path, name="coordinate_metadata")
    assert pkg.compute_result(values=pkg.data.VALUES_DEFAULT) == 1.0
    assert list(pkg.data.VALUES_DEFAULT.items())[-1] == (
        ("instrument with an authored descriptive identifier 0199",),
        200.0,
    )
    assert len(modules["data.py"]) < 32_000
