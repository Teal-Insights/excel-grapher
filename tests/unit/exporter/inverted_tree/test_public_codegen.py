import inspect
from pathlib import Path

import pytest

from excel_grapher.exporter import CodeGenerator
from excel_grapher.grapher import DependencyGraph
from excel_grapher.series_bindings import validate_bindings_document
from tests.unit.exporter.inverted_tree.helpers import (
    bindings_document,
    inverted_graph_parts,
    load_package,
    series_entry,
    write_workbook,
)


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
    assert set(modules) == {"__init__.py", "api.py", "data.py", "internals.py", "runtime.py"}
    assert package.compute_out(src=7) == (14,)
    assert not hasattr(package, "make_context")
    assert not hasattr(package, "set_src")


def test_generate_modules_signature_is_bindings_only() -> None:
    params = inspect.signature(CodeGenerator.generate_modules).parameters
    assert list(params) == ["self", "series_bindings", "bindings_workbook", "blank_ranges"]
    for removed in ("targets", "paradigm", "address_helpers", "include_compute_all"):
        assert removed not in params


def test_package_generation_requires_bindings() -> None:
    with pytest.raises(TypeError, match="series_bindings.*bindings_workbook"):
        CodeGenerator(DependencyGraph()).generate_modules()  # ty: ignore[missing-argument]


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
