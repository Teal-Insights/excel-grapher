"""`Model.from_workbook` binds declared inputs, including typed-over formulas."""

from __future__ import annotations

import subprocess
import zipfile
from pathlib import Path

import pytest
from fastpyxl import load_workbook

from tests.unit.exporter.inverted_tree.helpers import (
    bindings_document,
    call_compute,
    generate_inverted,
    inverted_graph_parts,
    load_package,
    named_input_kwargs,
    series_entry,
    write_workbook,
)
from tests.unit.exporter.inverted_tree.test_dynamic_axis_labels import (
    _labelled_bindings,
    _labelled_workbook,
)
from tests.unit.exporter.inverted_tree.test_input_domain import (
    _enum_flag_bindings,
    _enum_flag_workbook,
)


def _seed_workbook(path: Path, seed: float) -> Path:
    return write_workbook(path, {"Inputs": {"A1": seed, "B1": "=A1*3"}})


def _seed_bindings() -> dict:
    return bindings_document(
        series_entry("seed", "Inputs!A1", layout="scalar", direction="input"),
        series_entry("result", "Inputs!B1", layout="scalar", direction="output"),
    )


def _constant_workbook(path: Path, *, seed: float, factor: float) -> Path:
    return write_workbook(
        path,
        {
            "Inputs": {"A1": seed, "B1": factor},
            "Outputs": {"A1": "=Inputs!A1*Inputs!B1"},
        },
    )


def _constant_bindings() -> dict:
    return bindings_document(
        series_entry("seed", "Inputs!A1", layout="scalar", direction="input"),
        series_entry("factor", "Inputs!B1", layout="scalar", direction="constant"),
        series_entry("result", "Outputs!A1", layout="scalar", direction="output"),
    )


def _rate_workbook(path: Path, left: float, right: float) -> Path:
    return write_workbook(
        path,
        {
            "Inputs": {"B1": left, "C1": right, "B10": 1, "C10": 2},
            "Outputs": {"A1": "=Inputs!B1+Inputs!C1", "B1": 0, "A10": 1, "B10": 2},
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
        ),
        series_entry(
            "out",
            "Outputs!A1",
            layout="scalar",
            direction="output",
        ),
    )


def test_from_workbook_emits_mixin_and_catalog_input_ids(tmp_path: Path) -> None:
    modules = generate_inverted(_seed_workbook(tmp_path / "template.xlsx", 2.0), _seed_bindings())
    model = modules["model.py"]
    api = modules["api.py"]
    assert "class _BoundInputs:" in model
    assert "def from_workbook" in model
    assert "cls.__annotations__" not in model
    assert '_INPUT_IDS: tuple[str, ...] = ("seed",)' in model
    assert "class Model(_BoundInputs):" in model
    assert "class Model" not in api
    assert "workbook.py" in modules
    assert "def compute_result" in api
    assert "workbook" not in api[api.index("def compute_result") :]


def test_from_workbook_reads_case_workbook_and_overrides_win(tmp_path: Path) -> None:
    template = _seed_workbook(tmp_path / "template.xlsx", 2.0)
    pkg = load_package(
        generate_inverted(template, _seed_bindings()),
        tmp_path,
        name="from_workbook_seed",
    )
    Model = pkg.Model
    case = _seed_workbook(tmp_path / "case.xlsx", 4.0)
    extra = write_workbook(
        tmp_path / "extra_sheets.xlsx",
        {"Inputs": {"A1": 4.0, "B1": "=A1*3"}, "Notes": {"Z9": "ignored"}},
    )

    assert pkg.data.SEED_DEFAULT == 2.0
    assert Model.from_workbook(case).result == 12.0
    assert Model.from_workbook(extra).result == 12.0
    assert Model.from_workbook(case, seed=5.0).result == 15.0
    assert pkg.compute_result(pkg.ResultInputs(seed=Model.from_workbook(case).seed)) == 12.0
    assert pkg.compute_result(pkg.ResultInputs.from_workbook(case)) == 12.0


def test_from_workbook_missing_input_sheet_raises(tmp_path: Path) -> None:
    pkg = load_package(
        generate_inverted(_seed_workbook(tmp_path / "template.xlsx", 2.0), _seed_bindings()),
        tmp_path,
        name="from_workbook_missing",
    )
    missing = write_workbook(tmp_path / "missing.xlsx", {"Other": {"A1": 4.0}})
    with pytest.raises(KeyError, match="Inputs!A1"):
        pkg.Model.from_workbook(missing)


def test_from_workbook_rejects_unknown_overrides(tmp_path: Path) -> None:
    pkg = load_package(
        generate_inverted(_seed_workbook(tmp_path / "template.xlsx", 2.0), _seed_bindings()),
        tmp_path,
        name="from_workbook_unknown",
    )
    case = _seed_workbook(tmp_path / "case.xlsx", 4.0)
    with pytest.raises(TypeError, match="unknown inputs"):
        pkg.Model.from_workbook(case, discount_rate=0.04)


def test_from_workbook_leaves_constants_on_codegen_snapshot(tmp_path: Path) -> None:
    template = _constant_workbook(tmp_path / "template.xlsx", seed=2.0, factor=3.0)
    pkg = load_package(
        generate_inverted(template, _constant_bindings()),
        tmp_path,
        name="from_workbook_const",
    )
    Model = pkg.Model
    case = _constant_workbook(tmp_path / "case.xlsx", seed=4.0, factor=9.0)
    assert pkg.data.FACTOR == 3.0
    assert Model.from_workbook(case).result == 12.0
    with pytest.raises(TypeError, match="unknown inputs"):
        Model.from_workbook(case, factor=5.0)
    with pkg.data.overrides(FACTOR=5.0):
        assert Model.from_workbook(case).result == 20.0


def test_from_workbook_builds_tensors_with_records(tmp_path: Path) -> None:
    template = _rate_workbook(tmp_path / "template.xlsx", 0.25, 0.5)
    pkg = load_package(
        generate_inverted(template, _rate_bindings()),
        tmp_path,
        name="from_workbook_rate",
    )
    case = _rate_workbook(tmp_path / "case.xlsx", 1.5, 2.5)
    model = pkg.Model.from_workbook(case)
    assert model.out == pytest.approx(4.0)
    overlay = pkg.data.RATE.with_records((((1,), 3.0), ((2,), 4.0)))
    assert pkg.Model.from_workbook(case, rate=overlay).out == pytest.approx(7.0)


def test_from_workbook_uses_init_validation_checks(tmp_path: Path) -> None:
    pkg = load_package(
        generate_inverted(_enum_flag_workbook(tmp_path), _enum_flag_bindings()),
        tmp_path,
        name="from_workbook_flag",
    )
    Model = pkg.Model
    valid = write_workbook(
        tmp_path / "flag_case.xlsx",
        {"Inputs": {"A1": 1}, "Outputs": {"A1": "=Inputs!A1"}},
    )
    invalid = write_workbook(
        tmp_path / "flag_bad.xlsx",
        {"Inputs": {"A1": 2}, "Outputs": {"A1": "=Inputs!A1"}},
    )
    assert Model.from_workbook(valid).out == 1
    with pytest.raises(ValueError, match="flag out of domain"):
        Model.from_workbook(invalid)


def test_from_workbook_reuses_deferred_runtime_labeller_init(tmp_path: Path) -> None:
    workbook = _labelled_workbook(tmp_path)
    pkg = load_package(
        generate_inverted(workbook, _labelled_bindings()),
        tmp_path,
        name="from_workbook_labelled",
    )
    catalog, _deps, graph = inverted_graph_parts(workbook, _labelled_bindings())
    expected = call_compute(pkg, "path", named_input_kwargs(pkg, catalog, graph))
    model = pkg.Model.from_workbook(workbook)
    assert list(model.path.items()) == list(expected.items())
    shifted = _labelled_workbook(tmp_path, first=2025, shock_year=2026)
    shifted_model = pkg.Model.from_workbook(shifted)
    assert shifted_model.first_year == 2025
    assert tuple(shifted_model.growth.domain.axes[0].keys) == (2025, 2026, 2027)
    assert tuple(value for _coord, value in shifted_model.growth.items()) == (3.5, 4.0, 4.5)
    assert tuple(shifted_model.path.domain.axes[0].keys) == (2025, 2026, 2027)
    assert [value for _coord, value in shifted_model.path.items()] == pytest.approx([3.5, 4.5, 5.5])


def _formula_header_workbook(path: Path, *, anchor: int, left: float, right: float) -> Path:
    """Value row keyed by formula headers driven by `anchor`."""
    return write_workbook(
        path,
        {
            "Inputs": {
                "A1": anchor,
                "A2": "=A1",
                "B2": "=A2+1",
                "A3": left,
                "B3": right,
                "A4": "=A3+1",
                "B4": "=B3+1",
            }
        },
    )


def _formula_header_bindings() -> dict:
    return bindings_document(
        series_entry("anchor", "Inputs!A1", direction="input", dtype="int"),
        series_entry(
            "amount",
            "Inputs!A3:B3",
            layout="series",
            direction="input",
            header_row=2,
        ),
        series_entry(
            "path",
            "Inputs!A4:B4",
            layout="series",
            direction="output",
            header_row=2,
        ),
    )


def _literal_header_workbook(path: Path, *, left_year: int, right_year: int) -> Path:
    return write_workbook(
        path,
        {
            "Inputs": {
                "A2": left_year,
                "B2": right_year,
                "A3": 10,
                "B3": 20,
                "A4": "=A3+1",
                "B4": "=B3+1",
            }
        },
    )


def _literal_header_bindings() -> dict:
    return bindings_document(
        series_entry(
            "amount",
            "Inputs!A3:B3",
            layout="series",
            direction="input",
            header_row=2,
        ),
        series_entry(
            "path",
            "Inputs!A4:B4",
            layout="series",
            direction="output",
            header_row=2,
        ),
    )


def _row_label_workbook(path: Path, *, anchor: int, top: float, bottom: float) -> Path:
    return write_workbook(
        path,
        {
            "Inputs": {
                "A1": anchor,
                "A2": "=A1",
                "A3": "=A2+1",
                "B2": top,
                "B3": bottom,
                "C2": "=B2+1",
                "C3": "=B3+1",
            }
        },
    )


def _row_label_bindings() -> dict:
    return bindings_document(
        series_entry("anchor", "Inputs!A1", direction="input", dtype="int"),
        series_entry(
            "amount",
            "Inputs!B2:B3",
            layout="series",
            direction="input",
            label_column="A",
        ),
        series_entry(
            "path",
            "Inputs!C2:C3",
            layout="series",
            direction="output",
            label_column="A",
        ),
    )


def test_from_workbook_rekeys_formula_column_headers(tmp_path: Path) -> None:
    template = _formula_header_workbook(tmp_path / "template.xlsx", anchor=2024, left=10, right=20)
    _plant_formula_cache(template, "A2", "2024")
    _plant_formula_cache(template, "B2", "2025")
    pkg = load_package(
        generate_inverted(template, _formula_header_bindings()),
        tmp_path,
        name="from_workbook_formula_headers",
    )
    Model = pkg.Model
    assert tuple(Model.from_workbook(template).amount.domain.axes[0].keys) == (2024, 2025)
    assert Model.from_workbook(template).path[2024] == pytest.approx(11)
    assert Model.from_defaults().amount.domain.axes[0].keys == (2024, 2025)

    case = _formula_header_workbook(tmp_path / "case.xlsx", anchor=2025, left=7, right=8)
    _plant_formula_cache(case, "A2", "2025")
    _plant_formula_cache(case, "B2", "2026")
    model = Model.from_workbook(case)
    assert model.anchor == 2025
    assert tuple(model.amount.domain.axes[0].keys) == (2025, 2026)
    assert tuple(value for _coord, value in model.amount.items()) == (7, 8)
    assert tuple(model.path.domain.axes[0].keys) == (2025, 2026)
    assert model.path[2025] == pytest.approx(8)
    assert model.path[2026] == pytest.approx(9)


def test_from_workbook_formula_headers_fail_closed(tmp_path: Path) -> None:
    template = _formula_header_workbook(tmp_path / "template.xlsx", anchor=2024, left=10, right=20)
    _plant_formula_cache(template, "A2", "2024")
    _plant_formula_cache(template, "B2", "2025")
    pkg = load_package(
        generate_inverted(template, _formula_header_bindings()),
        tmp_path,
        name="from_workbook_formula_headers_closed",
    )
    missing = _formula_header_workbook(tmp_path / "missing.xlsx", anchor=2025, left=7, right=8)
    with pytest.raises(pkg.tensor.SchemaError, match="no cached value"):
        pkg.Model.from_workbook(missing)
    duplicate = _formula_header_workbook(tmp_path / "duplicate.xlsx", anchor=2025, left=7, right=8)
    _plant_formula_cache(duplicate, "A2", "2025")
    _plant_formula_cache(duplicate, "B2", "2025")
    with pytest.raises(pkg.tensor.SchemaError, match="duplicate"):
        pkg.Model.from_workbook(duplicate)


def test_from_workbook_keeps_literal_column_headers(tmp_path: Path) -> None:
    template = _literal_header_workbook(tmp_path / "template.xlsx", left_year=2024, right_year=2025)
    pkg = load_package(
        generate_inverted(template, _literal_header_bindings()),
        tmp_path,
        name="from_workbook_literal_headers",
    )
    edited = _literal_header_workbook(tmp_path / "edited.xlsx", left_year=2099, right_year=2100)
    model = pkg.Model.from_workbook(edited)
    assert tuple(model.amount.domain.axes[0].keys) == (2024, 2025)
    assert model.path[2024] == pytest.approx(11)


def test_from_workbook_rekeys_formula_row_labels(tmp_path: Path) -> None:
    template = _row_label_workbook(tmp_path / "template.xlsx", anchor=2024, top=10, bottom=20)
    _plant_formula_cache(template, "A2", "2024")
    _plant_formula_cache(template, "A3", "2025")
    pkg = load_package(
        generate_inverted(template, _row_label_bindings()),
        tmp_path,
        name="from_workbook_formula_rows",
    )
    case = _row_label_workbook(tmp_path / "case.xlsx", anchor=2025, top=7, bottom=8)
    _plant_formula_cache(case, "A2", "2025")
    _plant_formula_cache(case, "A3", "2026")
    model = pkg.Model.from_workbook(case)
    assert tuple(model.amount.domain.axes[0].keys) == (2025, 2026)
    assert model.path[2026] == pytest.approx(9)


def _override_workbook(path: Path, *, seed: float = 0.05, slot: object = "=A1") -> Path:
    """Template: `B1` echoes `A1` until an economist types over the formula."""
    return write_workbook(
        path,
        {"Inputs": {"A1": seed, "B1": slot, "C1": "=IF(ISNUMBER(B1),B1,A1)"}},
    )


def _override_bindings(*, domain: dict | None = None) -> dict:
    slot = series_entry("slot", "Inputs!B1", layout="scalar", direction="input")
    slot["input"]["mode"] = "override"
    if domain is not None:
        slot["input"]["domain"] = domain
    return bindings_document(
        series_entry("seed", "Inputs!A1", layout="scalar", direction="input"),
        slot,
        series_entry("result", "Inputs!C1", layout="scalar", direction="output"),
    )


def _plant_formula_cache(path: Path, coordinate: str, cached: str) -> None:
    """Keep the formula and replace its saved calculated value."""
    sheet = "xl/worksheets/sheet1.xml"
    with zipfile.ZipFile(path) as archive:
        xml = archive.read(sheet).decode()
        others = {name: archive.read(name) for name in archive.namelist() if name != sheet}
    marker = f'<c r="{coordinate}">'
    start = xml.find(marker)
    end = xml.find("</c>", start)
    if start < 0 or end < 0:
        raise AssertionError(f"{coordinate} formula cell not found")
    cell = xml[start:end]
    if "<f>" not in cell:
        raise AssertionError(f"{coordinate} is not a formula cell")
    if "<v />" in cell:
        cell = cell.replace("<v />", f"<v>{cached}</v>", 1)
    elif "<v/>" in cell:
        cell = cell.replace("<v/>", f"<v>{cached}</v>", 1)
    else:
        close = cell.find("</v>")
        open_end = cell.find(">", cell.find("<v"))
        if close < 0 or open_end < 0:
            raise AssertionError(f"{coordinate} has no cached value element")
        cell = cell[: open_end + 1] + cached + cell[close:]
    xml = xml[:start] + cell + xml[end:]
    with zipfile.ZipFile(path, "w") as archive:
        for name, payload in others.items():
            archive.writestr(name, payload)
        archive.writestr(sheet, xml)


def test_from_workbook_displaces_typed_over_formula(tmp_path: Path) -> None:
    template = _override_workbook(tmp_path / "template.xlsx")
    pkg = load_package(
        generate_inverted(template, _override_bindings()),
        tmp_path,
        name="from_workbook_override",
    )
    Model = pkg.Model
    assert pkg.data.SLOT_DEFAULT is pkg.tensor.UNSET
    assert "slot" in Model._INPUT_IDS

    assert Model.from_defaults().result == pytest.approx(0.05)
    assert Model.from_workbook(template).result == pytest.approx(0.05)
    assert Model(seed=0.07).result == pytest.approx(0.07)
    assert Model.from_defaults(slot=0.09).result == pytest.approx(0.09)
    assert pkg.compute_result(pkg.ResultInputs.from_defaults()) == pytest.approx(0.05)
    assert pkg.compute_result(pkg.ResultInputs.from_defaults(seed=0.07)) == pytest.approx(0.07)

    typed = _override_workbook(tmp_path / "typed.xlsx", slot=0.09)
    assert Model.from_workbook(typed).result == pytest.approx(0.09)
    assert Model.from_workbook(typed, slot=0.04).result == pytest.approx(0.04)

    stale = _override_workbook(tmp_path / "stale.xlsx", seed=0.07)
    _plant_formula_cache(stale, "B1", "0.05")
    assert Model.from_workbook(stale).result == pytest.approx(0.07)

    typed_output = _override_workbook(tmp_path / "typed_output.xlsx", slot="=A1")
    book = load_workbook(typed_output)
    book["Inputs"]["C1"] = 123
    book.save(typed_output)
    book.close()
    assert Model.from_workbook(typed_output).result == pytest.approx(0.05)


def test_override_package_imports_unset_and_passes_ruff(tmp_path: Path) -> None:
    modules = generate_inverted(
        _override_workbook(tmp_path / "template.xlsx"),
        _override_bindings(),
    )
    package = tmp_path / "override_lint"
    package.mkdir()
    for filename, content in modules.items():
        (package / filename).write_text(content, encoding="utf-8")
    assert "from .tensor import Domain, Series, Unset" in modules["internals.py"]
    targets = [str(package / name) for name in ("api.py", "model.py", "validation.py")]
    check = subprocess.run(
        ["uv", "run", "--no-sync", "ruff", "check", *targets],
        check=False,
        capture_output=True,
        text=True,
    )
    assert check.returncode == 0, check.stdout + check.stderr
    formatted = subprocess.run(
        ["uv", "run", "--no-sync", "ruff", "format", "--check", *targets],
        check=False,
        capture_output=True,
        text=True,
    )
    assert formatted.returncode == 0, formatted.stdout + formatted.stderr
    undefined = subprocess.run(
        [
            "uv",
            "run",
            "--no-sync",
            "ruff",
            "check",
            "--select",
            "F821",
            str(package / "internals.py"),
        ],
        check=False,
        capture_output=True,
        text=True,
    )
    assert undefined.returncode == 0, undefined.stdout + undefined.stderr


def test_from_workbook_override_literal_outside_domain_fails(tmp_path: Path) -> None:
    domain = {"real_between": {"min": 0, "max": 1}}
    template = _override_workbook(tmp_path / "template.xlsx")
    pkg = load_package(
        generate_inverted(template, _override_bindings(domain=domain)),
        tmp_path,
        name="from_workbook_override_domain",
    )
    outside = _override_workbook(tmp_path / "outside.xlsx", slot=1.5)
    with pytest.raises(ValueError, match="slot out of domain"):
        pkg.Model.from_workbook(outside)
    stale = _override_workbook(tmp_path / "stale_outside.xlsx", seed=0.07)
    _plant_formula_cache(stale, "B1", "9")
    assert pkg.Model.from_workbook(stale).result == pytest.approx(0.07)


def _override_row_workbook(
    path: Path,
    *,
    left: float = 0.05,
    right: float = 0.06,
    left_slot: object = "=B1",
    right_slot: object = "=C1",
) -> Path:
    return write_workbook(
        path,
        {
            "Inputs": {
                "B1": left,
                "C1": right,
                "B2": left_slot,
                "C2": right_slot,
                "B3": "=IF(ISNUMBER(B2),B2,B1)",
                "C3": "=IF(ISNUMBER(C2),C2,C1)",
                "B10": 1,
                "C10": 2,
            }
        },
    )


def _override_row_bindings() -> dict:
    slot = series_entry(
        "slot",
        "Inputs!B2:C2",
        layout="series",
        direction="input",
        header_row=10,
    )
    slot["input"]["mode"] = "override"
    return bindings_document(
        series_entry(
            "seed",
            "Inputs!B1:C1",
            layout="series",
            direction="input",
            header_row=10,
        ),
        slot,
        series_entry(
            "result",
            "Inputs!B3:C3",
            layout="series",
            direction="output",
            header_row=10,
        ),
    )


def test_from_workbook_displaces_one_cell_of_an_override_series(tmp_path: Path) -> None:
    template = _override_row_workbook(tmp_path / "template.xlsx")
    pkg = load_package(
        generate_inverted(template, _override_row_bindings()),
        tmp_path,
        name="from_workbook_override_row",
    )
    Model = pkg.Model
    baseline = Model.from_defaults()
    assert baseline.result[1] == pytest.approx(0.05)
    assert baseline.result[2] == pytest.approx(0.06)
    assert Model.from_workbook(template).result[1] == pytest.approx(0.05)
    assert Model.from_workbook(template).result[2] == pytest.approx(0.06)

    shifted = _override_row_workbook(tmp_path / "shifted.xlsx", left=0.2, right=0.3)
    assert Model.from_workbook(shifted).result[1] == pytest.approx(0.2)
    assert Model.from_workbook(shifted).result[2] == pytest.approx(0.3)

    partial = _override_row_workbook(
        tmp_path / "partial.xlsx",
        left=0.05,
        right=0.11,
        left_slot=0.09,
        right_slot="=C1",
    )
    model = Model.from_workbook(partial)
    assert model.result[1] == pytest.approx(0.09)
    assert model.result[2] == pytest.approx(0.11)
