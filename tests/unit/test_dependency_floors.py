"""Declared dependency floors must cover the APIs the package calls."""

from importlib.metadata import requires

from packaging.requirements import Requirement
from packaging.version import Version


def test_fastpyxl_floor_supports_keep_formula_cache() -> None:
    """`load_wb` passes `keep_formula_cache=`, which fastpyxl added in 1.1.0."""
    reqs = [Requirement(r) for r in requires("excel-grapher") or []]
    (fastpyxl,) = [r for r in reqs if r.name == "fastpyxl"]
    assert not fastpyxl.specifier.contains(Version("1.0.12"))
    assert fastpyxl.specifier.contains(Version("1.1.0"))
