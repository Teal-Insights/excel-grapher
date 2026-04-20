"""Package-boundary guards (#101).

`excel_grapher.grapher` must not import from `excel_grapher.evaluator` or
`excel_grapher.exporter`; reverse direction is allowed.
"""

from __future__ import annotations

import importlib
import pkgutil


def test_grapher_does_not_import_evaluator_or_exporter() -> None:
    import excel_grapher.grapher as grapher_pkg

    offenders: list[tuple[str, str]] = []
    for mod_info in pkgutil.walk_packages(grapher_pkg.__path__, prefix="excel_grapher.grapher."):
        mod = importlib.import_module(mod_info.name)
        for name, value in vars(mod).items():
            origin = getattr(value, "__module__", "") or ""
            if origin.startswith("excel_grapher.evaluator") or origin.startswith(
                "excel_grapher.exporter"
            ):
                offenders.append((mod_info.name, f"{name} <- {origin}"))
    assert not offenders, f"grapher leaked imports from evaluator/exporter: {offenders}"
