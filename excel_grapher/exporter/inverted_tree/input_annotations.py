"""Public input annotations read from `graph.cell_type_env`."""

from __future__ import annotations

import json
from dataclasses import dataclass
from typing import TYPE_CHECKING, Literal

from excel_grapher.core.cell_types import CellType, normalize_cell_type_env_key
from excel_grapher.exporter.inverted_tree.errors import InvertedTreeExportError
from excel_grapher.series_bindings.input_coerce import input_value_map_from_series

if TYPE_CHECKING:
    from collections.abc import Mapping

    from excel_grapher.exporter.inverted_tree.catalog import BoundSeries, SeriesCatalog
    from excel_grapher.grapher.graph import DependencyGraph

_MarkerKind = Literal["public", "pin", "open"]


@dataclass(frozen=True, slots=True)
class _PublicDomain:
    """One caller-facing constraint shared by every cell of an input series."""

    kind: Literal["enum", "between", "real_between"]
    enum: frozenset[object] | None = None
    minimum: int | float | None = None
    maximum: int | float | None = None


@dataclass(frozen=True, slots=True)
class _Marker:
    kind: _MarkerKind
    domain: _PublicDomain | None = None


def public_input_annotations(
    catalog: SeriesCatalog,
    graph: DependencyGraph,
) -> dict[str, str]:
    """Return element annotations for public inputs constrained on the graph.

    Singleton literals, `Literal[None]` blanks, and `constant` series are not
    public types. A series whose cells do not share one caller-facing domain
    fails closed. `input.value_map` workbook needles stay off the annotation;
    callers are still checked against the map keys. A missing `cell_type_env`
    yields no annotations; the check does not expand a lazy domain index.

    Args:
        catalog: Bound series for this export.
        graph: Graph whose `cell_type_env` holds the extracted domains.

    Returns:
        Series id to a `Literal` or `Annotated` element annotation. Inputs
        with no public domain are omitted.

    Raises:
        InvertedTreeExportError: A public input spans incompatible domains, or
            an interval kind disagrees with the measure dtype.
    """
    env = graph.cell_type_env
    if env is None:
        return {}
    annotations: dict[str, str] = {}
    for series in catalog.series.values():
        if series.direction != "input":
            continue
        if series.graph_cells is not None and not series.graph_cells:
            continue
        rendered = _annotation_for_series(series, env)
        if rendered is not None:
            annotations[series.series_id] = rendered
    return annotations


def _annotation_for_series(series: BoundSeries, env: Mapping[str, CellType]) -> str | None:
    if not series.cells:
        return None
    markers = [
        (
            address,
            _classify(
                _lookup(env, address),
                series_id=series.series_id,
                address=address,
            ),
        )
        for address in series.cells
    ]
    public = [(address, marker) for address, marker in markers if marker.kind == "public"]
    if not public:
        return None
    first_address, first = public[0]
    for address, marker in public[1:]:
        if marker.domain != first.domain:
            _fail(series.series_id, first, first_address, marker, address)
    other = next(
        ((address, marker) for address, marker in markers if marker.kind != "public"),
        None,
    )
    if other is not None:
        address, marker = other
        _fail(series.series_id, first, first_address, marker, address)
    domain = first.domain
    assert domain is not None
    if _is_value_map_needle(series, domain):
        return None
    _check_measure_dtype(series, domain)
    return _render(domain)


def _lookup(env: Mapping[str, CellType], address: str) -> CellType | None:
    key = normalize_cell_type_env_key(address)
    domain_for = getattr(env, "domain_for", None)
    if callable(domain_for):
        found = domain_for(key)
        return found if isinstance(found, CellType) else None
    found = env.get(key)
    return found if isinstance(found, CellType) else None


def _classify(cell_type: CellType | None, *, series_id: str, address: str) -> _Marker:
    if cell_type is None:
        return _Marker("open")
    kinds: list[str] = []
    if cell_type.enum is not None:
        kinds.append("enum")
    if cell_type.interval is not None:
        kinds.append("between")
    if cell_type.real_interval is not None:
        kinds.append("real_between")
    if len(kinds) > 1:
        raise InvertedTreeExportError(
            f"series {series_id!r}: incompatible constraint domains on {address} "
            f"({' and '.join(kinds)})"
        )
    if cell_type.enum is not None:
        values = cell_type.enum.values
        if not values:
            raise InvertedTreeExportError(
                f"series {series_id!r}: incompatible constraint domains on {address} (empty enum)"
            )
        if len(values) == 1:
            return _Marker("pin")
        return _Marker("public", _PublicDomain("enum", enum=frozenset(values)))
    if cell_type.interval is not None:
        return _Marker(
            "public",
            _PublicDomain(
                "between",
                minimum=cell_type.interval.min,
                maximum=cell_type.interval.max,
            ),
        )
    if cell_type.real_interval is not None:
        return _Marker(
            "public",
            _PublicDomain(
                "real_between",
                minimum=cell_type.real_interval.min,
                maximum=cell_type.real_interval.max,
            ),
        )
    return _Marker("open")


def _is_value_map_needle(series: BoundSeries, domain: _PublicDomain) -> bool:
    """True when an enum is the workbook needle set, not the caller-facing keys."""
    if domain.kind != "enum" or domain.enum is None:
        return False
    mapping = input_value_map_from_series(series.raw)
    if mapping is None:
        return False
    needles = frozenset(mapping.values())
    keys = frozenset(mapping)
    return domain.enum == needles and domain.enum != keys


def _check_measure_dtype(series: BoundSeries, domain: _PublicDomain) -> None:
    if domain.kind == "between" and series.python_dtype != "int":
        raise InvertedTreeExportError(
            f"series {series.series_id!r}: constraint Between requires integer "
            f"measure dtype, got {series.dtype!r}"
        )
    if domain.kind == "real_between" and series.python_dtype != "float":
        raise InvertedTreeExportError(
            f"series {series.series_id!r}: constraint RealBetween requires float "
            f"measure dtype, got {series.dtype!r}"
        )


def _fail(
    series_id: str,
    left: _Marker,
    left_address: str,
    right: _Marker,
    right_address: str,
) -> None:
    raise InvertedTreeExportError(
        f"series {series_id!r}: incompatible constraint domains "
        f"({_label(left)} on {left_address}, {_label(right)} on {right_address})"
    )


def _label(marker: _Marker) -> str:
    if marker.kind == "open":
        return "unconstrained"
    if marker.kind == "pin":
        return "extract-only pin"
    assert marker.domain is not None
    return _domain_text(marker.domain)


def _domain_text(domain: _PublicDomain) -> str:
    if domain.kind == "enum":
        assert domain.enum is not None
        rendered = ", ".join(_render_member(value) for value in _sorted_members(domain.enum))
        return f"enum {{{rendered}}}"
    name = "between" if domain.kind == "between" else "real_between"
    return f"{name}(min={domain.minimum!r}, max={domain.maximum!r})"


def _render(domain: _PublicDomain) -> str:
    if domain.kind == "enum":
        assert domain.enum is not None
        members = ", ".join(_render_member(value) for value in _sorted_members(domain.enum))
        return f"Literal[{members}]"
    if domain.kind == "between":
        return (
            "Annotated[int, Between("
            f"{_render_bound(domain.minimum)}, {_render_bound(domain.maximum)})]"
        )
    return (
        "Annotated[float, RealBetween("
        f"{_render_bound(domain.minimum)}, {_render_bound(domain.maximum)})]"
    )


def _sorted_members(values: frozenset[object]) -> tuple[object, ...]:
    return tuple(sorted(values, key=lambda value: (type(value).__name__, repr(value))))


def _render_member(value: object) -> str:
    if isinstance(value, bool):
        return "True" if value else "False"
    if value is None:
        return "None"
    if isinstance(value, str):
        return json.dumps(value)
    if isinstance(value, int | float):
        return repr(value)
    raise InvertedTreeExportError(
        f"cannot emit a public Literal member of type {type(value).__name__}"
    )


def _render_bound(value: int | float | None) -> str:
    if value is None:
        return "None"
    return repr(value)
