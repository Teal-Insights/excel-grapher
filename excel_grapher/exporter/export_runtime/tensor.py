"""Immutable named-coordinate values shared by generated model functions."""

from __future__ import annotations

import hashlib
import json
from collections.abc import Iterable, Iterator, Mapping, Sequence
from dataclasses import dataclass, field
from datetime import date, datetime
from itertools import product
from types import MappingProxyType
from typing import Any, ClassVar, Generic, Self, TypeVar, cast, overload

T = TypeVar("T")
Coordinate = tuple[str | int, ...]
SCHEMA_VERSION = 1


class AxisError(ValueError):
    """An axis has invalid identity, type, or keys."""


class DomainError(ValueError):
    """A domain or its supplied records are structurally invalid."""


class CoordinateError(KeyError):
    """A requested model coordinate does not exist."""


class SchemaError(ValueError):
    """A tensor violates a generated series contract."""


@dataclass(frozen=True, slots=True)
class Axis:
    """An ordered semantic dimension with explicitly typed keys."""

    name: str
    keys: tuple[str | int, ...]
    key_type: type[str] | type[int]

    def __post_init__(self) -> None:
        if not isinstance(self.name, str) or not self.name.strip():
            raise AxisError("axis name must be a nonempty string")
        if self.key_type not in (str, int):
            raise AxisError(f"axis {self.name!r}: expected str or int key type")
        keys = tuple(self.keys)
        for key in keys:
            if type(key) is not self.key_type:
                raise AxisError(f"axis {self.name!r}: key {key!r} must be {self.key_type.__name__}")
        if len(set(keys)) != len(keys):
            raise AxisError(f"axis {self.name!r}: duplicate keys")
        object.__setattr__(self, "keys", keys)


def coordinate_runs(
    axis: Axis, runs: Iterable[tuple[Coordinate, object, object]]
) -> tuple[Coordinate, ...]:
    """Expand `(prefix, first, last)` runs along `axis` into coordinates, in order.

    A ragged domain lists one run per prefix instead of every coordinate;
    each run covers the keys of `axis` from `first` through `last`.
    """
    coordinates: list[Coordinate] = []
    keys = axis.keys
    for prefix, first, last in runs:
        start = keys.index(cast(Any, first))
        stop = keys.index(cast(Any, last))
        if start > stop:
            raise DomainError(f"run {prefix!r}: {first!r} follows {last!r} on axis {axis.name!r}")
        coordinates.extend((*prefix, key) for key in keys[start : stop + 1])
    return tuple(coordinates)


@dataclass(frozen=True, slots=True)
class Domain:
    """Ordered axes and exact membership, independent of stored values."""

    axes: tuple[Axis, ...]
    coordinates: tuple[Coordinate, ...] | None = None
    _positions: tuple[Any, ...] = field(init=False, repr=False, compare=False)
    _members: Any = field(init=False, repr=False, compare=False)
    _strides: tuple[int, ...] = field(init=False, repr=False, compare=False)

    def __post_init__(self) -> None:
        axes = tuple(self.axes)
        if len({axis.name for axis in axes}) != len(axes):
            raise DomainError("domain axis names must be unique")
        object.__setattr__(self, "axes", axes)
        object.__setattr__(
            self,
            "_positions",
            tuple(MappingProxyType({key: i for i, key in enumerate(axis.keys)}) for axis in axes),
        )
        members = None
        strides: list[int] = []
        extent = 1
        for axis in reversed(axes):
            strides.append(extent)
            extent *= len(axis.keys)
        object.__setattr__(self, "_strides", tuple(reversed(strides)))
        if self.coordinates is not None:
            coords = tuple(tuple(coord) for coord in self.coordinates)
            for coord in coords:
                self._validate_components(coord)
            if len(set(coords)) != len(coords):
                raise DomainError("domain contains duplicate coordinates")
            coords = tuple(
                sorted(
                    coords,
                    key=lambda coord: tuple(
                        pos[key] for pos, key in zip(self._positions, coord, strict=True)
                    ),
                )
            )
            object.__setattr__(self, "coordinates", coords)
            # One membership index per domain, shared by every tensor over it.
            members = MappingProxyType({coord: index for index, coord in enumerate(coords)})
        object.__setattr__(self, "_members", members)

    @classmethod
    def product(cls, *axes: Axis) -> Domain:
        """Construct a Cartesian domain, with the final axis varying fastest."""
        return cls(tuple(axes))

    @classmethod
    def explicit(cls, *, axes: Iterable[Axis], coordinates: Iterable[Coordinate]) -> Domain:
        """Construct sparse membership without enumerating the Cartesian product."""
        return cls(tuple(axes), tuple(coordinates))

    def _validate_components(self, coord: Coordinate) -> None:
        if len(coord) != len(self.axes):
            raise DomainError(f"coordinate {coord!r}: expected rank {len(self.axes)}")
        for axis, positions, key in zip(self.axes, self._positions, coord, strict=True):
            if type(key) is not axis.key_type or key not in positions:
                raise DomainError(
                    f"coordinate {coord!r}: invalid key {key!r} for axis {axis.name!r}"
                )

    def require(self, coord: Coordinate) -> None:
        """Raise a coordinate error unless the complete coordinate exists."""
        self.position(coord)

    def position(self, coord: Coordinate) -> int:
        """Return the canonical position of `coord`, or raise a coordinate error.

        Product domains compute the position from axis strides; explicit
        domains look it up in the membership index they share with every
        tensor over them.
        """
        if self._members is not None:
            position = self._members.get(coord)
            if position is None:
                try:
                    self._validate_components(coord)
                except DomainError as exc:
                    raise CoordinateError(str(exc)) from exc
                raise CoordinateError(f"coordinate {coord!r} is absent from domain")
            return position
        if len(coord) != len(self.axes):
            raise CoordinateError(f"coordinate {coord!r}: expected rank {len(self.axes)}")
        position = 0
        for axis, positions, stride, key in zip(
            self.axes, self._positions, self._strides, coord, strict=True
        ):
            if type(key) is not axis.key_type or key not in positions:
                raise CoordinateError(
                    f"coordinate {coord!r}: invalid key {key!r} for axis {axis.name!r}"
                )
            position += positions[key] * stride
        return position

    def __iter__(self) -> Iterator[Coordinate]:
        return (
            iter(self.coordinates)
            if self.coordinates is not None
            else product(*(axis.keys for axis in self.axes))
        )

    def __len__(self) -> int:
        if self.coordinates is not None:
            return len(self.coordinates)
        size = 1
        for axis in self.axes:
            size *= len(axis.keys)
        return size

    def to_dict(self) -> dict[str, Any]:
        """Encode identity and order using JSON-compatible primitives."""
        return {
            "axes": [
                {"name": axis.name, "keys": list(axis.keys), "key_type": axis.key_type.__name__}
                for axis in self.axes
            ],
            "coordinates": None
            if self.coordinates is None
            else [list(coord) for coord in self.coordinates],
        }

    @property
    def fingerprint(self) -> str:
        """Deterministic content identity including axis and coordinate order."""
        return hashlib.sha256(
            json.dumps(self.to_dict(), sort_keys=True, separators=(",", ":")).encode()
        ).hexdigest()


@dataclass(frozen=True, slots=True)
class Tensor(Generic[T]):
    """An eager immutable tensor with complete domain coverage."""

    domain: Domain
    _values: tuple[T, ...]

    def __post_init__(self) -> None:
        values = tuple(self._values)
        if len(values) != len(self.domain):
            raise DomainError(f"expected {len(self.domain)} values, received {len(values)}")
        if any(
            type(value) not in (int, float, str, bool, date, datetime, type(None))
            for value in values
        ):
            raise SchemaError("tensor values must be immutable workbook scalars")
        object.__setattr__(self, "_values", values)

    @classmethod
    def _from_domain(cls, domain: Domain, values: tuple[T, ...]) -> Self:
        instance = object.__new__(cls)
        Tensor.__init__(instance, domain, values)
        return instance

    @classmethod
    def from_records(cls, *, domain: Domain, records: Iterable[tuple[Coordinate, T]]) -> Self:
        """Validate unique records and exact coverage before publication."""
        collected: dict[Coordinate, T] = {}
        for raw_coord, value in records:
            coord = tuple(raw_coord)
            try:
                domain.require(coord)
            except CoordinateError as exc:
                raise DomainError(str(exc)) from exc
            if coord in collected:
                raise DomainError(f"duplicate record coordinate {coord!r}")
            collected[coord] = value
        if len(collected) != len(domain):
            missing = next(coord for coord in domain if coord not in collected)
            raise DomainError(f"missing record coordinate {missing!r}")
        return cls._from_domain(domain, tuple(collected[coord] for coord in domain))

    @classmethod
    def from_nested(cls, *, domain: Domain, values: Any) -> Self:
        """Construct exactly shaped nested product values without padding."""
        if domain.coordinates is not None:
            raise DomainError("nested construction requires a product domain")
        flat: list[T] = []

        def visit(value: Any, depth: int) -> None:
            if depth == len(domain.axes):
                flat.append(value)
                return
            axis = domain.axes[depth]
            if (
                not isinstance(value, Sequence)
                or isinstance(value, str | bytes)
                or len(value) != len(axis.keys)
            ):
                raise DomainError(f"axis {axis.name!r}: expected {len(axis.keys)} nested values")
            for child in value:
                visit(child, depth + 1)

        visit(values, 0)
        return cls._from_domain(domain, tuple(flat))

    def __getitem__(self, key: Any) -> T:
        coord = key if isinstance(key, tuple) else (key,)
        return self._values[self.domain.position(coord)]

    def items(self) -> Iterator[tuple[Coordinate, T]]:
        """Iterate complete coordinates and values in canonical order."""
        return zip(self.domain, self._values, strict=True)

    def sel(self, **selectors: str | int) -> Tensor[T] | T:
        """Select exact named keys, dropping the fixed axes."""
        axes = self.domain.axes
        names = {axis.name for axis in axes}
        if unknown := selectors.keys() - names:
            raise CoordinateError(f"unknown selectors: {sorted(unknown)!r}")
        for axis in axes:
            if axis.name in selectors:
                key = selectors[axis.name]
                if type(key) is not axis.key_type or key not in axis.keys:
                    raise CoordinateError(f"axis {axis.name!r}: invalid key {key!r}")
        keep = [i for i, axis in enumerate(axes) if axis.name not in selectors]
        records = [
            (tuple(coord[i] for i in keep), value)
            for coord, value in self.items()
            if all(
                coord[i] == selectors[axis.name]
                for i, axis in enumerate(axes)
                if axis.name in selectors
            )
        ]
        if not records:
            raise CoordinateError(f"selection {selectors!r} matches no coordinates")
        if not keep:
            return records[0][1]
        selected_axes = tuple(
            Axis(
                axes[i].name,
                tuple(key for key in axes[i].keys if any(coord[j] == key for coord, _ in records)),
                axes[i].key_type,
            )
            for j, i in enumerate(keep)
        )
        domain = Domain.explicit(axes=selected_axes, coordinates=(coord for coord, _ in records))
        return Tensor.from_records(domain=domain, records=records)

    def isel(self, **selectors: int) -> Tensor[T] | T:
        """Select using explicit nonnegative positions along named axes."""
        axes = {axis.name: axis for axis in self.domain.axes}
        keys: dict[str, str | int] = {}
        for name, position in selectors.items():
            if (
                name not in axes
                or type(position) is not int
                or not 0 <= position < len(axes[name].keys)
            ):
                raise CoordinateError(f"axis {name!r}: invalid position {position!r}")
            keys[name] = axes[name].keys[position]
        return self.sel(**keys)

    @classmethod
    def from_legacy(
        cls, *, domain: Domain, values: Iterable[T], coordinate_order: Iterable[Coordinate]
    ) -> Self:
        """Adapt a flat buffer using an explicit, validated catalog order."""
        values, order = tuple(values), tuple(coordinate_order)
        if len(values) != len(order):
            raise DomainError("legacy values and coordinate order must have equal length")
        return cls.from_records(domain=domain, records=zip(order, values, strict=True))

    def to_legacy(self, *, coordinate_order: Iterable[Coordinate]) -> tuple[T, ...]:
        """Export values in an explicitly declared complete catalog order."""
        order = tuple(coordinate_order)
        if len(order) != len(self.domain) or len(set(order)) != len(order):
            raise DomainError("legacy coordinate order must cover the domain exactly once")
        return tuple(self[coord] for coord in order)

    def to_json(self) -> str:
        """Serialize a versioned scalar value schema and exact domain records."""
        return json.dumps(
            {
                "schema_version": SCHEMA_VERSION,
                "value_schema": "workbook_scalar",
                "domain": self.domain.to_dict(),
                "records": [
                    (
                        coord,
                        {"type": type(value).__name__, "value": value.isoformat()}
                        if isinstance(value, date)
                        else value,
                    )
                    for coord, value in self.items()
                ],
            }
        )

    @classmethod
    def from_json(cls, source: str) -> Self:
        """Restore and validate a versioned serialized tensor."""
        payload = json.loads(source)
        if (
            payload["schema_version"] != SCHEMA_VERSION
            or payload["value_schema"] != "workbook_scalar"
        ):
            raise SchemaError("unsupported tensor schema version or value schema")
        definition = payload["domain"]
        types = {"int": int, "str": str}
        axes = tuple(
            Axis(axis["name"], tuple(axis["keys"]), types[axis["key_type"]])
            for axis in definition["axes"]
        )
        coordinates = definition["coordinates"]
        domain = (
            Domain.product(*axes)
            if coordinates is None
            else Domain.explicit(axes=axes, coordinates=coordinates)
        )
        records = []
        for coord, value in payload["records"]:
            if isinstance(value, dict):
                if value.get("type") == "date":
                    value = date.fromisoformat(value["value"])
                elif value.get("type") == "datetime":
                    value = datetime.fromisoformat(value["value"])
                else:
                    raise SchemaError("unsupported serialized workbook value type")
            records.append((coord, value))
        return cls.from_records(domain=domain, records=records)


@dataclass(frozen=True, slots=True)
class TensorSchema:
    """Required series coordinates and permitted uncoerced value types."""

    series_id: str
    domain: Domain
    value_types: tuple[type, ...]

    def __post_init__(self) -> None:
        if not self.series_id or not self.domain or not self.value_types:
            raise SchemaError(
                "required series schema needs an identifier, nonempty domain, and value types"
            )
        object.__setattr__(self, "value_types", tuple(self.value_types))

    def validate(self, tensor: object, *, exact: bool = False) -> None:
        """Check semantic axes, required coordinates, and values."""
        if not isinstance(tensor, Tensor):
            raise SchemaError(f"{self.series_id}: expected Tensor, use an explicit legacy adapter")
        expected = tuple((axis.name, axis.key_type) for axis in self.domain.axes)
        actual = tuple((axis.name, axis.key_type) for axis in tensor.domain.axes)
        if actual != expected:
            raise SchemaError(f"{self.series_id}: expected axes {expected!r}, received {actual!r}")
        for coord in self.domain:
            try:
                tensor.domain.require(coord)
            except CoordinateError as exc:
                raise SchemaError(
                    f"{self.series_id}: missing required coordinate {coord!r}"
                ) from exc
        if exact and len(tensor.domain) != len(self.domain):
            raise SchemaError(f"{self.series_id}: result must exactly cover its declared domain")
        for coord, value in tensor.items():
            if type(value) not in self.value_types:
                raise SchemaError(
                    f"{self.series_id}: coordinate {coord!r} has invalid value type {type(value).__name__}"
                )


class Series(Tensor[T]):
    """A tensor validated against its generated series schema on construction.

    Generated `data` modules subclass this once per bound series and set
    `schema`, so every instance carries the axes and required coordinates of
    the authored series.
    """

    __slots__ = ()
    schema: ClassVar[TensorSchema]

    def __post_init__(self) -> None:
        super().__post_init__()
        type(self).schema.validate(self)

    @classmethod
    def collect(cls, records: Iterable[tuple[Coordinate, T]]) -> Self:
        """Publish coordinate/value records over the series' required domain."""
        return cls.from_records(domain=cls.schema.domain, records=records)


class YearSeries(Tensor[T]):
    """A schema-enforcing year facade using integer labels."""

    __slots__ = ()

    def __init__(self, *, years: Iterable[int], values: Iterable[T]) -> None:
        super().__init__(Domain.product(Axis("year", tuple(years), int)), tuple(values))

    def __post_init__(self) -> None:
        super().__post_init__()
        axes = self.domain.axes
        if len(axes) != 1 or axes[0].name != "year" or axes[0].key_type is not int:
            raise SchemaError("YearSeries requires one integer year axis")

    @property
    def years(self) -> tuple[int, ...]:
        """The ordered year labels actually present in this path."""
        return tuple(cast(int, coord[0]) for coord in self.domain)

    def __getitem__(self, key: int) -> T:
        return super().__getitem__(key)


class ScenarioSeries(Tensor[T]):
    """A schema-enforcing scenario/year facade with possibly unequal horizons."""

    __slots__ = ()

    def __post_init__(self) -> None:
        super().__post_init__()
        axes = self.domain.axes
        if (
            len(axes) != 2
            or axes[0].name != "scenario"
            or axes[0].key_type is not str
            or axes[1].name != "year"
            or axes[1].key_type is not int
        ):
            raise SchemaError("ScenarioSeries requires scenario/year axes with integer years")

    @classmethod
    def from_paths(cls, paths: Mapping[str, YearSeries[T]]) -> ScenarioSeries[T]:
        """Construct an exact ragged domain in mapping and first-occurrence order."""
        scenarios = Axis("scenario", tuple(paths), str)
        years = Axis(
            "year",
            tuple(dict.fromkeys(year for path in paths.values() for year in path.years)),
            int,
        )
        records = [
            ((scenario, year), path[year])
            for scenario, path in paths.items()
            for year in path.years
        ]
        domain = Domain.explicit(
            axes=(scenarios, years), coordinates=(coord for coord, _ in records)
        )
        return cls(
            domain,
            tuple(
                value
                for _, value in sorted(
                    records,
                    key=lambda record: (
                        scenarios.keys.index(record[0][0]),
                        years.keys.index(record[0][1]),
                    ),
                )
            ),
        )

    @overload
    def __getitem__(self, key: str) -> YearSeries[T]: ...

    @overload
    def __getitem__(self, key: tuple[str, int]) -> T: ...

    def __getitem__(self, key: str | tuple[str, int]) -> T | YearSeries[T]:
        if isinstance(key, tuple):
            return super().__getitem__(key)
        selected = self.sel(scenario=key)
        assert isinstance(selected, Tensor)
        return YearSeries[T](
            years=(cast(int, coord[0]) for coord, _ in selected.items()),
            values=(value for _, value in selected.items()),
        )
