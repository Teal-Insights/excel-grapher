"""Immutable named-coordinate values shared by generated model functions."""

from __future__ import annotations

import hashlib
import json
from collections.abc import Iterable, Iterator, Mapping, Sequence
from dataclasses import dataclass, field
from datetime import date, datetime
from itertools import product
from types import MappingProxyType
from typing import TYPE_CHECKING, Any, Generic, Self, TypeVar, cast, overload

if TYPE_CHECKING:
    from .provenance import ProvenanceTemplate

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
    _index: Mapping[str | int, int] = field(init=False, repr=False, compare=False)

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
        object.__setattr__(self, "_index", MappingProxyType({key: i for i, key in enumerate(keys)}))

    def __contains__(self, key: object) -> bool:
        """True when `key` is a typed member of this axis."""
        return type(key) is self.key_type and key in self._index

    def position(self, key: str | int) -> int:
        """Return the 0-based index of `key` on this axis."""
        if key not in self:
            raise CoordinateError(f"axis {self.name!r}: invalid key {key!r}")
        return self._index[key]


def label_axis(name: str, values: Iterable[object], key_type: type[str] | type[int]) -> Axis:
    """Build an axis from evaluated labels, naming duplicate or mistyped values."""
    keys: list[str | int] = []
    seen: set[object] = set()
    for value in values:
        if type(value) is not key_type:
            raise AxisError(f"axis {name!r}: label {value!r} must be {key_type.__name__}")
        if value in seen:
            raise AxisError(f"axis {name!r}: duplicate label {value!r}")
        seen.add(value)
        keys.append(cast(str | int, value))
    return Axis(name, tuple(keys), key_type)


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

    def sparse_get(self, coord: Coordinate) -> int | None:
        """Return the position, or `None` when a sparse domain lacks `coord`.

        Invalid axis keys still raise `CoordinateError`. A missing membership
        in an explicit domain is a blank, not a bad key.
        """
        if self._members is None:
            return self.position(coord)
        position = self._members.get(coord)
        if position is not None:
            return position
        try:
            self._validate_components(coord)
        except DomainError as exc:
            raise CoordinateError(str(exc)) from exc
        return None

    def position(self, coord: Coordinate) -> int:
        """Return the canonical position of `coord`, or raise a coordinate error.

        Product domains compute the position from axis strides; explicit
        domains look it up in the membership index they share with every
        tensor over them.
        """
        if self._members is not None:
            position = self.sparse_get(coord)
            if position is None:
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
class AxisTemplate:
    """An axis whose keys are known only after its labeller is evaluated."""

    name: str
    key_type: type[str] | type[int]
    size: int
    labeller: str
    snapshot: tuple[str | int, ...]
    source: tuple[str | int, ...] | None = None

    def bind(self, axis: Axis) -> Axis:
        """Return `axis`, or the snapshot-aligned slice of a longer labeller axis.

        `axis` may already be the subset (`len(keys) == size`) or the full
        labeller (`len(keys) == len(source)`). Other lengths fail closed.
        """
        if axis.name != self.name:
            raise AxisError(
                f"axis {self.name!r}: expected name {self.name!r}, received {axis.name!r}"
            )
        if axis.key_type is not self.key_type:
            raise AxisError(
                f"axis {self.name!r}: expected {self.key_type.__name__} keys, "
                f"received {axis.key_type.__name__}"
            )
        source = self.source if self.source is not None else self.snapshot
        n_keys = len(axis.keys)
        if n_keys == self.size:
            return axis
        if n_keys == len(source) and source != self.snapshot:
            keys = tuple(axis.keys[source.index(key)] for key in self.snapshot)
            return Axis(self.name, keys, self.key_type)
        raise AxisError(f"axis {self.name!r}: expected {self.size} keys, received {n_keys}")


def _bound_axis(axis: Axis | AxisTemplate, labellers: Mapping[str, Tensor[Any]]) -> Axis:
    if isinstance(axis, Axis):
        return axis
    labeller = labellers.get(axis.name)
    if labeller is None:
        raise AxisError(f"axis {axis.name!r}: missing labeller {axis.labeller!r}")
    if not labeller.domain.axes:
        raise AxisError(f"axis {axis.name!r}: labeller {axis.labeller!r} has no axes")
    return axis.bind(labeller.domain.axes[0])


@dataclass(frozen=True, slots=True)
class DomainTemplate:
    """A domain with one or more axes supplied by labellers at call time."""

    axes: tuple[Axis | AxisTemplate, ...]
    coordinates: tuple[tuple[int, ...], ...] | None = None

    @classmethod
    def product(cls, *axes: Axis | AxisTemplate) -> DomainTemplate:
        """Construct a Cartesian template, with the final axis varying fastest."""
        return cls(tuple(axes))

    @classmethod
    def explicit(
        cls, *, axes: Iterable[Axis | AxisTemplate], coordinates: Iterable[tuple[int, ...]]
    ) -> DomainTemplate:
        """Store sparse membership as positions along each template axis."""
        return cls(tuple(axes), tuple(tuple(coord) for coord in coordinates))

    def bind(self, **labellers: Tensor[Any]) -> Domain:
        """Materialize axes from `labellers` keyed by axis name."""
        axes = tuple(_bound_axis(axis, labellers) for axis in self.axes)
        return self._domain(axes)

    def domain_from_axes(self, axes: Sequence[Axis]) -> Domain:
        """Materialize this template over already-bound `axes`."""
        if len(axes) != len(self.axes):
            raise DomainError(f"expected {len(self.axes)} axes, received {len(axes)}")
        return self._domain(tuple(axes))

    def _domain(self, axes: tuple[Axis, ...]) -> Domain:
        if self.coordinates is None:
            return Domain.product(*axes)
        coordinates = tuple(
            tuple(axes[index].keys[position] for index, position in enumerate(coord))
            for coord in self.coordinates
        )
        return Domain.explicit(axes=axes, coordinates=coordinates)

    def __iter__(self) -> Iterator[Coordinate]:
        raise TypeError("bind a DomainTemplate before iterating its coordinates")


@dataclass(frozen=True, slots=True)
class SchemaTemplate:
    """A series contract whose runtime axes bind from labellers or the tensor itself."""

    series_id: str
    domain: DomainTemplate
    value_types: tuple[type, ...]

    def __post_init__(self) -> None:
        if not self.series_id or not self.domain.axes or not self.value_types:
            raise SchemaError(
                "required series schema needs an identifier, nonempty domain, and value types"
            )
        object.__setattr__(self, "value_types", tuple(self.value_types))

    def bind(self, **labellers: Tensor[Any]) -> TensorSchema:
        """Bind runtime axes and return a concrete `TensorSchema`."""
        return TensorSchema(self.series_id, self.domain.bind(**labellers), self.value_types)

    def validate(self, tensor: object, *, exact: bool = False) -> None:
        """Check size and type on runtime axes, resolving keys from `tensor`."""
        if not isinstance(tensor, Tensor):
            raise SchemaError(f"{self.series_id}: expected Tensor")
        if len(tensor.domain.axes) != len(self.domain.axes):
            expected = tuple(axis.name for axis in self.domain.axes)
            actual = tuple(axis.name for axis in tensor.domain.axes)
            raise SchemaError(f"{self.series_id}: expected axes {expected!r}, received {actual!r}")
        bound: list[Axis] = []
        for template, actual in zip(self.domain.axes, tensor.domain.axes, strict=True):
            if isinstance(template, AxisTemplate):
                try:
                    bound.append(template.bind(actual))
                except AxisError as exc:
                    raise SchemaError(f"{self.series_id}: {exc}") from exc
                continue
            if actual.name != template.name or actual.key_type is not template.key_type:
                expected = tuple((axis.name, axis.key_type) for axis in self.domain.axes)
                received = tuple((axis.name, axis.key_type) for axis in tensor.domain.axes)
                raise SchemaError(
                    f"{self.series_id}: expected axes {expected!r}, received {received!r}"
                )
            bound.append(template)
        TensorSchema(
            self.series_id, self.domain.domain_from_axes(bound), self.value_types
        ).validate(tensor, exact=exact)


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
        """Return the value at `key`, or `None` for a valid sparse hole.

        Invalid axis keys still raise `CoordinateError`. A coordinate whose
        components are on the axes but missing from an explicit domain is a
        blank, matching Excel's empty triangle cells.
        """
        coord = key if isinstance(key, tuple) else (key,)
        position = self.domain.sparse_get(coord)
        if position is None:
            return cast(T, None)
        return self._values[position]

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
        """Check semantic axes, required coordinates, and values.

        A same-rank tensor whose keys are neither a subset nor a superset of
        the required keys raises `SchemaError` naming the unknown labels.
        That message is not labelled-axis-specific.
        """
        if not isinstance(tensor, Tensor):
            raise SchemaError(f"{self.series_id}: expected Tensor")
        expected = tuple((axis.name, axis.key_type) for axis in self.domain.axes)
        actual = tuple((axis.name, axis.key_type) for axis in tensor.domain.axes)
        if actual != expected:
            raise SchemaError(f"{self.series_id}: expected axes {expected!r}, received {actual!r}")
        for expected_axis, actual_axis in zip(self.domain.axes, tensor.domain.axes, strict=True):
            expected_keys = set(expected_axis.keys)
            unknown = tuple(key for key in actual_axis.keys if key not in expected_keys)
            missing = tuple(key for key in expected_axis.keys if key not in set(actual_axis.keys))
            if unknown and missing:
                raise SchemaError(
                    f"{self.series_id}: axis {expected_axis.name!r} has unknown labels "
                    f"{unknown!r}; accepted labels are {expected_axis.keys!r}"
                )
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


@dataclass(frozen=True, slots=True)
class Series(Tensor[T]):
    """A tensor bound to a series schema and optional cell provenance.

    `define_series` constructs instances for inputs and constants. Formula
    series without observations use `SeriesSpec` with the same attributes.
    """

    schema: TensorSchema | SchemaTemplate = field(kw_only=True, repr=False, compare=False)
    cells: Mapping[Coordinate, str] | ProvenanceTemplate | None = field(
        default=None, kw_only=True, repr=False, compare=False
    )

    def __post_init__(self) -> None:
        # Zero-arg super() closes over the pre-slots class dataclass replaces.
        Tensor.__post_init__(self)
        self.schema.validate(self)

    @property
    def required(self) -> Domain | DomainTemplate:
        """The schema's required coordinates, which may be a subset of `domain`."""
        return self.schema.domain

    def collect(
        self, records: Iterable[tuple[Coordinate, T]], domain: Domain | None = None
    ) -> Self:
        """Publish coordinate/value records over the series' required domain."""
        return cast(Self, _collect_series(type(self), self.schema, self.cells, records, domain))

    @classmethod
    def from_labels(cls, source: SeriesSpec[T] | Series[T], axis: Axis) -> Series[T]:
        """Publish the identity tensor mapping each label to itself.

        `source` supplies the series id, value types, and cell provenance.
        """
        domain = Domain.product(axis)
        schema = TensorSchema(source.schema.series_id, domain, source.schema.value_types)
        return cls(domain, cast(tuple[T, ...], tuple(axis.keys)), schema=schema, cells=source.cells)

    def with_values(self, values: Sequence[T]) -> Self:
        """Return a series over the same domain, schema, and cells."""
        return type(self)(self.domain, tuple(values), schema=self.schema, cells=self.cells)

    def with_nested(self, values: Any) -> Self:
        """Return a series from nested product values over the authored domain."""
        tensor = Tensor.from_nested(domain=self.domain, values=values)
        return type(self)(tensor.domain, tensor._values, schema=self.schema, cells=self.cells)

    def with_records(self, records: Iterable[tuple[Coordinate, T]]) -> Self:
        """Return a series from records over the authored domain."""
        tensor = Tensor.from_records(domain=self.domain, records=records)
        return type(self)(tensor.domain, tensor._values, schema=self.schema, cells=self.cells)


@dataclass(frozen=True, slots=True)
class SeriesSpec(Generic[T]):
    """Schema, authored domain, and cell provenance without observations."""

    schema: TensorSchema | SchemaTemplate
    domain: Domain | DomainTemplate
    cells: Mapping[Coordinate, str] | ProvenanceTemplate | None

    @property
    def required(self) -> Domain | DomainTemplate:
        """The schema's required coordinates, which may be a subset of `domain`."""
        return self.schema.domain

    def collect(
        self, records: Iterable[tuple[Coordinate, T]], domain: Domain | None = None
    ) -> Series[T]:
        """Publish coordinate/value records over the series' required domain."""
        return _collect_series(Series, self.schema, self.cells, records, domain)

    def with_values(self, values: Sequence[T]) -> Series[T]:
        """Bind observations over the authored domain."""
        if not isinstance(self.domain, Domain):
            raise TypeError("with_values requires a bound domain")
        return Series(self.domain, tuple(values), schema=self.schema, cells=self.cells)

    def with_nested(self, values: Any) -> Series[T]:
        """Bind nested product values over the authored domain."""
        if not isinstance(self.domain, Domain):
            raise TypeError("with_nested requires a bound domain")
        tensor = Tensor.from_nested(domain=self.domain, values=values)
        return Series(tensor.domain, tensor._values, schema=self.schema, cells=self.cells)

    def with_records(self, records: Iterable[tuple[Coordinate, T]]) -> Series[T]:
        """Bind records over the authored domain."""
        if not isinstance(self.domain, Domain):
            raise TypeError("with_records requires a bound domain")
        tensor = Tensor.from_records(domain=self.domain, records=records)
        return Series(tensor.domain, tensor._values, schema=self.schema, cells=self.cells)


def _collect_series(
    series_type: type[Series[T]],
    schema: TensorSchema | SchemaTemplate,
    cells: Mapping[Coordinate, str] | ProvenanceTemplate | None,
    records: Iterable[tuple[Coordinate, T]],
    domain: Domain | None,
) -> Series[T]:
    schema_domain = schema.domain
    if domain is None:
        if not isinstance(schema_domain, Domain):
            raise TypeError("collect requires a bound domain")
        domain = schema_domain
    tensor = Tensor.from_records(domain=domain, records=records)
    bound_schema: TensorSchema | SchemaTemplate = schema
    if isinstance(schema, SchemaTemplate):
        bound_schema = TensorSchema(schema.series_id, domain, schema.value_types)
    return series_type(tensor.domain, tensor._values, schema=bound_schema, cells=cells)


@overload
def define_series(
    series_id: str,
    domain: Domain,
    values: Sequence[T],
    *,
    cells: Mapping[Coordinate, str] | ProvenanceTemplate | None,
    value_types: tuple[type, ...],
    required: Domain | DomainTemplate | None = None,
) -> Series[T]: ...


@overload
def define_series(
    series_id: str,
    domain: Domain | DomainTemplate,
    values: None = None,
    *,
    cells: Mapping[Coordinate, str] | ProvenanceTemplate | None,
    value_types: tuple[type, ...],
    required: Domain | DomainTemplate | None = None,
) -> SeriesSpec[T]: ...


def define_series(
    series_id: str,
    domain: Domain | DomainTemplate,
    values: Sequence[T] | None = None,
    *,
    cells: Mapping[Coordinate, str] | ProvenanceTemplate | None,
    value_types: tuple[type, ...],
    required: Domain | DomainTemplate | None = None,
) -> Series[T] | SeriesSpec[T]:
    """Bind a series schema, provenance, and optional default observations.

    `required` defaults to `domain`. Pass it only when the required
    coordinates are a proper subset of the authored domain.
    """
    schema_domain = domain if required is None else required
    schema: TensorSchema | SchemaTemplate
    if isinstance(schema_domain, DomainTemplate):
        schema = SchemaTemplate(series_id, schema_domain, value_types)
    else:
        schema = TensorSchema(series_id, schema_domain, value_types)
    if values is None:
        return SeriesSpec(schema=schema, domain=domain, cells=cells)
    if not isinstance(domain, Domain):
        raise TypeError("define_series with values requires a bound Domain")
    return Series(domain, tuple(values), schema=schema, cells=cells)
