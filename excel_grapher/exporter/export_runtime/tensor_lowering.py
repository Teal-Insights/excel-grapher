"""Private positional view for scheduled kernels over immutable tensor inputs."""

from collections.abc import Sequence
from dataclasses import dataclass
from typing import Generic, TypeVar, overload

from .tensor import Coordinate, Tensor

T = TypeVar("T")


@dataclass(frozen=True, slots=True)
class CoordinateBuffer(Sequence[T], Generic[T]):
    """Resolve a private kernel slot against its explicit semantic coordinate."""

    tensor: Tensor[T]
    coordinate_order: tuple[Coordinate, ...]

    def __len__(self) -> int:
        return len(self.coordinate_order)

    @overload
    def __getitem__(self, index: int) -> T: ...

    @overload
    def __getitem__(self, index: slice) -> tuple[T, ...]: ...

    def __getitem__(self, index: int | slice) -> T | tuple[T, ...]:
        if isinstance(index, slice):
            return tuple(self.tensor[coord] for coord in self.coordinate_order[index])
        return self.tensor[self.coordinate_order[index]]
