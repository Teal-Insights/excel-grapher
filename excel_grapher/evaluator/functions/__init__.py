from __future__ import annotations

from collections.abc import Callable
from typing import ParamSpec, TypeVar

from ..types import CellValue

FUNCTIONS: dict[str, Callable[..., CellValue]] = {}

P = ParamSpec("P")
R = TypeVar("R", bound=CellValue)


def register(
    name: str,
) -> Callable[[Callable[P, R]], Callable[P, R]]:
    def decorator(fn: Callable[P, R]) -> Callable[P, R]:
        FUNCTIONS[name.upper()] = fn
        return fn

    return decorator


# Import modules for side-effect registration.
from . import info as _info  # noqa: E402,F401
from . import logic as _logic  # noqa: E402,F401
from . import lookup as _lookup  # noqa: E402,F401
from . import math as _math  # noqa: E402,F401
from . import reference as _reference  # noqa: E402,F401
from . import text as _text  # noqa: E402,F401
