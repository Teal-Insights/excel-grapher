"""Export-owned runtime primitives for generated Python code.

The lazy range value is named `Range`. It stores rectangular worksheet geometry
and a resolver callable with the shape `resolver(address: str) -> CellValue`.
"""

from .errors import XlErrorException
from .ranges import Range

__all__ = ["Range", "XlErrorException"]
