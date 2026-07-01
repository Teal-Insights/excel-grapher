"""Excel error exceptions for the exported Python runtime."""

from __future__ import annotations

from excel_grapher.core import XlError

__all__ = ["XlErrorException"]


class XlErrorException(Exception):
    """Exception form of an Excel error code in exported Python code."""

    code: XlError

    def __init__(self, code: XlError) -> None:
        """Initialize the exception with an Excel error code."""
        if not isinstance(code, XlError):
            raise TypeError(f"Expected XlError, got {type(code).__name__}")
        self.code = code
        super().__init__(code.value)
