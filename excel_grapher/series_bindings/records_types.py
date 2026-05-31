"""Shared record type aliases for generated binding APIs."""

from __future__ import annotations

Record = dict[str, object]
Records = list[Record]
Scalar = str | int | float | bool | None

__all__ = ["Record", "Records", "Scalar"]
