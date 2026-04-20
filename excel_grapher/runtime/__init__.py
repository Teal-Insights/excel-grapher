"""Shared Excel runtime.

Houses the implementations used both by :mod:`excel_grapher.evaluator` at eval time
and by :mod:`excel_grapher.exporter` when embedding the runtime into generated
standalone code (see :func:`excel_grapher.exporter.embed.emit_runtime`). This
package must not import from ``evaluator``, ``exporter``, or ``grapher``.
"""

from __future__ import annotations

__all__: list[str] = []
