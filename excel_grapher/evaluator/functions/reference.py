"""Register Excel reference functions against the shared runtime.

ROW / COLUMN / COLUMNS / OFFSET are special-cased in
`excel_grapher.evaluator.evaluator` because they need range context; they
are not dispatched through `FUNCTIONS` and therefore not registered here.
"""

from __future__ import annotations

from excel_grapher.runtime.reference import xl_address

from . import register

register("ADDRESS")(xl_address)
