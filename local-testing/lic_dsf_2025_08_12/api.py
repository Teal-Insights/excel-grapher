from __future__ import annotations

from .data import CONSTANTS, DEFAULT_INPUTS
from .internals import _resolve_formula
from .runtime import EvalContext, coerce_inputs_dict, xl_cell, xl_range_rows
import warnings


def make_context(inputs: dict[str, object] | None = None) -> EvalContext:
    """Create an EvalContext with merged inputs."""
    merged: dict[str, object] = dict(DEFAULT_INPUTS)
    merged.update(CONSTANTS)
    if inputs is not None:
        merged.update(inputs)
    return EvalContext(inputs=coerce_inputs_dict(merged), resolver=_resolve_formula, iterative_enabled=False, iterate_count=100, iterate_delta=0.001)


def list_setters() -> list[str]:
    """Return generated series-binding setter function names."""
    return []


def list_computes() -> list[str]:
    """Return generated series-binding compute function names."""
    return []


TARGETS = {
    "'Chart Data'!D10": xl_cell,
    "'Chart Data'!D103:X103": xl_range_rows,
    "'Chart Data'!D104:X104": xl_range_rows,
    "'Chart Data'!D105:X105": xl_range_rows,
    "'Chart Data'!D106:X106": xl_range_rows,
    "'Chart Data'!D108:X108": xl_range_rows,
    "'Chart Data'!D11": xl_cell,
    "'Chart Data'!D12": xl_cell,
    "'Chart Data'!D13": xl_cell,
    "'Chart Data'!D135:X135": xl_range_rows,
    "'Chart Data'!D14": xl_cell,
    "'Chart Data'!D145:X145": xl_range_rows,
    "'Chart Data'!D146:X146": xl_range_rows,
    "'Chart Data'!D147:X147": xl_range_rows,
    "'Chart Data'!D148:X148": xl_range_rows,
    "'Chart Data'!D15": xl_cell,
    "'Chart Data'!D150:X150": xl_range_rows,
    "'Chart Data'!D16": xl_cell,
    "'Chart Data'!D17": xl_cell,
    "'Chart Data'!D177:X177": xl_range_rows,
    "'Chart Data'!D187:X187": xl_range_rows,
    "'Chart Data'!D188:X188": xl_range_rows,
    "'Chart Data'!D189:X189": xl_range_rows,
    "'Chart Data'!D190:X190": xl_range_rows,
    "'Chart Data'!D192:X192": xl_range_rows,
    "'Chart Data'!D239:X239": xl_range_rows,
    "'Chart Data'!D240:X240": xl_range_rows,
    "'Chart Data'!D241:X241": xl_range_rows,
    "'Chart Data'!D242:X242": xl_range_rows,
    "'Chart Data'!D243:X243": xl_range_rows,
    "'Chart Data'!D244:X244": xl_range_rows,
    "'Chart Data'!D245:X245": xl_range_rows,
    "'Chart Data'!D246:X246": xl_range_rows,
    "'Chart Data'!D248:X248": xl_range_rows,
    "'Chart Data'!D249:X249": xl_range_rows,
    "'Chart Data'!D250:X250": xl_range_rows,
    "'Chart Data'!D251:X251": xl_range_rows,
    "'Chart Data'!D252:X252": xl_range_rows,
    "'Chart Data'!D263:X263": xl_range_rows,
    "'Chart Data'!D264:X264": xl_range_rows,
    "'Chart Data'!D265:X265": xl_range_rows,
    "'Chart Data'!D267:X267": xl_range_rows,
    "'Chart Data'!D281:X281": xl_range_rows,
    "'Chart Data'!D282:X282": xl_range_rows,
    "'Chart Data'!D283:X283": xl_range_rows,
    "'Chart Data'!D284:X284": xl_range_rows,
    "'Chart Data'!D285:X285": xl_range_rows,
    "'Chart Data'!D286:X286": xl_range_rows,
    "'Chart Data'!D287:X287": xl_range_rows,
    "'Chart Data'!D288:X288": xl_range_rows,
    "'Chart Data'!D290:X290": xl_range_rows,
    "'Chart Data'!D291:X291": xl_range_rows,
    "'Chart Data'!D292:X292": xl_range_rows,
    "'Chart Data'!D293:X293": xl_range_rows,
    "'Chart Data'!D294:X294": xl_range_rows,
    "'Chart Data'!D306:X306": xl_range_rows,
    "'Chart Data'!D318:X318": xl_range_rows,
    "'Chart Data'!D319:X319": xl_range_rows,
    "'Chart Data'!D320:X320": xl_range_rows,
    "'Chart Data'!D321:X321": xl_range_rows,
    "'Chart Data'!D322:X322": xl_range_rows,
    "'Chart Data'!D323:X323": xl_range_rows,
    "'Chart Data'!D324:X324": xl_range_rows,
    "'Chart Data'!D325:X325": xl_range_rows,
    "'Chart Data'!D327:X327": xl_range_rows,
    "'Chart Data'!D328:X328": xl_range_rows,
    "'Chart Data'!D329:X329": xl_range_rows,
    "'Chart Data'!D330:X330": xl_range_rows,
    "'Chart Data'!D331:X331": xl_range_rows,
    "'Chart Data'!D341:X341": xl_range_rows,
    "'Chart Data'!D342:X342": xl_range_rows,
    "'Chart Data'!D343:X343": xl_range_rows,
    "'Chart Data'!D351:X351": xl_range_rows,
    "'Chart Data'!D352:X352": xl_range_rows,
    "'Chart Data'!D353:X353": xl_range_rows,
    "'Chart Data'!D354:X354": xl_range_rows,
    "'Chart Data'!D355:X355": xl_range_rows,
    "'Chart Data'!D356:X356": xl_range_rows,
    "'Chart Data'!D357:X357": xl_range_rows,
    "'Chart Data'!D358:X358": xl_range_rows,
    "'Chart Data'!D360:X360": xl_range_rows,
    "'Chart Data'!D361:X361": xl_range_rows,
    "'Chart Data'!D362:X362": xl_range_rows,
    "'Chart Data'!D363:X363": xl_range_rows,
    "'Chart Data'!D364:X364": xl_range_rows,
    "'Chart Data'!D51:X51": xl_range_rows,
    "'Chart Data'!D61:X61": xl_range_rows,
    "'Chart Data'!D62:X62": xl_range_rows,
    "'Chart Data'!D63:X63": xl_range_rows,
    "'Chart Data'!D64:X64": xl_range_rows,
    "'Chart Data'!D66:X66": xl_range_rows,
    "'Chart Data'!D93:X93": xl_range_rows,
    "'Chart Data'!E25": xl_cell,
    "'Chart Data'!E26": xl_cell,
    "'Chart Data'!E27": xl_cell,
    "'Chart Data'!I10": xl_cell,
    "'Chart Data'!I11": xl_cell,
    "'Chart Data'!I12": xl_cell,
    "'Chart Data'!I13": xl_cell,
    "'Chart Data'!I14": xl_cell,
    "'Chart Data'!I17": xl_cell,
    "'Chart Data'!I18": xl_cell,
    "'Chart Data'!I19": xl_cell,
    "'Chart Data'!L10": xl_cell,
    "'Chart Data'!L11": xl_cell,
}


def compute_all(ctx: EvalContext | None = None, *, inputs: dict[str, object] | None = None) -> dict[str, object]:
    """Compute all target cells and return results."""
    if ctx is None:
        ctx = make_context(inputs)
    elif inputs is not None:
        warnings.warn(
            "inputs will be ignored because ctx was provided",
            UserWarning,
            stacklevel=2,
        )
    return {target: handler(ctx, target) for target, handler in TARGETS.items()}
