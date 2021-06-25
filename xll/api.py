import logging

from _xlcall import lib
from .xloper12 import XLOPER12

logger = logging.getLogger(__name__)


def Excel(xlf, *args):
    res = XLOPER12()

    if logger.isEnabledFor(logging.DEBUG):
        logger.debug(f"Excel({', '.join(repr(a) for a in args)})")

    args = [XLOPER12.from_python(x) if not isinstance(x, XLOPER12) else x for x in args]

    if args:
        ret = lib.Excel12(int(xlf), res.ptr, len(args), *(x.ptr for x in args))
    else:
        ret = lib.Excel12(int(xlf), res.ptr, len(args))
    if ret:
        logger.error(f"\tExcel12 call failed with error code {ret!r}")
        raise RuntimeError(str(ret))

    if logger.isEnabledFor(logging.DEBUG):
        logger.debug(f"\t = {res!r}")

    return res
