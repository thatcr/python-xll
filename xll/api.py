import logging

from _xlcall import lib
from .xloper12 import XLOPER12

logger = logging.getLogger(__name__)


def Excel(xlf, *args):
    res = XLOPER12()

    args = [XLOPER12.from_python(x) for x in args]

    if args:
        ret = lib.Excel12(int(xlf), res.ptr, len(args), [x.ptr for x in args])
    else:
        ret = lib.Excel12(int(xlf), res.ptr, len(args))

    if ret:
        logger.debug(f"\tCall Failed {ret!r}")
        raise RuntimeError(str(ret))

    return res
