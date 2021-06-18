import logging

from _xlcall import ffi, lib
from .convert import to_xloper, from_xloper

logger = logging.getLogger(__name__)


def Excel(xlf, *args, convert=True):
    res = ffi.new("LPXLOPER12")

    args = list(map(to_xloper, args))

    if args:
        ret = lib.Excel12(int(xlf), res, len(args), *args)
    else:
        ret = lib.Excel12(int(xlf), res, len(args))

    if ret:
        logger.debug(f"\tCall Failed {ret!r}")
        raise RuntimeError(str(ret))

    if not convert:
        return res

    return from_xloper(res)
