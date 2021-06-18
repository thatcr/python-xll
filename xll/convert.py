import logging
from _xlcall import ffi, lib

logger = logging.getLogger(__file__)


def from_xloper(xloper):
    xltype = 0x0FFF & xloper.xltype
    if xltype == lib.xltypeErr:
        raise ValueError(xloper.val.err)
    if xltype == lib.xltypeNil:
        return None
    if xltype == lib.xltypeBool:
        return bool(xloper.val.xbool)
    if xltype == lib.xltypeNum:
        return xloper.val.num
    if xltype == lib.xltypeInt:
        return xloper.val.w
    if xltype == lib.xltypeStr:
        return ffi.string(xloper.val.str + 1, ord(xloper.val.str[0]))
    raise RuntimeError(f"unknown xloper type {xloper.xltype:d}")


def to_xloper(value):
    if isinstance(value, ffi.CData):
        return value

    # initialize the result, and
    result = ffi.new("LPPYXLOPER12")
    result.ptr = ffi.cast("void*", 0)

    xlo = result.xlo

    if value is None:
        xlo.xltype = lib.xltypeBool
        return result
    elif type(value) is bool:
        xlo.xltype = lib.xltypeBool
        xlo.val.xbool = 1 if value else 0
    elif type(value) is int:
        xlo.xltype = lib.xltypeInt
        xlo.val.w = value
    elif type(value) is float:
        xlo.xltype = lib.xltypeNum
        xlo.val.num = value
    elif type(value) is str:
        xlo.xltype = lib.xltypeStr

        xlo.val.str = lib.malloc((len(value) + 1) * 2)
        xlo.val.str[0] = chr(len(value))
        xlo.val.str[1 : len(value) + 1] = value

        def _free(x):
            logger.debug(
                f"freeing xltypeStr {ffi.string(x.xlo.val.str[1:1+ord(x.xlo.val.str[0])])}"
            )
            lib.free(x.xlo.val.str)

        result = ffi.gc(result, _free)

    else:
        raise TypeError(f"cannot convert {type(value)!r} to XLOPER12")

    return result


def to_xloper_result(value):
    # do the normal conversion first
    result = to_xloper(value)

    logger.info("converting to result")

    # include the object pointer in the data after the XLOPER
    result.ptr = ffi.cast("void*", id(result))

    # signal to excel that it needs to call xl
    result.xlo.xltype |= lib.xlbitDLLFree

    # hack to get the refcount from the object and increment it for excel's reference
    prefcnt = ffi.cast("unsigned int*", id(result))
    prefcnt[0] = prefcnt[0] + 1

    logger.info("done converting to result")

    return result
