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

    # how to create  LPXLOPER12 owning an array? Can we even do that...
    # can we just hack GC here.

    xlo = ffi.new("XLOPER12[2]")

    xlo[1].xltype = 0

    if value is None:
        xlo[0].xltype = lib.xltypeBool
    elif type(value) is bool:
        xlo[0].xltype = lib.xltypeBool
        xlo[0].val.xbool = 1 if value else 0
    elif type(value) is int:
        xlo[0].xltype = lib.xltypeInt
        xlo[0].val.w = value
    elif type(value) is float:
        xlo[0].xltype = lib.xltypeNum
        xlo[0].val.num = value
    elif type(value) is str:
        xlo[0].xltype = lib.xltypeStr

        xlo[0].val.str = lib.malloc((len(value) + 1) * 2)
        xlo[0].val.str[0] = chr(len(value))
        xlo[0].val.str[1 : len(value) + 1] = value

        def _free(xlo):
            logger.debug(
                f"freeing xltypeStr {ffi.string(xlo[0].val.str[1:1+ord(xlo[0].val.str[0])])}"
            )
            lib.free(xlo[0].val.str)
            xlo[0].xltype = 0

        xlo = ffi.gc(xlo, _free)

    else:
        raise TypeError(f"cannot convert {type(value)!r} to XLOPER12")

    return xlo


def to_xloper_result(value):
    # do the normal conversion first
    xlo = to_xloper(value)

    logger.info("converting to result")

    # signal to excel that it needs to call xlAutoFree12, and setup the pointer
    # in the extra xloper
    xlo[0].xltype |= lib.xlbitDLLFree
    xlo[1].xltype = 0xFFFF
    xlo[1].val.str = ffi.cast("wchar_t*", id(xlo))

    # hack to get the refcount from the object and increment it for excel's reference
    prefcnt = ffi.cast("unsigned int*", id(xlo))
    prefcnt[0] = prefcnt[0] + 1

    logger.info("done converting to result")

    # convert to pointer to return to excel
    return ffi.cast("LPXLOPER12", xlo)
