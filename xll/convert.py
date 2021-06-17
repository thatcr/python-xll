import logging
from _python_xll import ffi, lib

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

    if value is None:
        result = ffi.new("LPXLOPER12")
        result.xltype = lib.xltypeBool
        return result

    if type(value) is bool:
        result = ffi.new("LPXLOPER12")
        result.xltype = lib.xltypeBool
        result.val.xbool = 1 if value else 0
        return result

    if type(value) is int:
        result = ffi.new("LPXLOPER12")
        result.xltype = lib.xltypeInt
        result.val.w = value
        return result

    if type(value) is float:
        result = ffi.new("LPXLOPER12")
        result.xltype = lib.xltypeNum
        result.val.num = value
        return result

    if type(value) is str:
        result = ffi.new("LPXLOPER12")
        result.xltype = lib.xltypeStr | lib.xlbitDLLFree

        # TODO expose an incref here? nope. can't go back to the CDATA
        #      use malloc/free? or store a map of cdata -> xloper and back.
        #
        result.val.str = lib.malloc((len(value) + 1) * 2)
        result.val.str[0] = chr(len(value))
        result.val.str[1 : len(value) + 1] = value

        return result

    raise TypeError(f"cannot convert {type(value)!r} to XLOPER12")


def xlAutoFree12(xloper):
    # cleanup memory from strings
    logger.debug(f"xlAutoFree12({xloper!r}")
    if xloper.xltype == lib.xltypeStr | lib.xlbitDLLFree:
        lib.free(xloper.val.str)
        xloper.xltype = 0
