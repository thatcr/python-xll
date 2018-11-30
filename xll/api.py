from _xlcall import ffi, lib
from .convert import to_xloper, from_xloper

print(__file__)

def Excel(xlf, *args, convert=True):
    res = ffi.new('LPXLOPER12')

    args = list(map(to_xloper, args))

    if args:
        ret = lib.Excel12(int(xlf), res, len(args), *args)
    else:
        ret = lib.Excel12(int(xlf), res, len(args))

    if ret:
        raise RuntimeError(str(ret))

    if not convert:
        return res

    return from_xloper(res)

