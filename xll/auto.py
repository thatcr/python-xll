import xll
import sys
import logging

from xll.api import Excel

from _python_xll import ffi as addin
import _xlthunk

from _xlcall import lib, ffi

from .convert import to_xloper

logger = logging.getLogger(__name__)


# allocate xlopers to return as error values

xlerrNull = ffi.new("LPXLOPER12", {"xltype": lib.xltypeErr})
xlerrNull.val.err = lib.xlerrNull

xlerrValue = ffi.new("LPXLOPER12", {"xltype": lib.xltypeErr})
xlerrValue.val.err = lib.xlerrValue


def onerror(exception, exc_value, traceback):
    sys.excepthook(exception, exc_value, traceback)

    # caller = Excel(lib.xlfCaller, convert=False)
    # text = Excel(lib.xlSheetNm, caller)

    # logger.error(
    #     f"{ffi.string(text.val.str)}", exc_info=(exception, exc_value, traceback)
    # )
    return xlerrValue


@ffi.callback("LPXLOPER12 (*)()", onerror=onerror)
def xlfCaller():
    logger.info("Called xlfCaller")

    caller = Excel(lib.xlfCaller, convert=False)

    text = Excel(lib.xlSheetNm, caller)

    logger.info(f"{ffi.string(text.val.str)}")

    result = to_xloper(text)
    result.xltype |= lib.xlbitDLLFree

    # use the sneaky incref hack here to fix the pyobject?

    # how is result kept as a return value?

    return result


_cache = []

# TODO wrap the xloper in a pointer, and use an offset/cast to include it.
# so PyXLOPER as a C type. That way we can ref-count in sync with excel.


@addin.callback("LPXLOPER12 (*)(const char*)", onerror=onerror)
def py_eval(source):
    logger.info(f"Called py.eval")

    value = sys.version

    # result = to_xloper(eval(ffi.string(source), locals(), {}))
    # using  crashes!!! hardocding the hex works.

    def _free(x):
        logger.info("freeing xloper!")
        lib.free(x.xlo.val.str)

    result = ffi.new("struct PyXLOPER12*")
    result.xlo.xltype = lib.xltypeStr | lib.xlbitDLLFree  # 0x4002

    result.xlo.val.str = lib.malloc((len(value) + 1) * 2)
    result.xlo.val.str[0] = chr(len(value))
    result.xlo.val.str[1 : len(value) + 1] = value
    result = ffi.gc(result, _free)

    result.ptr = ffi.cast("void*", id(result))
    logger.info(f"PyXLOPER12 @ {result.ptr}, { hex(id(result))}")

    pref = ffi.cast("unsigned int*", id(result))

    pref[0] = pref[0] + 1

    logger.info(f"pref = {pref[0]}, sys.getrefcount = {sys.getrefcount(result)}")

    # _cache.append(result)

    return ffi.addressof(result.xlo)


@addin.def_extern(error=0)
def xlAutoOpen():
    logger.info("xlAutoOpen!")
    name = xll.Excel(lib.xlGetName)

    logger.info(f"Python XLL Loaded from {name}")

    # import debugpy
    # debugpy.listen(5678)

    _xlthunk.lib.SetThunkProc(b"_0000", py_eval)
    xll.Excel(
        lib.xlfRegister,
        _xlthunk.__file__,
        "_0000",
        "QC",
        "PY.EVAL",
        "source",
        1,
        "Python Builtins",
        None,
        None,
        "Evaluate a python expression",
        "Python expression to evaluate",
    )

    # exports['_001'] = xlfCaller
    _xlthunk.lib.SetThunkProc(b"_0001", xlfCaller)
    xll.Excel(
        lib.xlfRegister,
        _xlthunk.__file__,
        "_0001",
        "Q",
        "xlfCaller",
        "",
        1,
        "Python Builtins",
        None,
        None,
        None,
        None,
    )
    return 1


@addin.def_extern(error=0)
def xlAutoClose():
    logger.info("xlAutoClose")
    return 1


@addin.def_extern(error=0)
def xlAutoAdd():
    logger.info("xlAutoAdd")
    return 1


@addin.def_extern(error=0)
def xlAutoRemove():
    logger.info("xlAutoRemove")
    return 1


@addin.def_extern(error=xlerrNull)
def xlAutoRegister12(xloper):
    # cleanup memory from strings
    logger.debug(f"xlAutoFree12({xloper!r}")
    raise NotImplementedError


@addin.def_extern(error=xlerrNull)
def xlAddInManagerInfo12(xloper):
    # cleanup memory from strings
    logger.debug(f"xlAddInManagerInfo12({xloper!r}")

    result = to_xloper(sys.version)
    result.xltype |= lib.xlbitDLLFree
    return result
