import xll
import sys
import logging

from xll.api import Excel

from _python_xll import ffi as addin
import _xlthunk

from _xlcall import lib, ffi
from .convert import to_xloper, to_xloper_result

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

    return result


_cache = []

# TODO wrap the xloper in a pointer, and use an offset/cast to include it.
# so PyXLOPER as a C type. That way we can ref-count in sync with excel.


@addin.callback("LPXLOPER12 (*)(const char*)", onerror=onerror)
def py_eval(source):
    logger.info(f"Called py.eval with {ffi.string(source)!r}")

    # execute the python and get the result
    value = eval(ffi.string(source), locals(), {"sys": sys})

    result = to_xloper_result(value)

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
    # this will invoke xlAutoFree on this addin, not the thunk...
    logger.debug(f"xlAddInManagerInfo12({xloper!r}")

    result = to_xloper_result(sys.version)
    return result
