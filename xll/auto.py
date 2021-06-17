import xll
import sys
import logging

from xll.api import Excel

from _python_xll import ffi, lib
import _xlthunk

from .convert import to_xloper

logger = logging.getLogger(__name__)


# allocate xlopers to return as error values

xlerrNull = ffi.new("LPXLOPER12", {"xltype": lib.xltypeErr})
xlerrNull.val.err = lib.xlerrNull

xlerrValue = ffi.new("LPXLOPER12", {"xltype": lib.xltypeErr})
xlerrValue.val.err = lib.xlerrValue


def onerror(exception, exc_value, traceback):
    sys.excepthook(exception, exc_value, traceback)

    logger.error("Function call failed!", exc_info=(exception, exc_value, traceback))
    return xlerrValue


@ffi.callback("LPXLOPER12 (*)()", onerror=onerror)
def xlfCaller():
    logger.info("Called xlfCaller")
    caller = Excel(lib.xlfCaller, convert=False)

    text = Excel(lib.xlSheetNm, caller)

    result = to_xloper(text)
    result.xltype |= lib.xlbitDLLFree

    # how is result kept as a return value?

    return result


@ffi.callback("LPXLOPER12 (*)(const char*)", onerror=onerror)
def py_eval(source):
    logging.info(f"Called py.eval with {source}")
    result = to_xloper(eval(ffi.string(source), locals(), {}))
    result.xltype |= lib.xlbitDLLFree
    return result


@ffi.def_extern(error=0)
def xlAutoOpen():
    logger.info("xlAutoOpen!")
    name = xll.Excel(lib.xlGetName)

    logger.info(f"Python XLL Loaded from {name}")

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


@ffi.def_extern(error=0)
def xlAutoClose():
    logger.info("xlAutoClose", flush=True)
    return 1


@ffi.def_extern(error=0)
def xlAutoAdd():
    logger.info("xlAutoAdd", flush=True)
    return 1


@ffi.def_extern(error=0)
def xlAutoRemove():
    logger.info("xlAutoRemove", flush=True)
    return 1


@ffi.def_extern()
def xlAutoFree12(xloper):
    # cleanup memory from strings
    logger.debug(f"xlAutoFree12({xloper!r}")
    if xloper.xltype == lib.xltypeStr | lib.xlbitDLLFree:
        lib.free(xloper.val.str)
        xloper.xltype = 0
    return None


@ffi.def_extern(error=xlerrNull)
def xlAutoRegister12(xloper):
    # cleanup memory from strings
    logger.debug(f"xlAutoFree12({xloper!r}")
    raise NotImplementedError


@ffi.def_extern(error=xlerrNull)
def xlAddInManagerInfo12(xloper):
    # cleanup memory from strings
    logger.debug(f"xlAddInManagerInfo12({xloper!r}")

    result = to_xloper(sys.version)
    result.xltype |= lib.xlbitDLLFree
    return result
