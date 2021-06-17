import xll
import sys
import logging

from xll.api import Excel

import inspect

from _python_xll import ffi, lib
#from _xlcall import lib

from .convert import to_xloper
from .export import ExportDirectory

logger = logging.getLogger(__name__)

def onerror(exception, exc_value, traceback):
    sys.excepthook(exception, exc_value, traceback)

    xlerrValue = ffi.new('LPXLOPER12')
    xlerrValue.xltype = lib.xltypeErr | lib.xlbitDLLFree
    xlerrValue.val.err = lib.xlerrValue

    return xlerrValue

@ffi.def_extern(error=None)
def xlAutoFree12(xloper):
    print('xlAutoFree12', xloper, id(xloper))

@ffi.callback("LPXLOPER12 (*)()")
def xlfCaller():
    logger.info("Called xlfCaller")
    caller = Excel(lib.xlfCaller, convert=False)

    text = Excel(lib.xlSheetNm, caller)

    result = to_xloper(text)
    result.xltype |= lib.xlbitDLLFree

    print(result)

    ptr = ffi.gc(result, destructor=lambda x : None)
    print(ptr, id(ptr))

    return result

@ffi.callback("LPXLOPER12 (*)(const char*)")
def py_eval(source):
    logging.info(f"Called py.eval with {source}")
    result = to_xloper(eval(ffi.string(source), locals(), {}))
    result.xltype |= lib.xlbitDLLFree
    return result

@ffi.def_extern(error=0)
def xlAutoOpen():
    logger.info("xlAutoOpen!")    
    name =  xll.Excel(lib.xlGetName)

    logger.info(f"Python XLL Loaded from {name}")

    exports = ExportDirectory(lib.pExportDirectory)

    # # xll.Excel(lib.xlcAlert, "Hello World")
    exports['_000'] = py_eval
    xll.Excel(lib.xlfRegister,
              name,
              "_000",
              "QC",
              "PY.EVAL",
              "source",
              1,
              "Python Builtins",
              None,
              None,
              "Evaluate a python expression",
              "Python expression to evaluate"
              )

    exports['_001'] = xlfCaller
    xll.Excel(lib.xlfRegister,
              name,
              "_001",
              "Q",
              "xlfCaller",
              "",
              1,
              "Python Builtins",
              None,
              None,
              None,
              None
              )
    print ('done')

    return 1


@ffi.def_extern(error=0)
def xlAutoClose():
    print('xlAutoClose', flush=True)
    return 1


@ffi.def_extern(error=0)
def xlAutoAdd():
    print('xlAutoAdd', flush=True)
    return 1


@ffi.def_extern(error=0)
def xlAutoRemove():
    print('xlAutoRemove', flush=True)
    return 1


