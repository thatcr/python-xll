import xll
import sys
import os.path
import logging

from xll.api import Excel

import inspect

from _python_xll import ffi, lib
import _xlthunk

from .convert import to_xloper

logger = logging.getLogger(__name__)

def onerror(exception, exc_value, traceback):
    sys.excepthook(exception, exc_value, traceback)

    xlerrValue = ffi.new('LPXLOPER12')
    xlerrValue.xltype = lib.xltypeErr | lib.xlbitDLLFree
    xlerrValue.val.err = lib.xlerrValue

    return xlerrValue

@ffi.def_extern(error=None)
def xlAutoFree12(xloper):
    logger.info(f'xlAutoFree12 {xloper!r} {id(xloper)}')

@ffi.callback("LPXLOPER12 (*)()")
def xlfCaller():
    logger.info("Called xlfCaller")
    caller = Excel(lib.xlfCaller, convert=False)

    text = Excel(lib.xlSheetNm, caller)

    result = to_xloper(text)
    result.xltype |= lib.xlbitDLLFree
    
    # how is result kept as a return value? 
        
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
    
    _xlthunk.lib.SetThunkProc(b"_0000", py_eval)                
    xll.Excel(lib.xlfRegister,
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
              "Python expression to evaluate"
              )

    # exports['_001'] = xlfCaller
    _xlthunk.lib.SetThunkProc(b"_0001", xlfCaller)        
    xll.Excel(lib.xlfRegister,
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
              None
              )    
    return 1


@ffi.def_extern(error=0)
def xlAutoClose():
    logger.info('xlAutoClose', flush=True)
    return 1


@ffi.def_extern(error=0)
def xlAutoAdd():
    logger.info('xlAutoAdd', flush=True)
    return 1


@ffi.def_extern(error=0)
def xlAutoRemove():
    logger.info('xlAutoRemove', flush=True)
    return 1


