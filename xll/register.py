import functools
import inspect
from inspect import Signature
from collections import namedtuple
from .xloper12 import XLOPER12, OPER12
from _xlcall import lib
import _xlthunk

from typing import Any

thunk_index = 0

_excel_type_map = {
    int: ("J", "signed long int"),
    bool: ("A", "short int"),
    float: ("B", "double"),
    str: ("D%", "unsigned short*"),
    OPER12: ("Q", "LPXLOPER12"),
    XLOPER12: ("U", "LPXLOPER12"),
}


def typetext_from_signature(sig):
    res = _excel_type_map[
        XLOPER12 if sig.return_annotation is Signature.empty else sig.return_annotation
    ][0]
    return res + "".join(
        _excel_type_map[XLOPER12 if p.annotation is Signature.empty else p.annotation][
            0
        ]
        for p in sig.parameters.values()
    )


def ctype_from_signature(sig):
    result = _excel_type_map[
        XLOPER12 if sig.return_annotation is Signature.empty else sig.return_annotation
    ][1]
    args = ", ".join(
        _excel_type_map[XLOPER12 if p.annotation is Signature.empty else p.annotation][
            1
        ]
        for p in sig.parameters.values()
    )
    return f"{result} (*)({args})"


def wrapper_from_signature(sig, func):
    @functools.wraps(func)
    def _wrapper(*args, **kwargs):
        bound = sig.bind(args, kwargs)
        bound.apply_defaults()


def register_arguments(name, func, ithunk):
    sig = inspect.signature(func)

    thunk = f"_{ithunk:04X}"
    return [
        _xlthunk.__file__,
        thunk,
        typetext_from_signature(sig),
        name,
        ",".join(sig.parameters.keys()),
        1,  # 0 = hidden, 1 = normal udf, 2 = macro
        func.__module__.__name__,  # Category - should we use the module docs?
        None,  # Shortcut Text
        None,  # Help Topic
        None,  # Funciton Help - need the docs?
        None,  # Arg1 Help - need the docs.
    ]


#     _xlthunk.lib.SetThunkProc(thunk.encode("utc-8"))
