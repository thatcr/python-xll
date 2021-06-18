import sys
from _xlcall import lib


def test_py_eval_simple(excel_application):
    assert 2 == excel_application.ExecuteExcel4Macro('PY.EVAL("1+1")')
    assert "fish" == excel_application.ExecuteExcel4Macro("PY.EVAL(\"'fish'\")")


def test_exception(excel_application):
    assert excel_application.ExecuteExcel4Macro('TYPE(PY.EVAL("1/0"))') == lib.xltypeErr
    assert excel_application.ExecuteExcel4Macro('ERROR.TYPE(PY.EVAL("1/0"))') == 3.0

    # https://support.microsoft.com/en-us/office/error-type-function-10958677-7c8d-44f7-ae77-b9a9ee6eefaa
    # codes from ERROR.TYPE are not th esame as xlerrValue...


def test_one_plus_one(excel_application):
    assert sys.version == excel_application.ExecuteExcel4Macro('PY.EVAL("sys.version")')
    #
    # assert "fish" == excel_application.ExecuteExcel4Macro("PY.EVAL(\"'fish'\")")
