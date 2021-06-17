import sys
import os.path


def test_xll(excel_application):
    filename = os.path.join(os.path.dirname(os.path.dirname(__file__)), "python.xll")

    assert excel_application.RegisterXll(filename)
    assert 2 == excel_application.ExecuteExcel4Macro('PY.EVAL("1+1")')
