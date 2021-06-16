import sys
import os.path

def test_xll(excel_application):
    filename = os.path.join(os.path.dirname(os.path.dirname(__file__)), "python.xll")

    assert excel_application.RegisterXll(filename)

    assert os.name == excel_application.ExecuteExcel4Macro('os.name()')
    assert os.environ['PATH'] == excel_application.ExecuteExcel4Macro('os.getenv("PATH")')






