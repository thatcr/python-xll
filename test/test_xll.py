import sys
import os.path

def test_xll(Application):
    filename = os.path.join(os.path.dirname(os.path.dirname(__file__)), "python.dll")

    assert Application.RegisterXll(filename)

    assert os.name == Application.ExecuteExcel4Macro('os.name()')
    assert os.environ['PATH'] == Application.ExecuteExcel4Macro('os.getenv("PATH")')






