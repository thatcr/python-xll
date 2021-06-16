# def test_ExcelProcess(ExcelProcess):
#     assert ExcelProcess.pid

def test_application(excel_application):
    assert Application.Version =='16.0'
