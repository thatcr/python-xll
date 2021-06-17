# def test_ExcelProcess(ExcelProcess):
#     assert ExcelProcess.pid

def test_application(excel_application):
    assert excel_application.Version =='16.0'
