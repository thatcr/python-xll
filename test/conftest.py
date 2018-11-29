import pytest
import shutil
import os

from comtypes.client import CreateObject
#
# from _ctypes_win32 import (
#     EnumWindows, EnumChildWindows,
#     WNDENUMPROC, GetWindowThreadProcessId,
#     AccessibleObjectFromWindow,
#     OBJID_NATIVEOM
# )
#
# from ctypes import byref, POINTER
# from ctypes.wintypes import DWORD
# from comtypes.automation import IDispatch
# from comtypes.client import GetBestInterface, GetModule
#
# import subprocess

EXCEL = r"C:\Program Files (x86)\Microsoft Office\root\Office16\EXCEL.EXE"

# @pytest.fixture(scope='session')
# def ExcelProcess():
#     executable = shutil.which('EXCEL.EXE') or EXCEL
#
#     args = [executable]
#
#     startupinfo = subprocess.STARTUPINFO()
#
#     startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
#     startupinfo.wShowWindow = subprocess.SW_HIDE
#
#     process = subprocess.Popen(args, executable=executable, startupinfo=startupinfo, cwd=os.getcwd())
#
#     yield process
#     process.kill()
#     process.wait()


@pytest.fixture(scope='session')
def Application():
    Application = CreateObject('Excel.Application')
    yield Application


# @pytest.fixture(scope='session')
# def Application(ExcelProcess):
#     """Figure out the Excel.Application object for the process that we just launched.
#     This relies on the application having at least one window open so that it exposes
#     it's Window object to the windows accessibility layer.
#
#     We determine that the accssible object that exposes an 'Application' property
#     is supplying the application object. We could do this more safely by using
#     and GUID on the interface to make sure we have what we need.
#     """
#
#     _module = GetModule(
#         ("{00020813-0000-0000-C000-000000000046}", 1, 7))  # 1.7 is Excel 2010?
#
#     # place that closures below can put their results
#     res = {}
#
#     def enum_child(hwnd, _):
#         """loop over all child windows, find the application object"""
#         ptr = POINTER(IDispatch)()
#         try:
#             AccessibleObjectFromWindow(
#                 hwnd, OBJID_NATIVEOM,
#                 byref(IDispatch._iid_),  # pylint: disable=W0212
#                 byref(ptr)
#             )
#             ptr = GetBestInterface(ptr)
#             res['Application'] = ptr.Application
#             return False
#         except (AttributeError, WindowsError, NameError):
#             return True
#
#     def enum_toplevel(hwnd, _):
#         """loop over all top level windows"""
#         epid = DWORD()
#         GetWindowThreadProcessId(hwnd, byref(epid))
#         if epid.value == ExcelProcess.pid:
#             EnumChildWindows(hwnd, WNDENUMPROC(enum_child), 0)
#         return True
#
#     EnumWindows(WNDENUMPROC(enum_toplevel), 0)
#
#     if not 'Application' in res:
#         raise RuntimeError("Could not find Application object for pid %d" % ExcelProcess.pid)
#
#     return GetBestInterface(res['Application'])  # pylint
