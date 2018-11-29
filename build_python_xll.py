import os.path
from cffi import FFI

ffi = FFI()

ffi.embedding_api(f'''
    extern "Python" int __stdcall xlAutoOpen(void);
    extern "Python" int __stdcall xlAutoClose(void);
    extern "Python" int __stdcall xlAutoAdd(void);
    extern "Python" int __stdcall xlAutoRemove(void);    
''')


ffi.set_source("_python_xll", r"""
#pragma comment(linker, "/export:xlAutoOpen=_xlAutoOpen@0")
#pragma comment(linker, "/export:xlAutoClose=_xlAutoClose@0")
#pragma comment(linker, "/export:xlAutoAdd=_xlAutoAdd@0")
#pragma comment(linker, "/export:xlAutoRemove=_xlAutoRemove@0")

void _set_python_home()
{
    CHAR   DllName[MAX_PATH];
    sprintf(DllName, "python%d%d.dll", PY_MAJOR_VERSION, PY_MINOR_VERSION);

    HMODULE hModule = GetModuleHandle(DllName);

    static WCHAR   DllPath[MAX_PATH] = {0};
    GetModuleFileNameW(hModule, DllPath, _countof(DllPath));

    *wcsrchr(DllPath, L'\\') = 0;

    OutputDebugString("PySetPythonHome:");
    OutputDebugStringW(DllPath);
    OutputDebugString("\n");
        
    Py_SetPythonHome(DllPath);
}

BOOL WINAPI DllMain(HINSTANCE hInstDLL, DWORD fdwReason, LPVOID lpvReason)
{    
    if (Py_IsInitialized())
        return TRUE;

    DisableThreadLibraryCalls(hInstDLL);
    if (fdwReason == DLL_PROCESS_ATTACH) {        
        _set_python_home();        
    }
    return TRUE;
}
 
""")

ffi.embedding_init_code(r"""
from _python_xll import ffi

import sys
import os.path

from xll.output import OutputDebugStringWriter

sys.stderr = OutputDebugStringWriter()
sys.stdout = OutputDebugStringWriter()

print("Hello World")

import _xlcall
import xll
    
@ffi.def_extern(error=0)
def xlAutoOpen():
    print('xlAutoOpen', flush=True)
        
    from _xlcall import lib
    print('xlGetName', xll.Excel(lib.xlGetName), flush=True)
    
    xll.Excel(lib.xlcAlert, "Hello World");             
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
    
print("Loaded XLL")
""")

if __name__ == '__main__':
    ffi.compile(target=os.path.join(os.path.dirname(__file__), 'python.xll'))

