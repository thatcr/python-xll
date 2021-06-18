import os.path
import build_xlcall
from cffi import FFI

ffi = FFI()

ffi.include(build_xlcall.ffi)

# NOTE xlAutoFree12 has to exist in the same module as the function
ffi.embedding_api(
    """
    extern "Python" int xlAutoOpen(void);
    extern "Python" int xlAutoClose(void);
    extern "Python" int xlAutoAdd(void);
    extern "Python" int xlAutoRemove(void);    
    extern "Python" LPXLOPER12 xlAutoRegister12();
    extern "Python" LPXLOPER12 xlAddInManagerInfo12();    
"""
)

ffi.set_source(
    "_python_xll",
    r"""
#include <assert.h>
#include <WINDOWS.H>
#include <XLCALL.H>

#ifdef _M_IX86
// rename the mangled functions to plain as excel expected them
#pragma comment(linker, "/export:xlAutoOpen=_xlAutoOpen@0")
#pragma comment(linker, "/export:xlAutoClose=_xlAutoClose@0")
#pragma comment(linker, "/export:xlAutoAdd=_xlAutoAdd@0")
#pragma comment(linker, "/export:xlAutoRemove=_xlAutoRemove@0")
#pragma comment(linker, "/export:xlAutoFree12=_xlAutoFree12@4")
#endif


struct PyXLOPER12 
{ 
    struct xloper12 xlo; 
    void* ptr; 
};

void _set_python_home()
{
    // NOTE if we are inside a venv we see the pythonXX.dll inside
    // the Scripts/ subfolder, and need to setup the right home
    // from the pyvenv.cfg first? How does the python.exe do that?

    CHAR   DllName[MAX_PATH];

    sprintf(DllName, "python%d%d.dll", PY_MAJOR_VERSION, PY_MINOR_VERSION);

    HMODULE hModule = GetModuleHandle(DllName);

    // configure the python home location from the path of the python dll
    // found when we were loaded
    static WCHAR   DllPath[MAX_PATH + 1] = {0};
    GetModuleFileNameW(hModule, DllPath, MAX_PATH);    
        
    *wcsrchr(DllPath, L'\\') = 0;
    
    // if the last part of the stringis 'Scripts' it's a venv... 
    // how does that work? 

    Py_SetPythonHome(DllPath);       
}

BOOL WINAPI DllMain(HINSTANCE hInstDLL, DWORD fdwReason, LPVOID lpvReason)
{                        
    DisableThreadLibraryCalls(hInstDLL);
    
    if (fdwReason == DLL_PROCESS_ATTACH) {                                                  
        if (Py_IsInitialized())
            return TRUE;

        _set_python_home();                                 
    }
                
    return TRUE;
}

""",
    include_dirs=[build_xlcall.src_dir],
)

ffi.embedding_init_code(
    r"""
import logging
import os.path
import sys

# patch the exceutable path, or lots of modules get confused
sys.executable = os.path.join(sys.prefix, "python.exe")

from xll.output import OutputDebugStringWriter

import _python_xll

# send all output to the debugging console
sys.stderr = sys.stdout = OutputDebugStringWriter()

# log to the new stdout/stderr channels
logging.basicConfig(level=logging.DEBUG)

# defer to the python module to add the real registration bits
import xll.auto
"""
)


if __name__ == "__main__":
    ffi.compile(
        target=os.path.join(os.path.dirname(__file__), "python.xll"), verbose=True
    )
