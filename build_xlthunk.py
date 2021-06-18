import os.path
from cffi import FFI

ffi = FFI()

src_dir = os.path.join(os.path.dirname(__file__), "src")

ffi.cdef(
    """
    void* SetThunkProc(const char* name, void* ptr);        
"""
)

MAX_THUNKS = 0x7FFF

ffi.set_source(
    "_xlthunk",
    r"""
#include <WINDOWS.H>
#include <XLCALL.H>

static HMODULE hModule = NULL;

void __declspec(dllexport) xlAutoFree12(LPXLOPER12 lpXloper)
{
    OutputDebugString("xlAutoFree12!\n\n");

    PyObject** pobj = (void*) (lpXloper + 1);

    CHAR sz[1024];
    sprintf(sz, "XlAutoFree12 of XLOPER @ %I64X has pyobject @ %I64X with refcount = %I64d\n\n", (unsigned __int64) lpXloper, (unsigned __int64) *pobj, (*pobj)->ob_refcnt);
    OutputDebugString(sz);

    PyGILState_STATE state = PyGILState_Ensure();
    Py_DECREF(*pobj);
    PyGILState_Release(state);
}

BOOL WINAPI DllMain(HINSTANCE hInstDLL, DWORD fdwReason, LPVOID lpvReason)
{                        
    DisableThreadLibraryCalls(hInstDLL);
    
    if (fdwReason == DLL_PROCESS_ATTACH) {                                          
        hModule = hInstDLL;                           
    }
                
    return TRUE;
}

extern void* SetThunkProc(const char* name, void* ptr)
{        
    FARPROC proc = GetProcAddress(hModule, name);
    if (proc == NULL)
        return NULL;

    // cast and assign the location of the jump pointer 
    void** addr = (void**) ((unsigned char*) proc + 2);
    *addr = ptr;

    // change the page permissions so that we can exceute the thunk
    DWORD oldProtect;
    if (!VirtualProtect(proc, 13, PAGE_EXECUTE_READWRITE, &oldProtect))
        return NULL;

    return proc;          
}

// http://kylehalladay.com/blog/2020/11/13/Hooking-By-Example.html

// thunks are defined as small exported arrays of bytes that trampoline
// to a pointer injected by from the SetThunkProc call below
#define E(n) extern __declspec(dllexport) unsigned char  _ ## n ## [13] = \
    { 0x49, 0xBA, \
      0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, \
      0x41, 0xFF, 0xE2 \
    };

    

"""
    + "\n".join(f"E({d:04X})" for d in range(0, MAX_THUNKS)),
    include_dirs=[src_dir],
)


if __name__ == "__main__":
    ffi.compile(
        target=os.path.join(os.path.dirname(__file__), "_xlthunk.pyd"), verbose=True
    )
