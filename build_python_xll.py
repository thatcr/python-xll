import os.path
import build_xlcall
from cffi import FFI

ffi = FFI()

ffi.include(build_xlcall.ffi)

ffi.cdef('''
    typedef struct _IMAGE_EXPORT_DIRECTORY {
        DWORD	Characteristics;
        DWORD	TimeDateStamp;
        WORD	MajorVersion;
        WORD	MinorVersion;
        DWORD	Name;
        DWORD	Base;
        DWORD	NumberOfFunctions;
        DWORD	NumberOfNames;
        DWORD	AddressOfFunctions;
        DWORD	AddressOfNames;
        DWORD	AddressOfNameOrdinals;
    } IMAGE_EXPORT_DIRECTORY,*PIMAGE_EXPORT_DIRECTORY;
    
    DWORD                   pRVABase;
    IMAGE_EXPORT_DIRECTORY* pExportDirectory;
    
    void* PointerFromAddress(DWORD Address);    
    DWORD AddressFromPointer(void* Pointer);          
''')

ffi.embedding_api('''      
    extern "Python" int __stdcall xlAutoOpen(void);
    extern "Python" int __stdcall xlAutoClose(void);
    extern "Python" int __stdcall xlAutoAdd(void);
    extern "Python" int __stdcall xlAutoRemove(void);
    
    extern "Python" void __stdcall xlAutoFree12(LPXLOPER12);
    
    extern "Python" LPXLOPER12 __stdcall _xlfCaller();
    extern "Python" LPXLOPER12 __stdcall py_eval(const char* expression);             
''')


ffi.set_source("_python_xll", r"""
#include <assert.h>
#include <WINDOWS.H>
#include <XLCALL.H>

#pragma comment(linker, "/export:xlAutoOpen=_xlAutoOpen@0")
#pragma comment(linker, "/export:xlAutoClose=_xlAutoClose@0")
#pragma comment(linker, "/export:xlAutoAdd=_xlAutoAdd@0")
#pragma comment(linker, "/export:xlAutoRemove=_xlAutoRemove@0")
#pragma comment(linker, "/export:xlAutoFree12=_xlAutoFree12@4")

#pragma comment(linker, "/export:py.eval=_py_eval@4")
#pragma comment(linker, "/export:xlfCaller=__xlfCaller@0")


extern DWORD pRVABase = 0;
IMAGE_EXPORT_DIRECTORY* pExportDirectory = 0; 

extern void* PointerFromAddress(DWORD Address)
{
    return (void*)(Address + pRVABase);    
}


extern DWORD AddressFromPointer(void* Pointer)
{
    return (DWORD)(Pointer) - pRVABase;    
}

// initialize this from the 
extern void** exports= 0;

void _set_export_directory(HINSTANCE hInstDll)
{            
    // figure out the location of the export able, this is easier than 
    // using python
    pRVABase = (DWORD) hInstDll;
    
    IMAGE_DOS_HEADER* pDosHeader = 
        (IMAGE_DOS_HEADER*) hInstDll;
    assert(pDosHeader->e_magic == 0x4550);    
    
    IMAGE_NT_HEADERS32* pNtHeaders32 = 
        (IMAGE_NT_HEADERS32*) (pDosHeader->e_lfanew + pRVABase);
    assert(pNtHeaders32->Signature == 0x4550);
    
    assert(pNtHeaders32->OptionalHeader.DataDirectory[0].VirtualAddress != 0);
    assert(pNtHeaders32->OptionalHeader.DataDirectory[0].Size != 0);
        
    pExportDirectory = 
        (IMAGE_EXPORT_DIRECTORY*)     
        (pNtHeaders32->OptionalHeader.DataDirectory[0].VirtualAddress + pRVABase);                             
}

#define E(n) extern CFFI_DLLEXPORT __declspec(naked) _ ## n() { ExitProcess(0x00FF0000); }

void _set_python_home()
{
    CHAR   DllName[MAX_PATH];
    sprintf(DllName, "python%d%d.dll", PY_MAJOR_VERSION, PY_MINOR_VERSION);

    HMODULE hModule = GetModuleHandle(DllName);

    // configure the python home location from the path of the python dll
    // found when we were loaded
    static WCHAR   DllPath[MAX_PATH] = {0};
    GetModuleFileNameW(hModule, DllPath, _countof(DllPath));
    *wcsrchr(DllPath, L'\\') = 0;
    Py_SetPythonHome(DllPath);                
}

BOOL WINAPI DllMain(HINSTANCE hInstDLL, DWORD fdwReason, LPVOID lpvReason)
{                        
    DisableThreadLibraryCalls(hInstDLL);
    
    if (fdwReason == DLL_PROCESS_ATTACH) {                  
        _set_export_directory(hInstDLL);
        
        if (Py_IsInitialized())
            return TRUE;

        _set_python_home();                
    }
            
    
    return TRUE;
}
 
"""+ '\n'.join(f"E({d:03X})" for d in range(0, 0xFFF)), include_dirs=[build_xlcall.src_dir])

ffi.embedding_init_code(r"""
import sys
import os.path

from xll.output import OutputDebugStringWriter

# send all output to the debugging console
sys.stderr = sys.stdout = OutputDebugStringWriter()

# if we have a PYDEVD port set, then attach to the debugger
if 'PYDEVD_PORT' in os.environ:
    import pydevd
    pydevd.settrace(
        host=os.getenv('PYDEVD_HOST', 'localhost'),
        port=int(os.getenv('PYDEVD_PORT')),
        suspend=False,
        stdoutToServer=True,
        stderrToServer=True
    )

# defer to the python module
import xll.auto
""")

if __name__ == '__main__':
    ffi.compile(target=os.path.join(os.path.dirname(__file__), 'python.xll'))

