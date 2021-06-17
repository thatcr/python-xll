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
    
    IMAGE_EXPORT_DIRECTORY* pExportDirectory;
    
    // these convert to/from the DWORD offsets in the export table
    void* PointerFromAddress(DWORD Address);    
    DWORD AddressFromPointer(void* Pointer);          
''')

ffi.embedding_api('''      
    extern "Python" int xlAutoOpen(void);
    extern "Python" int xlAutoClose(void);
    extern "Python" int xlAutoAdd(void);
    extern "Python" int xlAutoRemove(void);
    
    extern "Python" void xlAutoFree12(LPXLOPER12);
    
    extern "Python" LPXLOPER12 _xlfCaller();
    extern "Python" LPXLOPER12 py_eval(const char* expression);             
''')


ffi.set_source("_python_xll", r"""
#undef NDEBUG
#include <assert.h>
#include <WINDOWS.H>
#include <XLCALL.H>

#ifdef _M_IX86
#pragma comment(linker, "/export:xlAutoOpen=_xlAutoOpen@0")
#pragma comment(linker, "/export:xlAutoClose=_xlAutoClose@0")
#pragma comment(linker, "/export:xlAutoAdd=_xlAutoAdd@0")
#pragma comment(linker, "/export:xlAutoRemove=_xlAutoRemove@0")
#pragma comment(linker, "/export:xlAutoFree12=_xlAutoFree12@4")

#pragma comment(linker, "/export:py.eval=_py_eval@4")
#pragma comment(linker, "/export:xlfCaller=__xlfCaller@0")
#endif


// use a pointer to bytes so that the offset arithmetic works
static LPBYTE pRVABase = 0; 

IMAGE_EXPORT_DIRECTORY* pExportDirectory = 0; 

extern void* PointerFromAddress(DWORD Address)
{
    return (void*)(pRVABase + Address);    
}

extern DWORD AddressFromPointer(void* Pointer)
{
    return (DWORD) (((LPBYTE) Pointer) - pRVABase);    
}

// initialize this from the 
extern void** exports= 0;

void _set_export_directory(HINSTANCE hInstDll)
{            
    // figure out the location of the export table, this is easier than 
    // using python and working around type/pointer casts in cffi
    pRVABase = (void*) hInstDll;
    
    IMAGE_DOS_HEADER* pDosHeader = 
        (IMAGE_DOS_HEADER*) hInstDll;
    assert(pDosHeader->e_magic == 0x5A4D);             

    IMAGE_NT_HEADERS* pNtHeaders = 
        (IMAGE_NT_HEADERS*) (pDosHeader->e_lfanew + pRVABase);    
       
    assert(pNtHeaders->Signature == 0x4550);
        
    assert(pNtHeaders->OptionalHeader.DataDirectory[0].VirtualAddress != 0);
    assert(pNtHeaders->OptionalHeader.DataDirectory[0].Size != 0);
        
    pExportDirectory = 
        (IMAGE_EXPORT_DIRECTORY*)     
        (pNtHeaders->OptionalHeader.DataDirectory[0].VirtualAddress + pRVABase);                             
    
}

# this won't work on x74 - as the PE image is loaded above the heap address
# space so we cannot patch the export table with heap allocated stuff. on 32-bit
# we were lucky.

# so how to do the same - a JMP direct to the pointer would work, so just
# use two arrays next to each other. however that cannot be done easily here
# as we have no assembly inlines.

# we could embed a bunch of static thunks here that we _poke_ into
# as entry points... with enough code to run.
# if we call a zero-args entry point, how do we fix the stack
# so that we can exctract the values nicely - problem is that we don't know
# the codes - but all we need is the size... however some args aren't stack
# passed, so we're stuck - we need cffi magic here, and a JMP or direct export
# table reference. 

# can we use setjmp/longjmp to preseve the stack... 

# can we write assembler that does a JMP funcs[i] for any given I and patch it in.




#define E(n) extern CFFI_DLLEXPORT _ ## n() { ExitProcess(0x00FF0000); }

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
        _set_export_directory(hInstDLL);
        
        if (Py_IsInitialized())
            return TRUE;

        _set_python_home();                                 
    }
                
    return TRUE;
}
 
"""+ '\n'.join(f"E({d:03X})" for d in range(0, 0xFFF)), include_dirs=[build_xlcall.src_dir])

ffi.embedding_init_code(r"""
import logging
import sys
import sys
import os.path

from xll.output import OutputDebugStringWriter

# send all output to the debugging console
sys.stderr = sys.stdout = OutputDebugStringWriter()

print("Hello World!")

# logging.basicConfig(filename=r"C:\Users\thatc\workspaces\python-xll\python-xll.log", level=logging.DEBUG)
logging.basicConfig(level=logging.DEBUG)
logging.info("Hello World!")
logging.info(f"sys.prefix = {sys.path!r}")
logging.info(f"sys.path = {sys.path!r}")


# defer to the python module to 
import xll.auto

""")

if __name__ == '__main__':
    ffi.compile(target=os.path.join(os.path.dirname(__file__), 'python.xll'), debug=False)

