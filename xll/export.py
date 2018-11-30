from _python_xll import ffi, lib

import cffi

ffi = cffi.FFI()
ffi.cdef("""
    BOOL __stdcall VirtualProtect(
        LPVOID lpAddress, 
        SIZE_T dwSize, 
        DWORD flNewProtect,
        PDWORD lpflOldProtect
        );

    #define PAGE_EXECUTE_READ 32
    #define PAGE_EXECUTE 16
    #define PAGE_WRITECOPY 8
    #define PAGE_NOCACHE 512 
    #define PAGE_READONLY 2 
    #define PAGE_READWRITE 4 
    #define PAGE_EXECUTE_READWRITE 64
    #define PAGE_WRITECOMBINE 1024 
    #define PAGE_GUARD 256 
    #define PAGE_NOACCESS 1 
    #define PAGE_EXECUTE_WRITECOPY 128 
""")

kernel32 = ffi.dlopen('kernel32.dll')

class ExportDirectory:
    def __init__(self, base, directory):
        self.base = base
        self.directory = directory

        self.AddressOfNames = ffi.cast(
            f"DWORD[{directory.NumberOfNames}]",
            lib.PointerFromAddress(directory.AddressOfNames)
        )
        self.AddressOfNameOrdinals = ffi.cast(
            f"WORD[{directory.NumberOfNames}]",
            lib.PointerFromAddress(directory.AddressOfNameOrdinals)
        )
        self.AddressOfFunctions = ffi.cast(
            f"DWORD[{directory.NumberOfFunctions}]",
            lib.PointerFromAddress(directory.AddressOfFunctions),
        )

        # remove the vtable protection on the export table, so we can change it
        flOldProtect = ffi.new("DWORD*")
        kernel32.VirtualProtect(
            self.AddressOfFunctions,
            ffi.sizeof(self.AddressOfFunctions),
            kernel32.PAGE_READWRITE,
            flOldProtect
        )

    def __iter__(self):
        for i in range(self.directory.NumberOfFunctions):
            yield ffi.string(ffi.cast("char*",
                lib.PointerFromAddress(self.AddressOfNames[i])
            )).decode('utf-8')

    def __setitem__(self, name, function):
        name = name.encode('utf-8')
        for i in range(self.directory.NumberOfFunctions):
            if name != ffi.string(ffi.cast("char*",
                            lib.PointerFromAddress(self.AddressOfNames[i]))):
                continue

            index = self.AddressOfNameOrdinals[i] # - self.directory.Base
            self.AddressOfFunctions[index] = lib.AddressFromPointer(function)
            return

        raise KeyError(f'could not find export {name}')
