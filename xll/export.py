import logging
from _python_xll import ffi, lib

import cffi

logger = logging.getLogger(__name__)

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
    def __init__(self, directory):   
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

        logger.info(f"Export Directory has {directory.NumberOfNames} names and {directory.NumberOfFunctions} functions.")

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

    # TODO slice assignment to be faster... or allocate names? 
    def __setitem__(self, name, function):
        
        logger.info(f"pRVABase = {lib.PointerFromAddress(0)}")        

        # it look slike the addresses of these functions on the heap are 
        # lower than the pRVAbase, so we're shafted. going to need a better
        # way to thunk. Can we trampoline from the entry points? 
        
        assert lib.AddressFromPointer(lib.PointerFromAddress(0)) == 0
        
        name = name.encode('utf-8')
        for i in range(self.directory.NumberOfFunctions):
            if name != ffi.string(ffi.cast("char*",
                            lib.PointerFromAddress(self.AddressOfNames[i]))):
                continue

            logger.info(f"Assigning entry point {name!s} @ {i} = {function!r}")
            index = self.AddressOfNameOrdinals[i]
            voidp = ffi.cast("void*", function)
            
            addr = lib.AddressFromPointer(voidp)
            
            logger.info(f"Index is {index!r}, addr = {addr!r}")            
            self.AddressOfFunctions[index] = addr

            ptr = lib.PointerFromAddress(self.AddressOfFunctions[index])            
            
            logger.info(f" ptr = {ptr!r}, voidp = {voidp!r}")            
            return

        raise KeyError(f'could not find export {name}')
