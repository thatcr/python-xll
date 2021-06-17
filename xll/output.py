from cffi import FFI

ffi = FFI()
ffi.cdef(
    """
   void OutputDebugStringW(const wchar_t*);
"""
)
kernel32 = ffi.dlopen("kernel32.dll")


class OutputDebugStringWriter:
    def write(self, value):
        kernel32.OutputDebugStringW(value.rstrip())

    def flush(self):
        pass
