import logging

from _xlcall import ffi, lib

NoneType = type(None)

log = logging.getLogger(__name__)


_to_python = {
    lib.xltypeNum: float,
    lib.xltypeStr: str,
    lib.xltypeBool: bool,
    lib.xltypeNil: lambda _: None,
    lib.xltypeMissing: lambda _: Ellipsis,
    lib.xltypeInt: int,
}

# mapping of xltype names to their values
_xltypes = {
    k: v
    for k, v in lib.__dict__.items()
    if k.startswith("xltype") and k != "xltypeBigData"
}


class XLOPER12(object):
    def __init__(self, ptr=None):
        if ptr:
            self.ptr = ptr
            return

        # create an array of two XLOPERs, value is in the first one
        self.ptr = ffi.new("XLOPER12[2]")
        self.ptr[0].xltype = lib.xltypeMissing

        # second XLOPER has a pointer to this object so we can divine
        # it in the xlAutoFree12 callback
        self.ptr[1].xltype = 0xFFFF
        self.ptr[1].val.str = ffi.cast("wchar_t*", id(self))

    @property
    def xltype(self):
        return self.ptr[0].xltype & 0xFFF

    @classmethod
    def from_excel(cls, value):
        return cls(ptr=value)

    @classmethod
    def from_float(cls, value):
        self = cls()
        self.ptr[0].xltype = lib.xltypeNum
        self.ptr[0].val.num = float(value)
        return self

    def __float__(self):
        if self.xltype == lib.xltypeNum:
            return float(self.ptr[0].val.num)
        raise TypeError("XLOPER12 is not xltypeNum")

    @classmethod
    def from_string(cls, value):
        self = cls()
        self.ptr[0].xltype = lib.xltypeStr

        self.ptr[0].val.str = lib.malloc((len(value) + 1) * 2)
        self.ptr[0].val.str[0] = chr(len(value))
        self.ptr[0].val.str[1 : len(value) + 1] = value
        return self

    def __str__(self):
        if self.xltype == lib.xltypeStr:
            return ffi.string(self.ptr[0].val.str + 1, ord(self.ptr[0].val.str[0]))
        raise TypeError("XLOPER12 is not xltypeStr")

    @classmethod
    def from_bool(cls, value):
        self = cls()
        self.ptr[0].xltype = lib.xltypeBool
        self.ptr[0].val.xbool = 1 if value else 0
        return self

    def __bool__(self):
        if self.xltype == lib.xltypeBool:
            return bool(self.ptr[0].val.xbool)
        raise TypeError("XLOPER12 is not xltypeBool")

    @classmethod
    def from_int(cls, value):
        # TODO check that we are within the rnage for a short int integer

        self = cls()
        self.ptr[0].xltype = lib.xltypeInt
        self.ptr[0].val.w = int(value)
        return self

    def __int__(self):
        if self.xltype == lib.xltypeInt:
            return int(self.ptr[0].val.w)
        raise TypeError("XLOPER12 is not xltypeInt")

    @classmethod
    def from_none(cls, value=None):
        assert value is None
        self = cls()
        self.ptr[0].xltype = lib.xltypeNil
        return self

    @classmethod
    def from_python(cls, value):
        if isinstance(value, NoneType):
            return cls.from_none(value)
        if isinstance(value, bool):
            return cls.from_bool(value)
        if isinstance(value, int):
            return cls.from_int(value)
        if isinstance(value, float):
            return cls.from_float(value)
        if isinstance(value, str):
            return cls.from_string(value)
        raise TypeError(f"could not convert {type(value)} to {cls.__name__}")

    def to_python(self):
        xltype = self.ptr[0].xltype & 0xFFF
        return _to_python[xltype](self)

    def __repr__(self):
        value = self.to_python()
        xltype = "|".join([k for k, v in _xltypes.items() if self.ptr[0].xltype & v])
        return f"<{self.__class__.__name__} {xltype} {value!r}>"

    def to_result(self):
        # get the refcount from the object and increment it for excel's reference
        self.ptr[0].xltype |= lib.xlbitDLLFree
        prefcnt = ffi.cast("unsigned int*", id(self))
        prefcnt[0] = prefcnt[0] + 1

        return ffi.cast("LPXLOPER12", self.ptr)

    def __del__(self):
        log.debug("XLOPER.__del__ called")
        if self.ptr[0].xltype & lib.xltypeStr:
            lib.free(self.ptr[0].val.str)
            self.ptr[0].xltype = lib.xltypeMissing
