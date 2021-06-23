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
    def __init__(self, *, _from_excel=None):
        if _from_excel is not None:
            self.ptr = ffi.cast("LPXLOPER12", _from_excel)
            return

        # create an array of two XLOPERs, value is in the first one
        self._array = ffi.new("XLOPER12[2]")

        # second XLOPER has a pointer to this object so we can divine
        # it in the xlAutoFree12 callback
        self._array[1].xltype = 0xFFFF
        self._array[1].val.str = ffi.cast("wchar_t*", id(self))

        self.ptr = ffi.addressof(self._array[0])
        self.ptr.xltype = lib.xltypeMissing

    @property
    def xltype(self):
        return self.ptr.xltype & 0xFFF

    @classmethod
    def from_excel(cls, value):
        return cls(_from_excel=value)

    @classmethod
    def from_float(cls, value):
        self = cls()
        self.ptr.xltype = lib.xltypeNum
        self.ptr.val.num = float(value)
        return self

    def __float__(self):
        if self.xltype == lib.xltypeNum:
            return float(self.ptr.val.num)
        raise TypeError("XLOPER12 is not xltypeNum")

    @classmethod
    def from_string(cls, value):
        self = cls()
        self.ptr.xltype = lib.xltypeStr

        self.ptr.val.str = lib.malloc((len(value) + 1) * 2)
        self.ptr.val.str[0] = chr(len(value))
        self.ptr.val.str[1 : len(value) + 1] = value
        return self

    def __str__(self):
        if self.xltype == lib.xltypeStr:
            return ffi.string(self.ptr.val.str + 1, ord(self.ptr.val.str[0]))
        raise TypeError("XLOPER12 is not xltypeStr")

    @classmethod
    def from_bool(cls, value):
        self = cls()
        self.ptr.xltype = lib.xltypeBool
        self.ptr.val.xbool = 1 if value else 0
        return self

    def __bool__(self):
        if self.xltype == lib.xltypeBool:
            return bool(self.ptr.val.xbool)
        raise TypeError("XLOPER12 is not xltypeBool")

    @classmethod
    def from_int(cls, value):
        # TODO check that we are within the rnage for a short int integer

        self = cls()
        self.ptr.xltype = lib.xltypeInt
        self.ptr.val.w = int(value)
        return self

    def __int__(self):
        if self.xltype == lib.xltypeInt:
            return int(self.ptr.val.w)
        raise TypeError("XLOPER12 is not xltypeInt")

    @classmethod
    def from_none(cls, value=None):
        assert value is None
        self = cls()
        self.ptr.xltype = lib.xltypeNil
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
        xltype = self.ptr.xltype & 0xFFF
        return _to_python[xltype](self)

    def __repr__(self):
        value = self.to_python()
        xltype = "|".join([k for k, v in _xltypes.items() if self.ptr.xltype & v])
        return f"<{self.__class__.__name__} {xltype} {value!r}>"

    def as_result(self):
        if self.ptr.xltype & lib.xlbitDLLFree:
            raise RuntimeError("XLOPER12.as_result was called twice")

        # get the refcount from the object and increment it for excel's reference
        self.ptr.xltype |= lib.xlbitDLLFree
        prefcnt = ffi.cast("unsigned int*", id(self))
        prefcnt[0] = prefcnt[0] + 1

        return self.ptr

    def __del__(self):
        log.debug("XLOPER.__del__ called")
        if hasattr(self, "ptr") and self.ptr.xltype & lib.xltypeStr:
            lib.free(self.ptr.val.str)
            self.ptr.xltype = lib.xltypeMissing

        # TODO call xlfree here if the bit is set?
        # _xlcall import lib
