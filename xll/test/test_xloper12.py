import sys
import pytest

from ..xloper12 import XLOPER12

from _xlcall import lib
from _xlthunk.lib import xlAutoFree12


def test_null(value=None):
    assert XLOPER12.from_python(value).to_python() is None
    assert XLOPER12.from_none(value).to_python() is None


def test_ellipsis(value=Ellipsis):
    assert XLOPER12().to_python() is Ellipsis


@pytest.mark.parametrize("value", [True, False])
def test_bool(value):
    assert XLOPER12.from_python(value).to_python() == value
    assert XLOPER12.from_bool(value).to_python() == value
    assert bool(XLOPER12.from_bool(value)) is value
    assert repr(XLOPER12.from_bool(value)) == f"<XLOPER12 xltypeBool {value!r}>"

    with pytest.raises(TypeError):
        bool(XLOPER12.from_none())


@pytest.mark.parametrize("value", [-1.0, 1.0, 0.0, 123.123])
def test_float(value):
    assert XLOPER12.from_python(value).to_python() == value
    assert XLOPER12.from_float(value).to_python() == value
    assert float(XLOPER12.from_float(value)) == value
    assert repr(XLOPER12.from_float(value)) == f"<XLOPER12 xltypeNum {value!r}>"

    with pytest.raises(TypeError):
        float(XLOPER12.from_none())


@pytest.mark.parametrize("value", ["a", "fish", sys.version])
def test_string(value):
    assert XLOPER12.from_python(value).to_python() == value
    assert XLOPER12.from_string(value).to_python() == value
    assert str(XLOPER12.from_string(value)) == value
    assert repr(XLOPER12.from_string(value)) == f"<XLOPER12 xltypeStr {value!r}>"

    with pytest.raises(TypeError):
        str(XLOPER12.from_none())

    # check that we de-allocate correctly? use a mock?


@pytest.mark.parametrize("value", [-1, 0, 1, 123])
def test_int(value):
    assert XLOPER12.from_python(value).to_python() == value
    assert XLOPER12.from_int(value).to_python() == value
    assert int(XLOPER12.from_int(value)) == value
    assert repr(XLOPER12.from_int(value)) == f"<XLOPER12 xltypeInt {value!r}>"

    with pytest.raises(TypeError):
        int(XLOPER12.from_none())


def test_unknown_type(value=Ellipsis):
    with pytest.raises(TypeError):
        XLOPER12.from_python(value)


def test_result():
    xloper = XLOPER12()

    refcount = sys.getrefcount(xloper)

    result = xloper.to_result()
    assert sys.getrefcount(xloper) == refcount + 1
    xlAutoFree12(result)
    assert sys.getrefcount(xloper) == refcount
    del xloper


def test_wrong_pointer_type():
    with pytest.raises(TypeError):
        XLOPER12(ptr="The Wrong Type!")
