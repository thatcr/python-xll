import pytest
from _xlcall import lib
from ..api import Excel


# simple test that API is working. it won't in pytest run for now
@pytest.mark.xfail(reason="Cannot test Excel calls from pytest", raises=RuntimeError)
def test_xlfNa():
    Excel(lib.xlfNa)
