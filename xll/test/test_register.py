import inspect

from ..register import typetext_from_signature, ctype_from_signature


def test_typecode():
    def add(a: int, b: int) -> int:
        """
        Add two numbers together.

        Parameters:

        a:int The first number
        b:int The second number

        Returns:

        int: The sum of the numbers.
        """
        ...

    sig = inspect.signature(add)
    assert typetext_from_signature(sig) == "JJJ"
    assert (
        ctype_from_signature(sig)
        == "signed long int (*)(signed long int, signed long int)"
    )

    def untyped(a, b):
        ...

    sig = inspect.signature(untyped)
    assert typetext_from_signature(sig) == "UUU"
    assert ctype_from_signature(sig) == "LPXLOPER12 (*)(LPXLOPER12, LPXLOPER12)"
