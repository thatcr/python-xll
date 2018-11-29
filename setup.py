from setuptools import setup

# TODO addin xll as a python module
# TODO add the python.xll build to deliver into the prefix?
# TODO how to build the python.xll?

setup(
    name='python-xll',
    version="0",
    packages=[
        'xll'
    ],
    platforms = ['Windows'],
    cffi_modules=[
        'build_xlcall.py:ffi',
    ],
    entry_points = {
        'python.xll' : [
            "os.getenv=os:getenv",
            "os.cpu_count=os:cpu_count",
            "os.name=os.name"
        ]
    }
)

