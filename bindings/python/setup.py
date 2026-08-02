import os
import sys
from setuptools import setup
from pybind11.setup_helpers import Pybind11Extension, build_ext

xlpp_root = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))

debug = os.environ.get("XLPP_DEBUG", "0") == "1"
config = "Debug" if debug else "Release"

extra_compile_args = ["/std:c++20", "/EHsc", "/bigobj", "/DXLPP_STATIC"]
if debug:
    extra_compile_args.append("/MDd")
else:
    extra_compile_args.append("/MD")

ext_modules = [
    Pybind11Extension(
        "xlpp",
        ["src/xlpp_bindings.cpp"],
        include_dirs=[
            os.path.join(xlpp_root, "include"),
            os.path.join(xlpp_root, "third_party", "zlib"),
        ],
        library_dirs=[
            os.path.join(xlpp_root, "x64", config),
        ],
        libraries=["XLPP", "zlib"],
        extra_compile_args=extra_compile_args,
        language="c++",
    ),
]

setup(
    name="xlpp",
    version="1.0.0",
    author="XL++ contributors",
    description="High-performance C++ Excel xlsx library for Python",
    long_description="XL++ is a high-performance C++20 library for reading and writing .xlsx files, with a Python binding API modeled after openpyxl.",
    ext_modules=ext_modules,
    cmdclass={"build_ext": build_ext},
    python_requires=">=3.8",
    classifiers=[
        "Development Status :: 4 - Beta",
        "License :: OSI Approved :: MIT License",
        "Programming Language :: Python :: 3",
        "Programming Language :: Python :: 3.8",
        "Programming Language :: Python :: 3.9",
        "Programming Language :: Python :: 3.10",
        "Programming Language :: Python :: 3.11",
        "Programming Language :: Python :: 3.12",
        "Programming Language :: Python :: 3.13",
        "Programming Language :: Python :: 3.14",
    ],
)
