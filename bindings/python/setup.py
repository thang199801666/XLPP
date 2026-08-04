import os
from setuptools import setup
from pybind11.setup_helpers import Pybind11Extension, build_ext

xlpp_root = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))

debug = os.environ.get("XLPP_DEBUG", "0") == "1"
config = "Debug" if debug else "Release"

is_msvc = os.environ.get("CXX", "").lower().endswith(("cl", "cl.exe")) or os.name == "nt"
if is_msvc:
    extra_compile_args = ["/std:c++20", "/EHsc", "/bigobj", "/DXLPP_STATIC"]
    extra_link_args = []
else:
    extra_compile_args = ["-std=c++20", "-fPIC", "-DXLPP_STATIC"]
    extra_link_args = []

library_dir = os.path.join(xlpp_root, "x64", config)
if is_msvc:
    libraries = ["XLPP", "zlib"]
else:
    libraries = ["xlpp_static", "zlibstatic"]

ext_modules = [
    Pybind11Extension(
        "xlpp",
        ["src/xlpp_bindings.cpp"],
        include_dirs=[
            os.path.join(xlpp_root, "include"),
            os.path.join(xlpp_root, "third_party", "zlib"),
        ],
        library_dirs=[
            library_dir,
        ],
        libraries=libraries,
        extra_compile_args=extra_compile_args,
        extra_link_args=extra_link_args,
        language="c++",
    ),
]

try:
    with open(os.path.join(xlpp_root, "README.md"), encoding="utf-8") as readme_file:
        long_description = readme_file.read()
except OSError:
    # When built from an sdist the README lives outside the package; never let
    # a missing readme break `pip install` from source.
    long_description = (
        "High-performance C++ Excel xlsx library for Python. "
        "See https://github.com/thang199801666/XLPP for the full README."
    )

setup(
    name="xlpp",
    version="1.0.0",
    author="XL++ contributors",
    description="High-performance C++ Excel xlsx library for Python",
    long_description=long_description,
    long_description_content_type="text/markdown",
    url="https://github.com/thang199801666/XLPP",
    project_urls={
        "Source": "https://github.com/thang199801666/XLPP",
        "Issues": "https://github.com/thang199801666/XLPP/issues",
        "Documentation": "https://thang199801666.github.io/XLPP/",
    },
    license="MIT",
    ext_modules=ext_modules,
    cmdclass={"build_ext": build_ext},
    python_requires=">=3.8",
    classifiers=[
        "Development Status :: 4 - Beta",
        "Intended Audience :: Developers",
        "License :: OSI Approved :: MIT License",
        "Operating System :: OS Independent",
        "Programming Language :: C++",
        "Programming Language :: Python :: 3",
        "Programming Language :: Python :: 3.8",
        "Programming Language :: Python :: 3.9",
        "Programming Language :: Python :: 3.10",
        "Programming Language :: Python :: 3.11",
        "Programming Language :: Python :: 3.12",
        "Programming Language :: Python :: 3.13",
        "Programming Language :: Python :: 3.14",
        "Topic :: Office/Business :: Financial :: Spreadsheet",
    ],
)
