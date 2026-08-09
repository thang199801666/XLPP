"""Self-contained build for the xlpp Python module.

Compiles the XL++ library (src/XLPP/**/*.cpp), the bundled zlib, and the
pybind11 binding into the extension directly. This is what makes the sdist
buildable and lets cibuildwheel produce wheels without pre-built libraries.
"""

import os
import sys

import pybind11
from setuptools import Extension, setup
from setuptools.command.build_ext import build_ext as _build_ext


class build_ext(_build_ext):
    """Build the extension, compiling zlib's C sources with the C compiler.

    The extension mixes .cpp and .c sources. setuptools passes every
    extra_compile_args entry (including -std=c++20) to the C compiler for
    .c files, which clang/gcc reject ("invalid argument '-std=c++20' not
    allowed with 'C'"). Compile the C sources separately first, then hand
    the resulting objects to the normal C++ build so -std=c++20 only ever
    reaches a C++ translation unit.
    """

    def build_extensions(self):
        for ext in self.extensions:
            c_sources = [s for s in ext.sources if s.endswith(".c")]
            if not c_sources:
                continue
            others = [s for s in ext.sources if not s.endswith(".c")]
            c_args = [a for a in ext.extra_compile_args
                      if not a.startswith(("-std", "/std"))]
            objects = self.compiler.compile(
                c_sources,
                output_dir=self.build_temp,
                macros=ext.define_macros or [],
                include_dirs=ext.include_dirs or [],
                debug=self.debug,
                extra_postargs=c_args,
                depends=ext.depends,
            )
            ext.sources = others
            ext.extra_objects = (ext.extra_objects or []) + list(objects)
        super().build_extensions()

root = os.path.abspath(os.path.dirname(__file__))

with open(os.path.join(root, "VERSION"), encoding="utf-8") as version_file:
    package_version = version_file.read().strip()

is_msvc = os.name == "nt"

if is_msvc:
    compile_args = ["/std:c++20", "/EHsc", "/bigobj", "/DXLPP_STATIC"]
    link_args = ["bcrypt.lib"]
else:
    compile_args = ["-std=c++20", "-fPIC", "-DXLPP_STATIC"]
    link_args = []


def xlpp_sources():
    # Relative paths so setuptools includes them in the sdist file list.
    out = []
    for dirpath, _dirnames, filenames in os.walk(os.path.join("src", "XLPP")):
        for name in sorted(filenames):
            if name.endswith(".cpp") and name != "WorkbookCodec.cpp":
                out.append(os.path.join(dirpath, name))
    return out


def zlib_sources():
    zlib_dir = os.path.join("third_party", "zlib")
    out = []
    for name in sorted(os.listdir(zlib_dir)):
        # Exclude the stdio-based gz* layer: XLPP uses the raw inflate/deflate
        # API and gz*.c does not compile as C++ on all toolchains.
        if (name.endswith(".c") and name not in
                ("example.c", "minigzip.c", "gzclose.c", "gzlib.c", "gzread.c", "gzwrite.c")):
            out.append(os.path.join(zlib_dir, name))
    return out


binding = os.path.join("bindings", "python", "src", "xlpp_bindings.cpp")
sources = [binding] + xlpp_sources() + zlib_sources()

ext_modules = [
    Extension(
        "xlpp",
        sources,
        include_dirs=[
            pybind11.get_include(),
            os.path.join(root, "include"),
            os.path.join(root, "src", "XLPP"),
            os.path.join(root, "third_party", "zlib"),
        ],
        extra_compile_args=compile_args,
        extra_link_args=link_args,
        language="c++",
    ),
]

try:
    with open(os.path.join(root, "README.md"), encoding="utf-8") as f:
        long_description = f.read()
except OSError:
    long_description = (
        "High-performance C++ Excel xlsx library for Python. "
        "See https://github.com/thang199801666/XLPP for the full README."
    )

setup(
    name="xlpp",
    version=package_version,
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
