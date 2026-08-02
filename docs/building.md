# Building from Source

## Prerequisites

| Platform | Compiler | Build System |
|----------|----------|-------------|
| Windows | MSVC 2022 (Toolset v145) | MSBuild or CMake + Ninja |
| Linux | GCC 14+ | CMake + Ninja |
| macOS | Clang 18+ | CMake + Ninja |

zlib is bundled (`third_party/zlib/`). Set `-DXLPP_USE_BUNDLED_ZLIB=OFF` to use system zlib.

## Quick build

### Windows (MSBuild)

```powershell
msbuild XL++.sln /p:Configuration=Debug /p:Platform=x64 /m
msbuild XL++.sln /p:Configuration=Release /p:Platform=x64 /m
```

### Cross-platform (CMake)

```bash
cmake -B build -G Ninja -DCMAKE_BUILD_TYPE=Release
cmake --build build
# Run tests
cd build && ./tests/XLPP_UnitTests
```

## CMake options

| Option | Default | Description |
|--------|---------|-------------|
| `XLPP_BUILD_TESTS` | ON | Build unit tests |
| `XLPP_BUILD_SAMPLES` | ON | Build sample programs |
| `XLPP_BUILD_CAPI` | ON | Build C API DLL |
| `XLPP_USE_BUNDLED_ZLIB` | ON | Use bundled zlib |

## Build Python bindings

```bash
cd bindings/python
pip install pybind11 setuptools
pip install .
```

## Build C API DLL

```powershell
msbuild bindings\XLPP.CApi.vcxproj /p:Configuration=Release /p:Platform=x64 /m
```

Output: `bindings/x64/Release/XLPP.CApi.dll`

## Install (CMake)

```bash
cmake --install build --prefix /usr/local
# or
cmake --install build --prefix C:/xlpp
```

Then in your CMakeLists.txt:
```cmake
find_package(XLPP REQUIRED)
target_link_libraries(myapp XLPP::xlpp)
```
