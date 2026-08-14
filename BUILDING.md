# XL++ Bindings — Build Guide

## Prerequisites

- Visual Studio 2022+ (Platform Toolset v145)
- CMake 3.18+ (optional, for Python CMake build)
- Python 3.8+ with `pip`
- .NET SDK 8.0+ (for C#)
- vcpkg (optional, for pybind11 and non-Windows dependency management)
- **Windows:** no external crypto package is required; password-to-open uses CNG/`bcrypt.lib`.
- **Linux/macOS:** OpenSSL Crypto development package is required for password-to-open support.

---

## 1. Build XL++ Core Library

```powershell
# Debug
msbuild XL++.sln /p:Configuration=Debug /p:Platform=x64 /m

# Release (required for bindings/python/C# bindings)
msbuild XL++.sln /p:Configuration=Release /p:Platform=x64 /m
```

Outputs:
- `x64\Release\XLPP.lib` — static library
- `x64\Release\zlib.lib` — compression

---

## Core sanitizer profiles (CMake)

For parser/mutator hardening on Clang or GCC:

```bash
cmake -S . -B build-san -G Ninja \
  -DXLPP_ENABLE_ASAN=ON -DXLPP_ENABLE_UBSAN=ON \
  -DXLPP_BUILD_TESTS=ON
cmake --build build-san --target XLPP_UnitTests
```

MSVC supports the `XLPP_ENABLE_ASAN` profile; UBSan is currently intended for Clang/GCC. These profiles are development verification settings and do not alter the public ABI.


## P1J core verification notes

P1J adds standalone translation units for structural editing, reference transformation, Pivot structural repair, workbook sheet lifecycle operations and model validation. The Visual Studio project includes these files explicitly; CMake discovers them through the existing recursive source glob.

For release verification, build `xlpp_static`, `XLPP_HeaderCheck`, `XLPP_UnitTests`, `XLPP_P1P_LazyFormulaTests`, `XLPP_P1Q_CoreHardeningTests`, `xlpp_capi`, `xlpp_capi_smoke`, the package validator and the three standalone writer examples. `SaveOptions::validateModelBeforeSave` is an opt-in runtime semantic validation gate and has no build-time dependency.

## P1S three-pillar verification

The dedicated CMake target exercises template, generated-chart and typed-Pivot interoperability without growing the legacy unit-test translation unit:

```bash
cmake --build build --target XLPP_P1S_ThreePillarTests
ctest --test-dir build -R XLPP_P1S_ThreePillarTests --output-on-failure
```

For focused ownership/UB checks, build that target under `XLPP_ENABLE_ASAN=ON` and `XLPP_ENABLE_UBSAN=ON`. The P1S development verification passed under Clang 17 with leak detection enabled.

The project package also keeps `XLPP_HeaderCheck`, `xlpp_capi_smoke`, installed `find_package(XLPP CONFIG)` and source `add_subdirectory()` consumer checks. On constrained machines, use the low-optimization verification flags (`-O0 -DNDEBUG -g0`) for the source consumer because the remaining large `Workbook.cpp` still dominates a clean optimized build.

Optional interoperability validation can run the generated P1S artifacts through an external host library. The development run used openpyxl 3.1.5 and verified XLTX/XLTM template identity, native Scatter/Bubble types, Bar+Line combined plots with separate axis IDs and discovery of the generated PivotTable.

## 2. Python Binding

### Install

```powershell
cd bindings\python

# Install pybind11
pip install pybind11 setuptools build

# Build + install
pip install .

Build and publish a wheel (requires a native XLPP Release build and PyPI credentials):

```powershell
python -m pip install build twine
python -m build --wheel
python -m twine check dist/*
python -m twine upload dist/*
```

The repository workflow `.github/workflows/pypi-publish.yml` builds the native
wheel on GitHub Actions and publishes it through the existing `pypi` environment.
No local C++ build is required.
```

### Quick test

```python
import xlpp

wb = xlpp.Workbook()
ws = wb.add_worksheet("Sheet1")

# Values auto-convert between Python and C++ types
ws["A1"].value = 42                    # int → double
ws["B1"].value = "Hello"               # str → string
ws["C1"].value = True                  # bool → bool
ws["D1"].value = None                  # None → empty
import datetime
ws["E1"].value = datetime.date(2024, 1, 15)  # date → DateTime

# Styles (fluent chaining)
ws["A1"].font().bold = True
ws["A1"].font().size = 14
ws["A1"].font().color().set_argb("FFFF0000")

# Merge, freeze, append
ws.merge_cells("A2:C2")
ws.freeze_panes("A3")
ws.append(["Product", "Price", "Qty"])

# Save and reload
wb.save("output.xlsx")

wb2 = xlpp.Workbook()
wb2.load("output.xlsx")
print(wb2["Sheet1"]["A1"].value)
```

### API mapping (openpyxl-like)

| openpyxl | XL++ Python |
|----------|-------------|
| `wb = Workbook()` | `wb = xlpp.Workbook()` |
| `ws = wb.active` | `ws = wb.active` |
| `ws['A1'] = 42` | `ws['A1'].value = 42` |
| `ws['A1'].font.bold = True` | `ws['A1'].font().bold = True` |
| `ws.append([1,2,3])` | `ws.append([1, 2, 3])` |
| `ws.merge_cells('A1:B2')` | `ws.merge_cells('A1:B2')` |
| `wb.save('f.xlsx')` | `wb.save('f.xlsx')` |
| `wb = load_workbook('f.xlsx')` | `wb.load('f.xlsx')` |
| `ws.max_row` | `ws.max_row` |
| `ws.dimensions` | `ws.dimensions` |

### CMake build (alternative)

```powershell
cd bindings\python
mkdir build && cd build
cmake .. -Dpybind11_DIR="C:/path/to/pybind11/share/cmake/pybind11"
cmake --build . --config Release
```

---

## 3. C# Binding

### Build C API DLL

The C# binding requires `xlpp_capi.dll` — a C-compatible wrapper DLL.

```powershell
# Compile the C API wrapper
cl /std:c++20 /EHsc /O2 /MD /DXLPP_CAPI_EXPORTS /DXLPP_STATIC ^
   /I"include" /I"third_party\zlib" ^
   bindings\c\xlpp_capi.cpp ^
   /link x64\Release\XLPP.lib x64\Release\zlib.lib bcrypt.lib ^
   /DLL /OUT:xlpp_capi.dll
```

Or use the Visual Studio Developer Command Prompt:
```powershell
# In Developer PowerShell for VS 2022
cl /std:c++20 /EHsc /O2 /MD /DXLPP_CAPI_EXPORTS /DXLPP_STATIC ^
   /Iinclude /Ithird_party\zlib ^
   bindings\c\xlpp_capi.cpp ^
   /link x64\Release\XLPP.lib x64\Release\zlib.lib bcrypt.lib ^
   /DLL /OUT:xlpp_capi.dll
```

### Use in C# project

1. Copy `xlpp_capi.dll` to your project output directory
2. Add `bindings/cs/XlppNet.cs` to your project
3. Copy `zlib.dll` if needed (zlib is statically linked by default)

```csharp
using XLPP;

class Program
{
    static void Main()
    {
        using var wb = new Workbook();
        var ws = wb.AddWorksheet("Sheet1");

        ws["A1"].Value = 42.5;
        ws["B1"].Value = "Hello from C#";
        ws["C1"].Value = true;

        ws["A1"].Font.SetBold(true);
        ws["A1"].Font.SetSize(16);
        ws["A1"].Font.SetColor("FFFF0000");

        ws.MergeCells("A2:C2");
        ws["A2"].Value = "Merged title";

        wb.Properties.Title = "XL++ C# Demo";
        wb.Properties.Creator = "XL++";

        wb.Save("csharp_output.xlsx");
        Console.WriteLine("Saved!");
    }
}
```

### .csproj example

```xml
<Project Sdk="Microsoft.NET.Sdk">
  <PropertyGroup>
    <OutputType>Exe</OutputType>
    <TargetFramework>net8.0</TargetFramework>
    <PlatformTarget>x64</PlatformTarget>
  </PropertyGroup>
  <ItemGroup>
    <Compile Include="..\..\bindings\cs\XlppNet.cs" />
  </ItemGroup>
  <ItemGroup>
    <None Include="xlpp_capi.dll" CopyToOutputDirectory="PreserveNewest" />
  </ItemGroup>
</Project>
```

---

## 4. Project Structure

```
XLPP/
├── include/XLPP/          # Public headers
├── src/XLPP/              # Implementation
│   ├── Packaging/         # ZIP archive I/O + MappedFile
│   ├── Streaming/         # Streaming read/write + shared strings
│   ├── Threading/         # ThreadPool
│   ├── Workbook/
│   ├── Worksheet/
│   └── XML/               # XML scanning + utilities
├── third_party/zlib/      # zlib compression
├── bindings/              # Language bindings
│   ├── python/            # Python (pybind11)
│   │   ├── setup.py
│   │   ├── CMakeLists.txt
│   │   └── src/xlpp_bindings.cpp
│   ├── c/                 # C API (DLL for P/Invoke)
│   │   ├── xlpp_capi.h
│   │   └── xlpp_capi.cpp
│   └── csharp/            # C# P/Invoke wrapper
│       └── XlppNet.cs
├── tests/                 # C++ unit tests
├── samples/               # C++ sample
├── scripts/
│   └── pgo.ps1            # PGO build script
├── .github/workflows/     # CI/CD
└── BUILDING.md
```


## Crypto dependency notes

P1H password-to-open encryption selects the crypto backend by platform:

- Windows/MSVC: system CNG (`bcrypt.lib`), including direct Visual Studio solution builds.
- Linux/macOS: `find_package(OpenSSL REQUIRED COMPONENTS Crypto)` and exported `find_dependency(OpenSSL COMPONENTS Crypto)`.
- `vcpkg.json` requests OpenSSL only on non-Windows platforms.

The public version remains `1.1.2`; milestone labels such as P1K are development snapshots and do not change the public release version.


### P1H encryption build note

P1H supports Agile AES-128/192/256 with SHA-1/SHA-256/SHA-384/SHA-512 and Standard AES-128/192/256 + SHA-1 without adding new source-level dependencies beyond the existing platform crypto backend. Windows direct Visual Studio builds continue to use system `bcrypt.lib`; non-Windows installed-package consumers inherit `OpenSSL::Crypto` through the exported CMake config.


### P1I encrypted-package memory boundary

P1I does not add a new crypto dependency. Windows continues to use system CNG/BCrypt and non-Windows builds continue to use `OpenSSL::Crypto`. The encrypted file path now serializes the inner OOXML ZIP to memory before creating `EncryptedPackage`, and decrypts directly into the in-memory ZIP parser. This removes the former plaintext inner-workbook temporary file from password-to-open file/stream workflows.

`LoadOptions` also exposes `maxEncryptionInfoBytes`, `allowStandardEncryption`, and `requireAgileDataIntegrity`. The C ABI mirrors these through `xlpp_workbook_load_password_ex`.

### P1K core-hardening build note

P1K adds no new external runtime dependency. The structural-edit transaction, 3-D reference hardening and strict pre-write validation use the existing core/package graph. Public headers remain standalone-compilable.

For release/CI validation, a useful strict save path is:

```cpp
xlpp::SaveOptions options;
options.validateModelBeforeSave = true;
options.rejectModelWarningsBeforeSave = true;
options.validatePackageBeforeWrite = true;
workbook.save("validated.xlsx", options);
```

`StructuralEditOptions::cancel` is cooperative rather than asynchronous: callbacks run at deterministic mutation checkpoints on the calling thread. With the default `rollbackOnFailure=true`, cancellation restores the workbook model before `StructuralEditCancelled` escapes.


## P1O core hot-path benchmark

When benchmarks are enabled, `xlpp_core_hotpath_benchmark [rows]` builds `rows x 10` cells (default 30,000), exercises geometry/change tracking and a Store save, and reports `sizeof_cell_bytes`, `sizeof_style_bytes`, `peak_rss_bytes` and timing lines. `peak_rss_bytes` uses the native process API (`GetProcessMemoryInfo` on Windows, `getrusage` on Unix-like hosts) and is intended as a same-host regression signal rather than a cross-platform absolute comparison.


## P1P cell-density benchmark

With `XLPP_BUILD_BENCHMARKS=ON`, `xlpp_cell_density_benchmark [rows] [default|styled]` builds `rows x 10` cells and reports `sizeof_cell_bytes`, `peak_rss_bytes`, build time and Store-save time. `default` measures the allocation-free default style/formula-metadata path; `styled` applies a non-default style to every cell and is intended to expose the cost of materializing lazy style state. Compare only same-host/same-compiler builds.

P1P also adds `XLPP_P1P_LazyFormulaTests`, a small independent regression executable under `XLPP_BUILD_TESTS=ON`, to verify that package loading does not allocate FormulaMetadata for ordinary formulas while shared/array metadata round-trips correctly.


## P1R core hardening build note

P1R adds `src/XLPP/Workbook/WorkbookIO.cpp` and `WorkbookPackageReader.cpp` as separate core translation units; CMake source discovery and the Visual Studio v145 project include both. No new external dependency is introduced. The materialized reader still leaves `Workbook.cpp` as the dominant optimized/sanitized compile unit, so P1S should continue decomposition before increasing parser scope.

For hostile-input validation, configure Clang/GCC with `-DXLPP_ENABLE_ASAN=ON -DXLPP_ENABLE_UBSAN=ON` and run `XLPP_P1Q_CoreHardeningTests` plus `XLPP_P1R_TransactionalLimitsTests`.

## P1T three-pillar verification

P1T adds one public header, `include/XLPP/Chart/Chartsheet.h`; CMake installs it automatically and the Visual Studio v145 core project lists it explicitly. No new runtime dependency is introduced.

Recommended focused gates after a three-pillar change:

```text
XLPP_UnitTests
XLPP_P1S_ThreePillarTests
XLPP_P1T_ThreePillarTests
xlpp_capi_smoke
xlpp-package-validator
```

For memory/ownership hardening, Clang/GCC can build the P1T test with `XLPP_ENABLE_ASAN=ON` and `XLPP_ENABLE_UBSAN=ON`. `tests/interop/p1t_openpyxl_host_check.py` is an optional independent host gate when openpyxl is installed; generate P1T artifacts with `XLPP_KEEP_P1T_ARTIFACTS=1` and pass the Chartsheet/Pivot paths to the script.


## P1U template / Chartsheet verification

P1U does not add a runtime dependency. It extends existing public headers (`Workbook.h`, `Chartsheet.h`, `PageSetup.h`) and the C ABI wrapper.

Focused regression target:

```bash
cmake --build build --target XLPP_P1U_TemplateChartsheetTests
./build/tests/XLPP_P1U_TemplateChartsheetTests
```

To retain generated XLTX/XLTM artifacts for an independent host check:

```bash
XLPP_KEEP_P1U_ARTIFACTS=1 ./build/tests/XLPP_P1U_TemplateChartsheetTests
python tests/interop/p1u_openpyxl_host_check.py generated.xltx patched-import.xltx generated.xltm
```

Focused sanitizer profile:

```bash
cmake -S . -B build-san -G Ninja \
  -DCMAKE_CXX_COMPILER=clang++ \
  -DXLPP_BUILD_TESTS=ON \
  -DXLPP_ENABLE_ASAN=ON \
  -DXLPP_ENABLE_UBSAN=ON
cmake --build build-san --target XLPP_P1U_TemplateChartsheetTests
ASAN_OPTIONS=detect_leaks=1:halt_on_error=1 \
UBSAN_OPTIONS=halt_on_error=1 \
  ./build-san/tests/XLPP_P1U_TemplateChartsheetTests
```

The installed `find_package` and source `add_subdirectory` consumer fixtures also exercise template identity, active Chartsheet state and hidden mixed tabs.


## P1V template / Chartsheet depth verification

P1V adds no runtime dependency. The new internal translation unit `src/XLPP/Workbook/WorkbookChartsheetIO.cpp` is part of both the CMake source glob and the Visual Studio v145 project.

Focused regression:

```bash
cmake --build build --target XLPP_P1V_ChartsheetDepthTests
./build/tests/XLPP_P1V_ChartsheetDepthTests
```

Generate host-check artifacts and validate with openpyxl when it is installed:

```bash
XLPP_KEEP_P1V_ARTIFACTS=1 ./build/tests/XLPP_P1V_ChartsheetDepthTests
python tests/interop/p1v_openpyxl_chartsheet_depth_check.py \
  generated-advanced.xltx custom-view-mutated.xltx worksheet-pagesetup.xlsx
```

Focused sanitizer profile:

```bash
cmake -S . -B build-san -G Ninja \
  -DCMAKE_CXX_COMPILER=clang++ \
  -DXLPP_BUILD_TESTS=ON \
  -DXLPP_ENABLE_ASAN=ON \
  -DXLPP_ENABLE_UBSAN=ON
cmake --build build-san --target XLPP_P1V_ChartsheetDepthTests
ASAN_OPTIONS=detect_leaks=1:halt_on_error=1 \
UBSAN_OPTIONS=halt_on_error=1 \
  ./build-san/tests/XLPP_P1V_ChartsheetDepthTests
```

The aggregate `PackageConsumerTests` driver can exceed constrained command windows because it launches nested Release builds. P1V release verification therefore also executes the installed `find_package(XLPP CONFIG)` and source `add_subdirectory()` consumer binaries separately.


## P1W Chartsheet auxiliary package verification

Build and run the dedicated ownership regression:

```bash
cmake --build build --target XLPP_P1W_ChartsheetPackageTests
ctest --test-dir build -R XLPP_P1W_ChartsheetPackageTests --output-on-failure
```

The test covers generated/imported/replaced/cleared printer-settings binary parts, `legacyDrawingHF` VML preservation, chart-regeneration closure ownership, repeated saves and Chartsheet removal cleanup. For focused memory/UB checks, build the target with `XLPP_ENABLE_ASAN=ON` and `XLPP_ENABLE_UBSAN=ON`.

Optional independent host validation is provided by:

```bash
python tests/interop/p1w_openpyxl_chartsheet_package_check.py \
  <printer-template.xltx> <header-footer-template.xltx> <regenerated-template.xltx>
```

The installed `find_package(XLPP CONFIG)` and source `add_subdirectory()` consumers also exercise the new printer-settings API. Because the remaining `Workbook.cpp` is still large, verifying those two consumers separately is preferable to treating a timeout of the aggregate nested clean-build harness as a runtime failure.


## P1X-A verification target

`XLPP_P1X_ChartsheetWriterTests` exercises the extracted Chartsheet package writer and relationship-ID collision repair. For sanitizer verification, configure with `XLPP_ENABLE_ASAN=ON` and `XLPP_ENABLE_UBSAN=ON`, then build/run that target. `WorkbookChartsheetPackage.cpp` is listed in the Visual Studio v145 project and is discovered automatically by CMake.

## Visual Studio internal source include path hotfix

The hand-authored `src/XLPP/XLPP.vcxproj` must include `$(SolutionDir)src` in
`AdditionalIncludeDirectories`. Internal implementation headers are intentionally
included as `"XLPP/..."` (for example `XLPP/Worksheet/WorksheetName.h`), matching
the CMake target's private `${CMAKE_CURRENT_SOURCE_DIR}` include root.

For both Debug|x64 and Release|x64 the effective order is:

`$(SolutionDir)src;$(SolutionDir)include;$(SolutionDir)third_party\zlib;%(AdditionalIncludeDirectories)`
