# XL++ Bindings — Build Guide

## Prerequisites

- Visual Studio 2022+ (Platform Toolset v145)
- CMake 3.18+ (optional, for Python CMake build)
- Python 3.8+ with `pip`
- .NET SDK 8.0+ (for C#)
- vcpkg (optional, for pybind11)

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

## 2. Python Binding

### Install

```powershell
cd bindings\python

# Install pybind11
pip install pybind11 setuptools

# Build + install
pip install .
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
   /link x64\Release\XLPP.lib x64\Release\zlib.lib ^
   /DLL /OUT:xlpp_capi.dll
```

Or use the Visual Studio Developer Command Prompt:
```powershell
# In Developer PowerShell for VS 2022
cl /std:c++20 /EHsc /O2 /MD /DXLPP_CAPI_EXPORTS /DXLPP_STATIC ^
   /Iinclude /Ithird_party\zlib ^
   bindings\c\xlpp_capi.cpp ^
   /link x64\Release\XLPP.lib x64\Release\zlib.lib ^
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
