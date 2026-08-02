# Quick Start — C\#

## Install

1. Build `xlpp_capi.dll`:

```powershell
msbuild bindings\XLPP.CApi.vcxproj /p:Configuration=Release /p:Platform=x64 /m
```

2. Copy `XlppNet.cs` into your project, or reference the DLL directly.

## Usage

```csharp
using XLPP;

class Program
{
    static void Main()
    {
        // Create
        using var wb = new Workbook();
        var ws = wb.AddWorksheet("Report");

        // Write cells
        ws["A1"].Value = "Product";
        ws["B1"].Value = 42.5;
        ws["C1"].Value = true;
        ws["D1"].Value = "Hello from C#";

        // Style
        ws["A1"].Font.SetBold(true);
        ws["A1"].Font.SetSize(16);
        ws["A1"].Font.SetColor("FFFF0000");

        // Borders
        ws["A1"].Border.Bottom.SetStyle("thin");
        ws["A1"].Border.Left.SetStyle("thin");

        // Alignment
        ws["A1"].Alignment.SetHorizontal("center");

        // Merge
        ws.MergeCells("A2:C2");
        ws["A2"].Value = "Merged Title";

        // Freeze
        ws.FreezePanes("A3");

        // Metadata
        wb.Properties.Title = "C# Demo";
        wb.Properties.Creator = "XL++";

        // Save
        wb.Save("csharp_report.xlsx");

        // Load
        using var loaded = new Workbook();
        loaded.Load("csharp_report.xlsx");
        var loadedWs = loaded["Report"];
        Console.WriteLine(loadedWs["A1"].Value); // "Product"
        Console.WriteLine(loadedWs["B1"].Value); // 42.5

        // Multi-sheet
        loaded.AddWorksheet("Sheet2");
        Console.WriteLine(loaded.SheetCount);    // 2

        // Cell access
        Console.WriteLine(loadedWs["A1"].ValueType); // String
        Console.WriteLine(loadedWs.MaxRow);          // rows
        Console.WriteLine(loadedWs.MaxColumn);       // cols
    }
}
```

## Project file

```xml
<Project Sdk="Microsoft.NET.Sdk">
  <PropertyGroup>
    <OutputType>Exe</OutputType>
    <TargetFramework>net8.0</TargetFramework>
    <PlatformTarget>x64</PlatformTarget>
  </PropertyGroup>
  <ItemGroup>
    <Compile Include="XlppNet.cs" />
  </ItemGroup>
  <ItemGroup>
    <None Include="xlpp_capi.dll" CopyToOutputDirectory="PreserveNewest" />
  </ItemGroup>
</Project>
```

## Key API

```csharp
// Workbook
using var wb = new Workbook();           // create (IDisposable)
var ws = wb.AddWorksheet("Sheet1");      // add sheet
wb.Load("input.xlsx");                   // read
wb.Save("output.xlsx");                  // write
ws = wb["Sheet1"];                       // by name
ws = wb[0];                              // by index
int count = wb.SheetCount;               // sheet count
bool ok = wb.RemoveWorksheet("Sheet1");  // remove
wb.Properties.Title = "Title";           // metadata

// Cell
ws["A1"].Value = "text";                 // string
ws["A1"].Value = 42.5;                   // double
ws["A1"].Value = true;                   // bool
object? v = ws["A1"].Value;              // get (auto-typed)
ws["A1"].Formula = "=B1*2";              // formula
ws.Cell(1, 2).Value = "data";            // row, col access
bool exists = ws.HasCell("Z99");         // check

// Style
ws["A1"].Font.SetBold(true);
ws["A1"].Font.SetSize(14);
ws["A1"].Font.SetItalic(true);
ws["A1"].Font.SetColor("FFFF0000");
ws["A1"].Border.Bottom.SetStyle("thin");
ws["A1"].Border.Top.SetStyle("medium");
ws["A1"].Alignment.SetHorizontal("center");

// Layout
ws.MergeCells("A1:C1");
ws.UnmergeCells("A1:C1");
ws.FreezePanes("B2");
```

## Type mapping

| C# | XL++ |
|----|------|
| `string` | `CellValue{string}` |
| `double` | `CellValue{double}` |
| `int`, `long`, `float` | → `double` |
| `bool` | `CellValue{bool}` |
| `null` | (not set) |
