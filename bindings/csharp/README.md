# XL++ .NET SDK

`XLPP` is the .NET wrapper for the native XL++ C ABI. Native binaries are resolved from `runtimes/<rid>/native` in the NuGet package.

```csharp
using var workbook = new XLPP.Workbook();
var sheet = workbook.AddWorksheet("Report");
sheet.Cell("A1").Value = 42;
workbook.Save("report.xlsx");
```

The owning `Workbook` uses `SafeHandle`; child handles remain non-owning views and must not outlive their workbook.

The managed API includes path and memory I/O, streaming, formulas and dependency
graphs, styles, tables, charts, PivotTables, VBA, encryption, rich text,
validation, structural edits, preservation metadata, and enterprise feature
inspection. Enterprise payloads such as Power Query, Data Model/OLAP, slicers,
timelines, SmartArt, ActiveX, and UserForms remain preservation-oriented where
the native library does not provide semantic authoring.
