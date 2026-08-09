# XL++ .NET SDK

`XLPP` is the .NET wrapper for the native XL++ C ABI. Native binaries are resolved from `runtimes/<rid>/native` in the NuGet package.

```csharp
using var workbook = new XLPP.Workbook();
var sheet = workbook.AddWorksheet("Report");
sheet.Cell("A1").Value = 42;
workbook.Save("report.xlsx");
```

The owning `Workbook` uses `SafeHandle`; child handles remain non-owning views and must not outlive their workbook.
