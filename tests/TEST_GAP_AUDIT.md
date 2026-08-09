# XL++ Test Gap Audit — VBA Text and Reader Coverage

Updated: **2026-08-07**

## Request addressed

This update addresses two test gaps:

1. Generate a macro-enabled workbook by supplying VBA source code as text.
2. Add genuine reader tests because many earlier tests primarily exercised writing and then inspected the generated package.

## API added

```cpp
xlpp::Workbook workbook;
workbook.addWorksheet("MacroWorkbook");
workbook.setVbaModuleText(
    "XLPPGenerated",
    "Option Explicit\n"
    "Public Sub Hello()\n"
    "    MsgBox \"Hello from XL++\"\n"
    "End Sub\n");
workbook.save("macro.xlsm");

xlpp::Workbook loaded;
loaded.load("macro.xlsm");
auto source = loaded.vbaModuleText("XLPPGenerated");
```

Public operations now include:

- `setVbaModuleText(name, source)`
- `vbaModuleText(name)`
- `vbaModules()`
- `removeVbaModule(name)`

The original `addVbaProject()` and `setVbaProject()` APIs remain available for attaching an existing binary project.

## VBA project implementation

XL++ now builds a compact MS-CFB version-3 compound file containing:

- root `PROJECT` and `PROJECTwm` streams;
- `VBA/_VBA_PROJECT`;
- compressed `VBA/dir` project metadata;
- compressed source streams for `Sheet1`, `Sheet2`, ..., `ThisWorkbook`, and each standard module;
- a valid red-black directory sibling tree, FAT, miniFAT and mini stream.

Source is normalized to CRLF and stored in OVBA compressed containers. This is source packaging, not a VBA language compiler: syntax is not validated and forms/designer data/signatures are outside the current scope.

## Independent reader fixtures added

| New suite | Package origin | Reader behavior verified |
|---|---|---|
| `External OOXML cell and style reader fixture` | Handcrafted ZIP/XML | Cell types, formula cache, errors, rich text, dates, font/fill/border/alignment |
| `External OOXML worksheet feature reader fixture` | Handcrafted ZIP/XML | Dimensions, merges, views, protection, filters, CF/DXF, validation, links, comments, printing |
| `External OOXML workbook metadata reader fixture` | Handcrafted ZIP/XML | Date system, properties, protection, calculation settings, names, print area/titles |
| `VBA source text build and read` | CFB/OVBA builder plus package loader | Text-to-project, save/load source extraction, update/add/remove lifecycle |

Because the OOXML reader fixtures bypass the XL++ writer, a serializer bug cannot make the fixture and reader agree accidentally.

## Result

- Debug: **123/123 suites, 1,476 checks PASS**
- Release/NDEBUG: **123/123 suites, 1,476 checks PASS**
- Public headers: **56/56 compile independently**
- Example sources: **43/43 compile**
- `.cpp` line coverage: **92.86%**
- `.cpp` branch coverage: **73.31%**
- `.cpp` function coverage: **99.09%**

## External validation performed

- `37_vba_text.xlsm` is recognized as an Excel 2007+ package.
- Extracted `vbaProject.bin` is independently recognized as a CFB V2 compound document.
- LibreOffice opens the generated `.xlsm` and exports its worksheet to PDF.
- XL++ reloads the package and extracts the original module text.
- An independent CFB/OVBA profile parser discovers `XLPPGenerated.XLPP_Hello` as a public parameterless procedure.
- Worksheet document modules and worksheet `codeName` values stay synchronized even when sheets are added after VBA generation or after reload.

Macro execution itself still requires Microsoft Excel Desktop and an environment where macros are enabled.
