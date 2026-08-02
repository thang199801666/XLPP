# Quick Start — C++

## Requirements

- C++20 compiler (MSVC 2022+, GCC 14+, Clang 18+)
- CMake 3.18+ (recommended) or Visual Studio 2022
- zlib (bundled or system)

## Build

### CMake (recommended)

```bash
cmake -B build -G Ninja -DCMAKE_BUILD_TYPE=Release
cmake --build build
```

Options:
- `-DXLPP_BUILD_TESTS=OFF` — skip unit tests
- `-DXLPP_BUILD_CAPI=OFF` — skip C API DLL
- `-DXLPP_USE_BUNDLED_ZLIB=OFF` — use system zlib

### Visual Studio

```powershell
msbuild XL++.sln /p:Configuration=Release /p:Platform=x64 /m
```

## Usage

```cpp
#include <XLPP/XLPP.h>
using namespace xlpp;

int main() {
    // Create a workbook with one sheet
    Workbook wb;
    auto& ws = wb.addWorksheet("Report");

    // Write data
    ws.cell("A1").setValue("Item");
    ws.cell("B1").setValue("Quantity");
    ws.cell("C1").setValue("Price");

    // Style the header
    ws.cell("A1").font().setBold(true);
    ws.cell("A1").font().setSize(14);
    ws.cell("A1").fill().setPatternType("solid");
    ws.cell("A1").fill().foregroundColor().setArgb("FF4472C4");
    ws.cell("A1").font().color().setArgb("FFFFFFFF");

    // Number formatting
    ws.cell("C2").setNumberFormat("$#,##0.00");

    // Append rows
    ws.append({"Widget", 100, 9.99});
    ws.append({"Gadget", 50, 19.99});
    ws.append({"Sprocket", 200, 4.99});

    // Merge and freeze
    ws.mergeCells("A5:C5");
    ws.cell("A5").setValue("Total: $3,497.00");
    ws.freezePanes("A2");

    // Set column widths
    ws.columnDimension(1).width = 20;
    ws.columnDimension(2).width = 12;
    ws.columnDimension(3).width = 12;

    // Document properties
    wb.properties().setTitle("Sales Report");
    wb.properties().setCreator("XL++");

    // Save
    wb.save("report.xlsx");

    // Load and verify
    Workbook loaded;
    loaded.load("report.xlsx");
    auto& sheet = loaded["Report"];
    std::cout << sheet.cell("A1").stringValueOr("") << std::endl; // "Item"
    std::cout << sheet.cell("B2").numericValueOr(0) << std::endl;  // 100

    return 0;
}
```

## Key API

### Workbook

```cpp
Workbook wb;
wb.addWorksheet("Sheet1");           // add sheet
wb.load("input.xlsx");               // read file
wb.load("input.xlsx", LoadOptions{}); // with options
wb.save("output.xlsx");              // write file
wb.save("output.xlsx", SaveOptions{});// with options
wb.worksheet("Sheet1");              // find by name
wb[0];                               // index access
wb.sheetNames();                     // list names
wb.properties().setTitle("...");     // metadata

// Parallel save options
SaveOptions opt;
opt.parallelWorkers = 4;             // use 4 threads
opt.parallelSheets = true;           // parallel per sheet
opt.parallelRows = true;             // parallel per row
opt.compressionLevel = CompressionLevel::Best;
```

### Worksheet

```cpp
auto& ws = wb["Sheet1"];
ws.cell("A1");                       // by address
ws.cell(1, 2);                       // by (row, col)
ws.tryCell("Z99");                   // safe lookup
ws.cell("A1").setValue(42);          // set value
ws.cell("A1").setFormula("B1*2");    // formula
ws.append({1, 2, 3});                // append row
ws.mergeCells("A1:C1");              // merge
ws.freezePanes("B2");                // freeze at B2
ws.insertRows(2, 3);                 // insert 3 rows at row 2
ws.dimensions();                     // "A1:C10"
ws.maxRow();                         // 10
ws.empty();                          // false
```

### Cell styling

```cpp
auto& cell = ws.cell("A1");
cell.font().setBold(true);
cell.font().setSize(14);
cell.font().setItalic(true);
cell.font().color().setArgb("FFFF0000");          // red
cell.fill().setPatternType("solid");
cell.fill().foregroundColor().setArgb("FFCCFFCC"); // light green
cell.border().bottom().setStyle("thin");
cell.border().bottom().color().setArgb("FF000000");
cell.alignment().setHorizontal("center");
cell.alignment().setWrapText(true);
cell.setNumberFormat("#,##0.00");
```

### Cell values

```cpp
cell.setValue(42);                   // double
cell.setValue("text");               // string
cell.setValue(true);                 // boolean
cell.setValue(DateTime{2024,1,15});  // date
cell.setError(CellError::NA);        // #N/A error
cell.setValue(std::monostate{});     // empty/clear

cell.isNumeric();                    // value type checks
cell.numericValueOr(0.0);            // safe getter with fallback
cell.stringValueOr("n/a");
```

### Streaming (large files)

```cpp
// Write 100K+ rows without loading into memory
StreamingWorkbookWriter writer("big.xlsx");
auto& streamWs = writer.addWorksheet("Data");
for (int i = 0; i < 100000; ++i)
    streamWs.append({"Row" + std::to_string(i), 1.0, 2.0});
writer.close();

// Read 100K+ rows with callback
StreamingWorkbookReader reader("big.xlsx");
reader.forEachRow("Data", [](const RowData& row) {
    // process each row
});
```

### Filters, tables, validation

```cpp
auto& filter = ws.autoFilter();
filter.setReference("A1:C100");
auto& col = filter.column(0);
col.addValue("Open");

ws.addTable("Sales", "A1:C10");

ws.dataValidations().add(
    DataValidation::list("Lookup!$A$1:$A$10", "B2:B100")
);

ws.conditionalFormatting().add("A2:A100",
    ConditionalRule::cellIs(ConditionalOp::GreaterThan, {"10"})
);
```
