# XL++ API Reference

XL++ is a dependency-light C++20 `.xlsx` read/write library inspired by openpyxl.
It uses the C++ standard library and zlib only.

```cpp
#include <XLPP/XLPP.h>
```

---

## Workbook

The `Workbook` represents an entire spreadsheet document.

### Creating and loading

```cpp
xlpp::Workbook wb;                          // empty workbook
xlpp::Workbook wb2;
wb2.load("input.xlsx");                     // load from file
wb2.load("input.xlsx", LoadOptions{});      // load with options
std::ifstream in("data.xlsx", std::ios::binary);
wb2.load(in);                               // load from stream
```

### Saving

```cpp
wb.save("output.xlsx");                     // save to file
wb.save("output.xlsx", SaveOptions{});      // save with options
std::ofstream out("data.xlsx", std::ios::binary);
wb.save(out);                               // save to stream
```

### Worksheets

```cpp
auto& sheet = wb.addWorksheet("Sheet1");    // create sheet
wb.removeWorksheet("Sheet1");               // remove sheet, returns bool
wb.copyWorksheet(sheet, "Copy");            // copy existing sheet

auto* s = wb.worksheet("Sheet1");           // find by name, nullptr if missing
auto& s2 = wb[0];                           // access by index (0-based)
std::size_t idx = wb.index(s2);             // get index of a sheet
auto names = wb.sheetNames();               // {"Sheet1", "Copy", ...}
std::size_t count = wb.sheetCount();        // number of sheets
```

### Styles

```cpp
wb.addNamedStyle({"Accent", Style{}});                     // register named style
auto* ns = wb.namedStyle("Accent");                        // find by name
wb.applyNamedStyle(sheet.cell("A1"), "Accent");            // apply to cell
```

### Defined Names

```cpp
wb.addDefinedName({"Revenue", "'Sales'!$A$1:$A$100"});     // workbook-scoped name
auto* dn = wb.definedName("Revenue");
wb.addDefinedName({"SalesData", "'Sales'!$B$1:$B$50", 0}); // sheet-scoped (localSheetId=0)
```

### Properties and protection

```cpp
wb.properties().setTitle("Quarterly Report");
wb.properties().setCreator("XL++");
wb.protection().setLockStructure(true);
wb.setDate1904(false);  // Excel 1900 date system (default)
```

### Load options

```cpp
xlpp::LoadOptions opts;
opts.lenient = true;                    // continue past malformed sheets
opts.maxEntries = 1000;                 // limit ZIP entries
opts.maxEntryBytes = 10 * 1024 * 1024;  // limit entry size
opts.maxTotalBytes = 100 * 1024 * 1024; // limit total payload
opts.cancel = []{ return someFlag; };   // cancellation callback
opts.progress = [](auto done, auto total) { ... }; // progress callback
wb.load("large.xlsx", opts);
```

### Save options

```cpp
xlpp::SaveOptions opts;
opts.strictNamespace = true;                    // ISO 29500 strict OOXML
opts.parallelSheets = true;                     // parallel sheet serialization
opts.parallelWorkers = 4;                       // thread count
opts.compressionLevel = CompressionLevel::Best; // zlib level
opts.compressionStrategy = CompressionStrategy::HuffmanOnly;
wb.save("output.xlsx", opts);
```

### Diagnostics and inspection

```cpp
wb.clear();                                   // reset everything
bool strict = wb.strictNamespaces();          // was source file strict?
const auto& diag = wb.diagnostics();          // load diagnostics
for (const auto& err : diag.errors) { ... }   // error list
for (const auto& part : wb.preservedParts())  // unmodeled parts
    std::cout << part.name << '\n';
```

---

## Worksheet

### Cell access

```cpp
auto& cell = sheet.cell("A1");           // by address string
auto& cell2 = sheet.cell(2, 3);          // by row, column (1-based)
const auto* ptr = sheet.tryCell("A1");   // safe lookup, nullptr if missing
const auto* ptr2 = sheet.tryCell(2, 3);
```

### Ranges

```cpp
auto rng = sheet.range("A1:C3");          // range by address
auto rng2 = sheet.range(1, 1, 3, 3);      // range by coordinates
rng.setValue(42.0);                       // fill all cells
rng.clear();                              // clear all cells
rng.forEach([](Cell& c) { ... });         // visit every cell
auto vals = rng.values();                 // flat vector of values
auto fmts = rng.formulas();               // flat vector of formulas
```

### Iteration

```cpp
sheet.append({"Name", "Value"});   // append row of values at end

auto rows = sheet.rows();          // vector of Row proxies
for (auto& r : rows) {
    std::cout << "Row " << r.number() << ": ";
    for (auto* c : r.cells())
        std::cout << c->address() << ' ';
}

// iterRows / iterCols: extract rectangular blocks of values
// 0 = use sheet extents for that bound
auto block = sheet.iterRows(1, 10, 1, 5);   // rows 1-10, cols 1-5
auto cols  = sheet.iterCols(1, 5, 1, 5);    // cols 1-5, rows 1-5
```

### Dimensions

```cpp
sheet.dimensions();            // "A1:D10" string
sheet.maxRow();                // maximum populated row
sheet.maxColumn();             // maximum populated column
sheet.rowCount();              // alias for maxRow()
sheet.columnCount();           // alias for maxColumn()
auto e = sheet.extents();      // {minRow, minColumn, maxRow, maxColumn}
```

### Layout

```cpp
sheet.mergeCells("A1:C3");                   // merge range
sheet.unmergeCells("A1:C3");                 // unmerge
bool merged = sheet.isMerged("B2");          // check
const auto& ranges = sheet.mergedRanges();   // vector of "A1:C3" strings

sheet.freezePanes("B3");                     // freeze at B3
sheet.clearFreezePanes();
auto pane = sheet.frozenPane();              // optional<string>

auto& rd = sheet.rowDimension(1);
rd.height = 28.5;                            // set row height
rd.hidden = true;                            // hide row

auto& cd = sheet.columnDimension("B");
cd.width = 22.25;                            // set column width
cd.hidden = true;                            // hide column
cd.bestFit = true;                           // auto-fit
```

### Structural edits

```cpp
sheet.insertRows(2, 3);      // insert 3 rows at row 2
sheet.deleteRows(2, 3);      // delete 3 rows starting row 2
sheet.insertColumns(2, 1);   // insert 1 column at column 2
sheet.deleteColumns(3, 2);   // delete 2 columns starting column 3
```

### AutoFilter

```cpp
auto& af = sheet.autoFilter();
af.setReference("A1:D100");
auto& col = af.column(0);      // column index 0
col.addValue("Open");
col.addCustomFilter(FilterOperator::GreaterThan, "100");
col.setAndMode(true);

auto& sort = af.sortState();
sort.setReference("A2:D100");
sort.addCondition("B2:B100", true);  // descending
```

### Conditional formatting

```cpp
auto rule = ConditionalRule::cellIs(ConditionalOperator::LessThan, "0");
rule.setPriority(1);
rule.setStopIfTrue(true);
rule.differentialStyle().font().color().setArgb("FFFF0000");
rule.differentialStyle().fill().setPatternType("solid");
rule.differentialStyle().fill().foregroundColor().setArgb("FFFFC7CE");
sheet.conditionalFormatting().addRule("B2:B100", std::move(rule));

auto formulaRule = ConditionalRule::formula("MOD(A2,2)=0");
sheet.conditionalFormatting().addRule("A2:A100", std::move(formulaRule));
```

### Data validation

```cpp
auto dv = DataValidation::list("B2:B100", "\"Open,Closed,Pending\"");
dv.setAllowBlank(true);
dv.setShowDropDown(true);
dv.setShowInputMessage(true);
dv.setPromptTitle("Select status");
dv.setShowErrorMessage(true);
dv.setErrorTitle("Invalid");
dv.setError("Must be Open, Closed, or Pending.");
sheet.dataValidations().add(std::move(dv));
```

### Tables

```cpp
auto& tbl = sheet.addTable("SalesTable", "A1:C100");
tbl.addColumn("Product");
tbl.addColumn("Amount");
tbl.styleInfo().setName("TableStyleMedium9");
tbl.styleInfo().setShowFirstColumn(true);

auto* found = sheet.table("SalesTable");
```

### Page setup and printing

```cpp
sheet.pageSetup().setOrientation(PageOrientation::Landscape);
sheet.pageSetup().setPaperSize(PaperSize::A4);
sheet.pageSetup().setScale(85);
sheet.pageSetup().setFitToPage(true);
sheet.pageSetup().setFitToWidth(1);
sheet.pageSetup().setFitToHeight(1);
sheet.pageSetup().setBlackAndWhite(true);

sheet.pageMargins().setLeft(0.25);
sheet.pageMargins().setTop(0.75);

sheet.printOptions().setGridLines(true);
sheet.printOptions().setHorizontalCentered(true);

sheet.headerFooter().setOddHeader("&LXL++ Report&R&P");
```

### Protection

```cpp
sheet.protection().setEnabled(true);
sheet.protection().setPasswordHash("ABCD");
sheet.protection().setSelectLockedCells(false);
sheet.protection().setFormatCells(false);
sheet.protection().setSort(true);
sheet.protection().setAutoFilter(true);
```

### Images

```cpp
sheet.addImage("logo.png", "A1");  // from file path
auto& img = sheet.addImage(Image::fromFile("photo.jpg", "B2"));
img.widthPixels();   // pixel dimensions
img.heightPixels();
```

---

## Cell

### Values

```cpp
Cell& cell = sheet.cell("A1");
cell.setValue(3.14);                              // double
cell.setValue(true);                               // bool
cell.setValue(std::string("text"));                // string
cell.setValue("C string");                         // const char*
cell.setValue(xlpp::DateTime{2024, 1, 15});       // date
cell.setError(xlpp::CellError::DivisionByZero);   // error

// Strongly-typed setters
cell.setNumericValue(42.0);
cell.setStringValue("hello");
cell.setBoolValue(true);

// Reading
const auto& v = cell.value();                    // CellValue variant
auto num = std::get<double>(v);                  // extract by type
auto str = std::get<std::string>(v);

// Type checking
cell.hasValue();       // has any non-empty value
cell.isNumeric();      // holds double
cell.isString();        // holds string
cell.isBoolean();       // holds bool
cell.isDate();          // holds DateTime
cell.isError();         // holds CellError
cell.empty();           // no value and no formula

cell.valueType();       // "empty"/"numeric"/"string"/"bool"/"error"/"date"

// Safe extraction with defaults
cell.numericValueOr(0.0);     // returns value or fallback
cell.stringValueOr("n/a");    // returns value or fallback

// Date helpers
cell.setDate(2024, 1, 15);         // sets DateTime + "yyyy-mm-dd" format
cell.setDateTime(DateTime{...});    // sets DateTime + "yyyy-mm-dd h:mm:ss" format
auto d = cell.date();               // optional<DateTime>
```

### Styles

```cpp
cell.font().setName("Arial");
cell.font().setSize(14);
cell.font().setBold(true);
cell.font().setItalic(true);
cell.font().setUnderline(true);
cell.font().setStrike(true);
cell.font().color().setArgb("FFFF0000");      // red

cell.fill().setPatternType("solid");
cell.fill().foregroundColor().setArgb("FFFFFF00");  // yellow

cell.border().left().setStyle("thin");
cell.border().bottom().setStyle("double");
cell.border().bottom().color().setArgb("FF000000");

cell.alignment().setHorizontal("center");
cell.alignment().setVertical("center");
cell.alignment().setWrapText(true);
cell.alignment().setTextRotation(30);

cell.setNumberFormat("#,##0.00");
cell.numberFormat();      // get current format string
cell.setRawStyleIndex(5); // set raw cellXf index
auto idx = cell.styleIndex();  // optional<size_t>
```

### Formulas

```cpp
cell.setFormula("B1+C1");
cell.setSharedFormula("SUM(A1:A10)", 0, "A1:A10");
cell.setArrayFormula("TRANSPOSE(B1:C2)", "D1:E2");

cell.hasFormula();
cell.formula();            // formula text
cell.formulaMetadata().type();     // FormulaType enum
cell.formulaMetadata().reference();
cell.formulaMetadata().sharedIndex();
cell.clearFormula();
```

### Hyperlinks

```cpp
cell.setHyperlink(Hyperlink("https://example.com"));
cell.hyperlink().setDisplay("Click here");
cell.hyperlink().setTooltip("Go to example");
cell.hasHyperlink();
cell.clearHyperlink();
```

### Comments

```cpp
cell.setComment(Comment("Review this value", "Alice"));
cell.hasComment();
auto& c = cell.comment();
c.text();
c.author();
cell.clearComment();
```

### Named styles

```cpp
wb.applyNamedStyle(cell, "Header");
auto name = cell.namedStyle();  // optional<string>
cell.clear();                    // also resets namedStyle
```

### Address navigation

```cpp
cell.address();       // "A1"
cell.row();           // 1
cell.column();        // 1
auto ref = cell.offset(1, 2);   // CellReference{row+1, col+2}
```

---

## CellReference

```cpp
auto ref = CellReference::parse("A1");
auto ref2 = CellReference::parse("$AA$42");  // dollars ignored
ref.row;                                      // 1
ref.column;                                   // 1
ref.address();                                // "A1"

CellReference::columnName(27);                // "AA"
CellReference::columnIndex("AA");             // 27
```

---

## CellValue (variant)

```cpp
using CellValue = std::variant<
    std::monostate,  // empty
    double,          // number
    bool,            // boolean
    std::string,     // text
    CellError,       // error
    DateTime         // date/time
>;
```

---

## DateTime

```cpp
xlpp::DateTime dt{2024, 1, 15, 13, 30, 0};
auto serial = xlpp::toExcelSerial(dt);          // Excel serial number
auto parsed = xlpp::fromExcelSerial(45306.0);   // serial back to DateTime
auto iso = xlpp::toIso8601(dt);                 // "2024-01-15T13:30:00"
auto p = xlpp::parseIso8601("2024-01-15T13:30+05:00");  // UTC adjusted
bool date = xlpp::isDateFormatCode("yyyy-mm-dd");        // is date format?
bool date2 = xlpp::isDateFormatCode("General", 14);      // built-in id 14
```

---

## FormulaMetadata

```cpp
enum class FormulaType { Normal, Shared, Array, DataTable };
cell.formulaMetadata().type() == FormulaType::Shared;
cell.formulaMetadata().sharedIndex();           // optional<unsigned>
cell.formulaMetadata().reference();             // ref string
cell.formulaMetadata().alwaysCalculateArray();
cell.formulaMetadata().calculateOnLoad();
```

---

## CellError

```cpp
enum class CellError {
    Null, DivisionByZero, Value, Reference,
    Name, Number, NotAvailable, GettingData
};
xlpp::toString(CellError::DivisionByZero);  // "#DIV/0!"
xlpp::cellErrorFromString("#N/A");          // CellError::NotAvailable
```

---

## CellRange

```cpp
auto rng = sheet.range("A1:C3");
rng.minRow();    rng.maxRow();
rng.minColumn(); rng.maxColumn();
rng.rowCount();  rng.columnCount();
rng.cell(1, 2);  // cell at (1,2) within range (relative)
rng.address();   // "A1:C3"
rng.setValue(42.0);
rng.clear();
rng.values();    // flat CellValue vector
rng.formulas();  // flat string vector
rng.forEach([](Cell& c) { c.setValue(0.0); });
```

---

## Row

```cpp
auto row = sheet.row(1);
row.number();                     // 1
auto& c = row.cell(3);            // column C
const auto* cp = row.tryCell(3);  // nullptr if not created
auto cells = row.cells();         // vector<Cell*> (non-empty only)
auto values = row.values();       // vector<CellValue> (all cols in extents)
```

---

## Named Style

```cpp
wb.addNamedStyle({"MyStyle", Style{}});
auto* s = wb.namedStyle("MyStyle");   // find
s->name();    s->style();
s->style().font().setBold(true);
s->style().numberFormat() == "#,##0.00";
```

---

## Defined Name

```cpp
wb.addDefinedName({"Total", "'Data'!$C$100"});
auto* dn = wb.definedName("Total");
dn->name();           // "Total"
dn->value();          // "'Data'!$C$100"
dn->comment();        // optional comment
dn->hidden();
dn->localSheetId();   // optional<size_t> for sheet-scoped names
```

---

## Streaming API

For memory-efficient processing of large files without loading the
entire workbook into memory.

### Writing

```cpp
xlpp::StreamingWorkbookWriter writer("large.xlsx",
    SharedStringMode::Hash);               // deduplicate strings
auto& sheet = writer.addWorksheet("Data");
for (int i = 0; i < 100000; ++i)
    sheet.append({"row-" + std::to_string(i), static_cast<double>(i), true});
writer.close();  // must call to finalize the package

// Options
writer.setDate1904(false);
writer.setCompressionLevel(CompressionLevel::Fastest);
writer.setCompressionStrategy(CompressionStrategy::HuffmanOnly);
writer.setParallelWorkers(4);
```

### Reading

```cpp
xlpp::StreamingWorkbookReader reader("large.xlsx");
auto names = reader.worksheetNames();    // all sheet names
auto logSheet = reader.worksheet("Log");

// Range-for iteration
for (auto it = logSheet.begin(); it != logSheet.end(); ++it) {
    std::cout << "Row " << it.rowNumber() << "\n";
    for (const auto& cell : *it) {
        std::cout << "  " << cell.address << " = ";
        if (std::holds_alternative<double>(cell.value))
            std::cout << std::get<double>(cell.value);
        std::cout << "\n";
    }
}

// Callback API
reader.forEachRow("Log", [](std::size_t rowNum, const StreamingRow& row) {
    // return false to stop early
    return true;
});
```

### StreamingCell

```cpp
struct StreamingCell {
    std::string address;
    CellValue value;
    std::string formula;
    std::optional<std::size_t> styleIndex;  // cellXf index (s attribute)
};
```

---

## Package Preservation

XL++ preserves parts it does not model (VBA, charts, custom XML) across
load/save round-trips.

```cpp
wb.load("with_macros.xlsm");
// ... edit sheets ...
wb.save("edited.xlsm");  // macros intact

const auto& parts = wb.preservedParts();
for (const auto& p : parts) {
    p.name;         // package path, e.g. "xl/vbaProject.bin"
    p.data;         // raw bytes
    p.overrideType; // content type override
    p.extension;    // file extension
    p.compress;     // true for non-image parts
}
```

---

## Document Properties

```cpp
auto& props = wb.properties();
props.setTitle("Report");
props.setSubject("Finance");
props.setCreator("XL++");
props.setDescription("Quarterly summary");
props.setKeywords("report, q4");
props.setCategory("Finance");
props.setLastModifiedBy("XL++");
```

---

## Styles Reference

### Font

```cpp
Font f;
f.setName("Calibri");
f.setSize(11.0);
f.setBold(true);
f.setItalic(false);
f.setUnderline(false);
f.setStrike(false);
f.color().setArgb("FFFF0000");
```

### Fill

```cpp
Fill fill;
fill.setPatternType("solid");  // or "none", "gray125", "darkGrid", ...
fill.foregroundColor().setArgb("FFFFFF00");
fill.backgroundColor().setArgb("FF000000");
```

### Border

```cpp
Border b;
b.left().setStyle("thin");        // "thin", "medium", "thick", "double", "dotted", ...
b.left().color().setArgb("FF000000");
b.right().setStyle("double");
b.top().setStyle("hair");
b.bottom().setStyle("medium");
b.diagonal().setStyle("none");
```

### Alignment

```cpp
Alignment a;
a.setHorizontal("center");    // "left", "center", "right", "fill", "justify", ...
a.setVertical("center");      // "top", "center", "bottom", "justify"
a.setWrapText(true);
a.setShrinkToFit(false);
a.setTextRotation(45);
a.setIndent(2);
```

### Color

```cpp
Color c("FFFF0000");    // ARGB hex string
c.setArgb("FF00FF00");
c.argb();               // ARGB value
c.empty();              // true if argb_ is empty
```

---

## Enums Reference

### CompressionLevel
`Store`, `Fastest`, `Default`, `Best`

### CompressionStrategy
`Default`, `Filtered`, `HuffmanOnly`, `Rle`, `Fixed`

### SharedStringMode
`Disabled`, `Hash`, `BoundedLru`

### FilterOperator
`Equal`, `NotEqual`, `LessThan`, `LessThanOrEqual`, `GreaterThan`, `GreaterThanOrEqual`

### ConditionalOperator
`Equal`, `NotEqual`, `LessThan`, `LessThanOrEqual`, `GreaterThan`, `GreaterThanOrEqual`, `Between`, `NotBetween`

### ConditionalRuleType
`CellIs`, `Formula`

### DataValidationType
`None`, `Whole`, `Decimal`, `List`, `Date`, `Time`, `TextLength`, `Custom`

### DataValidationOperator
`Between`, `NotBetween`, `Equal`, `NotEqual`, `LessThan`, `LessThanOrEqual`, `GreaterThan`, `GreaterThanOrEqual`

### DataValidationErrorStyle
`Stop`, `Warning`, `Information`

### PageOrientation
`Default`, `Portrait`, `Landscape`

### PaperSize
`Default`, `Letter`, `Tabloid`, `Legal`, `A3`, `A4`, `A5`, `B4`, `B5`, ...

### CellError
`Null`, `DivisionByZero`, `Value`, `Reference`, `Name`, `Number`, `NotAvailable`, `GettingData`

### FormulaType
`Normal`, `Shared`, `Array`, `DataTable`

---

## Build

Open `XL++.sln` with Visual Studio 2022+ (Platform Toolset v145), select x64
Debug or Release, then build.

```bash
# The solution builds:
#   - XLPP.lib          (static library)
#   - XLPP.UnitTests    (test runner, 55 suites)
#   - XLPP.Sample       (demo application)
```

---

## Complete Example

```cpp
#include <XLPP/XLPP.h>

int main() {
    using namespace xlpp;

    Workbook wb;
    auto& sheet = wb.addWorksheet("Report");

    // Header row
    sheet.append({"Product", "Revenue", "Margin"});

    auto& header = sheet.cell("A1");
    header.font().setBold(true);
    header.font().setSize(14);
    header.fill().setPatternType("solid");
    header.fill().foregroundColor().setArgb("FF4472C4");
    header.font().color().setArgb("FFFFFFFF");

    // Data
    sheet.append({"Widget-A", 1500.0, 0.35});
    sheet.append({"Widget-B", 2300.0, 0.42});
    sheet.append({"Widget-C", 800.0, 0.28});

    // Column widths
    sheet.columnDimension("A").width = 15.0;
    sheet.columnDimension("B").width = 12.0;
    sheet.columnDimension("C").width = 10.0;

    // Number format for revenue
    for (std::size_t r = 2; r <= 4; ++r)
        sheet.cell(r, 2).setNumberFormat("#,##0.00");

    // Conditional formatting: highlight low margin
    auto rule = ConditionalRule::cellIs(ConditionalOperator::LessThan, "0.30");
    rule.differentialStyle().font().color().setArgb("FF9C0006");
    rule.differentialStyle().fill().setPatternType("solid");
    rule.differentialStyle().fill().foregroundColor().setArgb("FFFFC7CE");
    sheet.conditionalFormatting().addRule("C2:C4", std::move(rule));

    wb.properties().setTitle("Sales Report");
    wb.properties().setCreator("XL++");

    wb.save("report.xlsx");
    return 0;
}
```
