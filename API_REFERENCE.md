# XL++ API Reference

XL++ is a dependency-light C++20 `.xlsx` read/write library inspired by openpyxl.
It uses the C++ standard library, zlib, and a platform crypto backend (Windows CNG/BCrypt or OpenSSL Crypto on non-Windows builds).

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

### Workbook templates (P1S)

XL++ now models template identity independently from VBA identity, allowing `.xltx` and `.xltm` packages to round-trip with the correct workbook content type.

```cpp
xlpp::Workbook wb;
wb.setTemplate(true);
wb.addWorksheet("TemplateData");
wb.save("model.xltx");

xlpp::Workbook loaded;
loaded.load("model.xltx");
if (loaded.isTemplate()) {
    // Template identity came from the package content type.
}
```

For a macro-enabled template, set the template flag and attach/generate a VBA project before saving to `.xltm`.

### Chartsheet printer settings and auxiliary ownership (P1W)

Chartsheets can own the opaque binary printer-settings payload referenced by top-level `pageSetup`. XL++ preserves the bytes as a package object; it does not decode operating-system printer structures.

```cpp
xlpp::Workbook wb;
wb.load("template.xltx");
auto& cs = wb.chartsheet("Dashboard");

cs.setPrinterSettingsData(std::string("\\x01\\x00\\x02", 3));
wb.save("template-with-printer-settings.xltx");

if (cs.hasPrinterSettings()) {
    const std::string& raw = *cs.printerSettingsData();
}

cs.clearPrinterSettings();
```

Imported header/footer-picture ownership (`legacyDrawingHF` plus its VML relationship/part) is preservation-safe across Chartsheet metadata edits and chart regeneration. Regeneration replaces the chart drawing closure but keeps unrelated Chartsheet auxiliary relationships unless the caller explicitly replaces/clears them.

The C ABI exposes equivalent length-based binary APIs: `xlpp_chartsheet_set_printer_settings`, `xlpp_chartsheet_printer_settings_size`, `xlpp_chartsheet_copy_printer_settings`, and `xlpp_chartsheet_clear_printer_settings`.

### Password-to-open encryption

Password-to-open encryption is package encryption, not worksheet/workbook protection. P1I keeps the configurable P1H Agile/Standard profiles and adds an in-memory inner-ZIP boundary, certificate key-encryptor inspection/tolerance for mixed Agile descriptors, and stricter load policies for untrusted encrypted files.

```cpp
xlpp::SaveOptions saveOptions;
saveOptions.encryption.enabled = true;
saveOptions.encryption.password = "P@ssw0rd\xE2\x9C\x93"; // UTF-8
saveOptions.encryption.mode = xlpp::PackageEncryptionMode::Agile;
saveOptions.encryption.keyBits = 256;                        // 128 / 192 / 256
saveOptions.encryption.hashAlgorithm = xlpp::PackageEncryptionHash::Sha512;
saveOptions.encryption.spinCount = 100000;
wb.save("protected.xlsx", saveOptions);

// Standard CryptoAPI/AES compatibility writer.
saveOptions.encryption.mode = xlpp::PackageEncryptionMode::Standard;
saveOptions.encryption.keyBits = 128; // 128 / 192 / 256
wb.save("protected-standard.xlsx", saveOptions);

const auto profile = xlpp::Workbook::inspectPasswordEncryptionFile("protected.xlsx");
// profile.format / keyBits / hashAlgorithm / spinCount / hasDataIntegrity

xlpp::LoadOptions loadOptions;
loadOptions.passwordToOpen = "P@ssw0rd\xE2\x9C\x93";
loadOptions.maxEncryptionSpinCount = 1'000'000;  // 0 accepts spec maximum
loadOptions.maxDecryptedPackageBytes = 512u * 1024u * 1024u;
loadOptions.maxEncryptionInfoBytes = 1024u * 1024u;
loadOptions.allowStandardEncryption = true;
loadOptions.requireAgileDataIntegrity = true;
xlpp::Workbook protectedBook;
protectedBook.load("protected.xlsx", loadOptions);
```

`save(std::ostream&, SaveOptions)` and `load(std::istream&, LoadOptions)` support the same password options. Wrong passwords are rejected before OOXML parsing. Agile `DataIntegrity` HMAC is verified before decrypted package bytes are accepted.

**P1H Agile read/write profiles:** AES-128/192/256-CBC with SHA-1, SHA-256, SHA-384 or SHA-512. Default: AES-256/SHA-512 with `spinCount=100000`.

**P1H Standard read/write profiles:** CryptoAPI AES-128/192/256 + SHA-1. Standard derivation is fixed to 50,000 iterations; `PackageEncryptionOptions::spinCount` and `hashAlgorithm` are intentionally ignored for Standard writes.

Profile inspection does not require the password and returns `PackageEncryptionInfo`, including format/version, cipher/hash, key bits, block/hash size, spin count, data-integrity presence and current read/write support. P1I also reports total/password key-encryptor counts and decoded metadata for certificate key-encryptors. Certificate metadata inspection does **not** validate a certificate chain and does not implement private-key decryption.

P1I encrypted file load/save serializes/parses the inner OOXML ZIP directly in memory; it no longer writes a plaintext temporary inner `.xlsx` while wrapping/unwrapping the encrypted CFB package. Stream APIs still use their existing outer temporary-file bridge, but encrypted streams no longer create a plaintext inner workbook on disk.

C API adds `xlpp_workbook_save_password_ex` and `xlpp_workbook_encryption_profile` alongside the P1G password APIs. C# exposes the configurable `SaveEncrypted(...)` overload plus `InspectPasswordEncryptionFile`. Python exposes `PackageEncryptionMode`, `PackageEncryptionHash`, `PackageEncryptionFormat`, `PackageEncryptionInfo`, and the new resource guards.

Agile descriptors may contain certificate key-encryptors alongside the required single password key-encryptor; XL++ P1I selects the password descriptor by URI and exposes certificate metadata during inspection. Certificate-only decryption/writing, RC4 and Extensible Encryption remain unsupported and fail explicitly rather than being guessed.

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
opts.maxEntryBytes = 10 * 1024 * 1024;  // limit one uncompressed entry
opts.maxTotalBytes = 100 * 1024 * 1024; // limit total uncompressed ZIP payload
opts.maxFileBytes = 128 * 1024 * 1024;  // limit source file/stream bytes
opts.maxWorksheets = 256;               // limit materialized sheets (0 = unlimited)
opts.maxCells = 5'000'000;              // total unique materialized cells
opts.maxSharedStrings = 1'000'000;      // SST cardinality guard
opts.maxDefinedNames = 100'000;         // defined-name model guard
opts.cancel = []{ return someFlag; };   // cancellation callback
opts.progress = [](auto done, auto total) { ... }; // progress callback
wb.load("large.xlsx", opts);
```

P1R applies a strong exception guarantee to `Workbook::load`: package/parser/resource failures leave the pre-existing `Workbook` object unchanged. The same guarantee applies to `load(std::istream&, LoadOptions)` after its bounded RAII temporary-file bridge. The model-level guards complement ZIP byte limits and are useful against highly compressed XML that would otherwise expand into a large object graph.

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

## VBA source projects (P1A)

XL++ can preserve arbitrary external `vbaProject.bin` parts unchanged. For projects generated by XL++ it also exposes a source-oriented C++ authoring model for standard, class and workbook/worksheet document modules.

```cpp
xlpp::Workbook wb;
wb.addWorksheet("Data");

wb.setVbaModuleText("Utilities",
    "Public Function Twice(x As Double) As Double\n"
    "Twice = x * 2\nEnd Function");
wb.setVbaClassModuleText("WorkerClass", "Option Explicit\nPublic Name As String");
wb.setVbaDocumentModuleText("ThisWorkbook",
    "Private Sub Workbook_Open()\nWorksheets(1).Range(\"A1\").Value = \"Ready\"\nEnd Sub");

xlpp::VbaModule module;
module.name = "WorkerClass";
module.type = xlpp::VbaModuleType::Class;
module.source = "Option Explicit\nPublic Name As String";
module.docString = "Worker metadata";
module.readOnly = true;
module.isPrivate = true;
wb.setVbaModule(std::move(module));
```

Project metadata and registered type-library references are first class:

```cpp
xlpp::VbaProjectInfo info;
info.name = "AnalyticsProject";
info.description = "XL++ generated VBA source project";
info.helpFile = "project-help.chm";
info.helpContextId = 42;
info.constants = "Build = 1";
info.projectId = "{12345678-1234-4ABC-9DEF-1234567890AB}";
info.references.push_back({
    "Scripting",
    "*\\G{420B2830-E718-11CF-893D-00A0C9054228}#1.0#0#C:\\Windows\\System32\\scrrun.dll#Microsoft Scripting Runtime"
});
wb.setVbaProjectInfo(info);

const auto modules = wb.vbaModules();
const auto project = wb.vbaProjectInfo();
auto source = wb.vbaModuleText("ThisWorkbook");
wb.removeVbaModule("WorkerClass"); // host document modules cannot be removed here
```

XL++ rebuilds its own source-generated VBA project on save so `Sheet1`, `Sheet2`, ... document modules remain synchronized after worksheet insertion/removal while existing document/class/standard source and project metadata are retained.

**Current VBA boundary:** this authoring layer does not claim to compile VBA, author UserForms/FRX/designer streams, edit digital signatures, unlock password-protected projects, or reproduce arbitrary compiled p-code. Untouched external VBA projects continue to use opaque package preservation.

---

## Pivot tables (P1A)

New PivotTables can still be generated directly:

```cpp
xlpp::PivotTable pivot("SalesPivot");
pivot.setLocation("F2");
pivot.cache().setSourceData("'Data'!$A$1:$D$100");
pivot.addRowField("Region");
pivot.addColumnField("Quarter");
pivot.addPageField("Year");
auto& values = pivot.addDataField("Amount", "sum");
values.setDisplayName("Sales %");
values.setShowDataAs("percentOfTotal");
sheet.addPivotTable(std::move(pivot));
```

P1A also loads common existing `pivotTableDefinition`, `pivotCacheDefinition` and `pivotCacheRecords` parts into the native model. Discovery works both with normal worksheet `<pivotTableParts>` owner nodes and producer variants that expose a PivotTable only through worksheet relationships.

```cpp
xlpp::Workbook wb;
wb.load("existing-pivot.xlsx");

// Use a const view for inspection; untouched imported Pivot OOXML is preserved.
const auto& readOnly = static_cast<const xlpp::Workbook&>(wb);
const auto* ws = readOnly.worksheet("Data");
const auto& p = ws->pivotTables().front();
std::cout << p.name() << ' ' << p.location() << '\n';
for (const auto& field : p.cache().fields()) std::cout << field << '\n';

// Mutable access opts into regeneration through XL++'s modeled Pivot layer.
auto& edited = wb.worksheet("Data")->pivotTables().front();
edited.setStyleName("PivotStyleMedium9");
edited.setRowGrandTotals(false);
edited.cache().setRefreshOnLoad(false);
edited.rowFields().front().hideItem(1);
edited.dataFields().front().setDisplayName("Sales %");
edited.dataFields().front().setShowDataAs("percentOfTotal");
wb.save("edited-pivot.xlsx");
```

Modeled Pivot state includes cache source/options/fields/records, row/column/page/data-field indices, hidden items, sort/default-subtotal flags, page-field settings, common data aggregations, `showDataAs`, base field/item, number-format ID, style/layout/grand-total/header/stripe settings.

**Current Pivot boundary:** untouched advanced Pivot OOXML is preservation-safe, but choosing mutable imported-Pivot regeneration uses the modeled subset. OLAP caches, grouping, calculated members/items, slicers/timelines, pivot charts and complete shared-cache editing are not yet fully modeled. For conservative safety, a regenerated imported Pivot may leave an otherwise valid preserved workbook cache in place when that cache could be shared.

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

### Structural edits and transactions (P1K)

For a worksheet that belongs to a `Workbook`, prefer the workbook-level APIs. They transform references across the complete in-memory model instead of moving only the local cell grid.

```cpp
auto report = wb.insertRows("Data", 2, 3);
wb.deleteRows("Data", 10, 2);
wb.insertColumns("Data", 2, 1);
wb.deleteColumns("Data", 3, 2);
```

P1K retains the P1J workbook-wide A1/range/whole-row/whole-column repair and makes it transactional by default. Deleted references become `#REF!`; physical edits that would move cells outside Excel's `1,048,576 x 16,384 (XFD)` grid are rejected. `StructuralEditReport` exposes rewrite/invalidation counters, post-edit validation counts and diagnostics.

Transactional options:

```cpp
xlpp::StructuralEditOptions options;
options.rollbackOnFailure = true;   // default
options.validateResult = true;      // default
options.cancel = [&]() { return shouldCancel(); };

try {
    auto report = wb.insertRows("Data", 1000, 500, options);
} catch (const xlpp::StructuralEditCancelled&) {
    // Model has been restored when rollbackOnFailure=true.
}
```

Rollback restores worksheet contents in place, plus workbook defined names and calculation/cache state, so existing `Worksheet&` references keep their identity. Post-edit validation compares against the pre-edit error set and rejects only newly introduced semantic errors. Set `rollbackOnFailure=false` only when a caller explicitly accepts partial mutation on failure/cancellation.

3-D references such as `Start:End!A1` are recognized. Workbook-safe sheet rename/remove rewrites or invalidates a matching 3-D endpoint. Row/column structural edits deliberately preserve 3-D coordinate references and increment `StructuralEditReport::referencesSkippedUnsupported`, because guessing how a structural edit should change a 3-D area can silently corrupt semantics.

Direct worksheet-local calls remain available for standalone worksheet models:

```cpp
sheet.insertRows(2, 3);
sheet.deleteColumns(3, 2);
```

Because a `Worksheet` has no parent-workbook pointer, direct local calls cannot rewrite references owned by sibling worksheets or workbook defined names.

Safe workbook-level sheet lifecycle operations are also available:

```cpp
wb.renameWorksheet("Data", "Input Data"); // rewrites workbook-wide qualifiers
wb.removeWorksheet("Old Data");           // invalidates surviving refs with #REF!
```

Worksheet lookup and safe rename/remove are case-insensitive like Excel. Removing the final worksheet is rejected.

### In-memory model integrity validation (P1J)

```cpp
auto validation = wb.validateModelIntegrity();
if (!validation.ok()) {
    for (const auto& issue : validation.issues) {
        // issue.code is stable for automation, e.g.
        // "pivot.record_width_mismatch"
    }
}
```

The validator catches semantic problems that OPC relationship validation cannot see, including invalid grid geometry, table/Pivot schema mismatches, bad field indices, duplicate workbook object names, missing Pivot/chart source sheets, invalid defined-name scopes and broken object references. Warnings such as an intentional `#REF!` are kept separate from errors.

Serialization can opt into a pre-save semantic gate:

```cpp
xlpp::SaveOptions save;
save.validateModelBeforeSave = true;
save.rejectModelWarningsBeforeSave = true; // optional stricter semantic policy
save.validatePackageBeforeWrite = true;    // validate assembled OPC before write/encryption
wb.save("validated.xlsx", save);
```

All three options default to `false` for preservation compatibility. `validateModelBeforeSave` blocks semantic errors; `rejectModelWarningsBeforeSave` additionally promotes warnings such as `#REF!` to fatal when model validation is enabled. `validatePackageBeforeWrite` validates relationships, duplicate IDs, dangling/orphan parts, content types and owner references on the fully assembled inner OOXML package before bytes are written or encrypted.

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

Existing embedded images loaded from an XLSX expose package-origin metadata through
`stableId()` and `anchorInfo()`. Use the selective mutation APIs below when editing
those imported images so XL++ patches only the target DrawingML object instead of
regenerating the complete drawing:

```cpp
const auto& imported = static_cast<const Worksheet&>(sheet).images().front();
const std::string id = imported.stableId();

sheet.moveImage(id, "H6");                 // one-cell / two-cell anchors
sheet.resizeImage(id, 160.0, 90.0);         // pixels
sheet.replaceImage(id, "updated-logo.png"); // keeps anchor and sibling objects
sheet.removeImage(id);                       // removes relationship/media when unused

// Absolute anchors are positioned directly in EMU.
sheet.moveImageAbsolute(id, 250000, 500000);
```

`images()` mutable access still opts into full drawing regeneration. Prefer the
selective APIs for imported images whenever sibling charts, shapes or unknown
DrawingML must remain untouched. Selective stable-ID edits are routed to the
image's original `sourceDrawingPart()`, so producer workbooks containing more
than one preserved drawing relationship can be patched without rebuilding the
other drawing parts.

### Imported charts

Charts discovered in existing DrawingML expose package-origin metadata through
`stableId()`, `anchorInfo()`, `sourceDrawingPart()`, `sourceChartPart()` and
`sourceRelationshipId()`. Read access is namespace-tolerant for common Excel,
OpenPyXL and LibreOffice chart XML, including title, axis titles, series formulas,
legend position and anchor geometry. Imported charts also expose native OOXML axis
structure through `axes()` / `plots()`, including `axId`, `crossAx`, primary/secondary
classification and combined-chart plot membership. Plot-level and series-level data labels,
per-point `dLbl` overrides, series line/fill/marker formatting, series trendlines and X/Y
error bars (including custom plus/minus formula ranges) are also exposed for imported charts.

Use the stable-ID APIs to edit a controlled subset without regenerating the whole
chart/drawing:

```cpp
const auto& importedChart = static_cast<const Worksheet&>(sheet).charts().front();
const std::string chartId = importedChart.stableId();

sheet.setChartTitle(chartId, "Updated title");
sheet.setChartXAxisTitle(chartId, "Category");
sheet.setChartYAxisTitle(chartId, "Amount");

// Native-axis targeting is useful for combined/secondary-axis charts.
if (const auto* axis = importedChart.axisById(200); axis && axis->secondary)
    sheet.setChartAxisTitle(chartId, axis->id, "Secondary amount");

sheet.setChartLegend(chartId, true, "b"); // l, r, t, b, tr
sheet.setChartSeriesTitle(chartId, 0, "Revenue");
sheet.setChartSeriesReferences(chartId, 0,
                               "'Data'!$A$2:$A$20",
                               "'Data'!$B$2:$B$20");

// Preservation-aware labels / trendline / error-bar edits.
auto labels = importedChart.plots().front().dataLabels;
labels.showValue = true;
labels.showCategoryName = true;
sheet.setChartPlotDataLabels(chartId, 0, labels);

auto trend = importedChart.series().front().trendlines().front();
trend.type = ChartSeries::TrendlineType::Polynomial;
trend.order = 3;
sheet.setChartSeriesTrendline(chartId, 0, 0, trend);

ChartSeries::ErrorBars bars;
bars.direction = ChartSeries::ErrorBarDirection::Y;
bars.valueType = ChartSeries::ErrorValueType::FixedValue;
bars.value = 2.0;
sheet.setChartSeriesErrorBars(chartId, 0, bars);

// Per-point labels use the native zero-based c:idx value.
ChartDataLabelPoint pointLabel;
pointLabel.index = 3;
pointLabel.showValue = true;
pointLabel.position = "t";
sheet.setChartSeriesDataLabelPoint(chartId, 0, pointLabel);

// Custom error bars retain independent plus/minus formula ranges.
ChartSeries::ErrorBars customBars;
customBars.direction = ChartSeries::ErrorBarDirection::Y;
customBars.valueType = ChartSeries::ErrorValueType::Custom;
customBars.plusReference = "'Data'!$C$2:$C$20";
customBars.minusReference = "'Data'!$D$2:$D$20";
customBars.lineFormat.present = true;
customBars.lineFormat.color = {ChartColor::Kind::SRgb, "404040"};
customBars.lineFormat.widthPoints = 1.25;
sheet.setChartSeriesErrorBars(chartId, 0, customBars);

ChartLineFormat seriesLine;
seriesLine.present = true;
seriesLine.color = {ChartColor::Kind::SRgb, "1F4E78"};
seriesLine.widthPoints = 2.0;
seriesLine.dash = "dash";
sheet.setChartSeriesLineFormat(chartId, 0, seriesLine);

auto marker = importedChart.series().front().markerFormat();
marker.symbol = "diamond";
marker.size = 9;
marker.fill.present = true;
marker.fill.color = {ChartColor::Kind::SRgb, "70AD47"};
sheet.setChartSeriesMarkerFormat(chartId, 0, marker);

sheet.moveChart(chartId, "H6");
sheet.resizeChart(chartId, 640.0, 360.0);

// Absolute anchors use EMU coordinates.
sheet.moveChartAbsolute(chartId, 250000, 500000);

// Imported charts can also be removed selectively. If a loaded worksheet
// already owns a preserved drawing, addChart() appends a new chart without
// regenerating untouched sibling DrawingML.
sheet.removeChart(chartId);
sheet.addChart(replacementChart);
```

Selective title editing patches only the requested title subtree. Axis-title,
legend and series-title edits patch their own ChartML regions; series-reference
edits patch the selected `<ser>` formulas and discard only the corresponding
stale category/value caches. Scatter and bubble charts map X/Y titles to their native first/second plot axis IDs instead
of assuming a category axis. Combined charts retain each plot's type, grouping, series
span and native axis IDs; `setChartXAxisTitle()` / `setChartYAxisTitle()` target the
primary plot pair by `axId`, while `setChartAxisTitle()` can address a secondary axis
directly. P0K adds namespace-tolerant read models and selective mutation for plot/series
data labels, series trendlines and error bars. P0L extends that path with direct-child-safe
aggregate label editing, per-point `dLbl` overrides, custom plus/minus error-bar formula ranges,
and selective series line/fill/marker plus trendline/error-bar line formatting. P0M adds a
preservation-aware `dPt` model, rich chart-title / per-point label text runs, gradient and pattern
fills, scheme/RGB color transforms, line cap/compound/join metadata, and custom dash sequences.
Selective mutations patch only the targeted `title`, `dLbl`, `dPt`, or `spPr` subtree. Newly
appended trendlines and error bars also retain their requested line formatting. Unsupported
ChartML formatting/extensions and sibling DrawingML remain untouched unless their specific
selective API is invoked.
P0M selective APIs include `setChartTitleRichText()`,
`setChartSeriesDataLabelPointRichText()`, `setChartSeriesDataPointFormat()`, and
`removeChartSeriesDataPointFormat()`. `ChartSeries::dataPoints()` / `dataPoint(idx)` expose
imported `dPt` formatting without marking the drawing dirty. `ChartRichText` exposes styled
DrawingML runs, while `ChartColor::transforms`, `ChartFillFormat::gradientStops` /
`pattern`, and `ChartLineFormat::customDash` / `cap` / `compound` / `join` retain common
advanced formatting metadata.


P0N adds preservation-aware layout/axis/legend inspection and mutation. `Chart::plotAreaLayout()` and `Chart::legendFormat()` expose imported manual-layout geometry. Each `Chart::Axis` now retains `numFmt`, source-linked state, tick marks, tick-label position, major/minor units, crossing behavior, rich title text, axis line formatting and major/minor gridline line formatting. Selective APIs include:

```cpp
sheet.setChartAxisTitleRichText(chartId, axisId, richText);
sheet.setChartAxisNumberFormat(chartId, axisId, "0.000", false);
sheet.setChartAxisTicks(chartId, axisId, "out", "none", "nextTo");
sheet.setChartAxisUnits(chartId, axisId, 5.0, 1.0);
sheet.setChartAxisCrossing(chartId, axisId, "autoZero", "between");
sheet.setChartAxisLineFormat(chartId, axisId, axisLine);
sheet.setChartAxisGridlineFormat(chartId, axisId, true, majorGridline);
sheet.setChartPlotAreaLayout(chartId, plotLayout);
sheet.setChartLegendLayout(chartId, legendLayout);
sheet.setChartLegendOverlay(chartId, false);
sheet.setChartLegendLineFormat(chartId, legendLine);
sheet.setChartLegendFillFormat(chartId, legendFill);
```

These operations patch only the matching axis, `layout`, legend `spPr`, or overlay node. Unsupported sibling ChartML remains preserved.

P0O extends this model with axis scaling, crosses-at values, display units, gridline lifecycle, and chart/plot-area shape properties. Imported `Chart::Axis` instances now expose `scaling`, `hasCrossesAt` / `crossesAt`, `displayUnits`, and explicit major/minor-gridline presence. `Chart` also exposes chart-area and plot-area fill/line formatting. Example selective edits:

```cpp
ChartAxisScaling scaling;
scaling.hasMinimum = true; scaling.minimum = 0.5;
scaling.hasMaximum = true; scaling.maximum = 500.0;
scaling.hasLogBase = true; scaling.logBase = 10.0;
scaling.reverseOrder = false;
sheet.setChartAxisScaling(chartId, axisId, scaling);

sheet.setChartAxisCrossesAt(chartId, axisId, 5.5);
sheet.clearChartAxisCrossesAt(chartId, axisId);

ChartDisplayUnits units;
units.present = true;
units.hasCustomUnit = true;
units.customUnit = 1000000.0;
units.showLabel = true;
sheet.setChartAxisDisplayUnits(chartId, valueAxisId, units);
sheet.clearChartAxisDisplayUnits(chartId, valueAxisId);

sheet.removeChartAxisGridlines(chartId, axisId, false); // remove minor gridlines
sheet.setChartAreaLineFormat(chartId, chartAreaLine);
sheet.setChartAreaFillFormat(chartId, chartAreaFill);
sheet.setChartPlotAreaLineFormat(chartId, plotAreaLine);
sheet.setChartPlotAreaFillFormat(chartId, plotAreaFill);
```

`setChartAxisCrossing()` removes an existing `crossesAt` when a categorical crossing mode is selected; `setChartAxisCrossesAt()` removes `crosses`. Display-unit mutation is currently restricted to value axes and accepts the OOXML built-in unit names or one positive custom unit. Scaling validation rejects inverted min/max ranges and invalid logarithmic bases.

P0P adds inspection and selective mutation for common auxiliary ChartML objects. `Chart::dataTable()` exposes the plot-area data table, each `Chart::Plot` exposes `hasDropLines`, `dropLinesFormat`, `hasHighLowLines`, `highLowLinesFormat`, and `upDownBars`, while `ChartDataLabels` exposes explicit `leaderLines` presence and line formatting. Example:

```cpp
ChartDataTable table;
table.showHorizontalBorder = true;
table.showOutline = true;
sheet.setChartDataTable(chartId, table);

ChartLineFormat drop;
drop.present = true;
drop.color = {ChartColor::Kind::SRgb, "4472C4"};
sheet.setChartPlotDropLines(chartId, 0, drop);
sheet.removeChartPlotHighLowLines(chartId, 0);

ChartUpDownBars bars;
bars.gapWidth = 100;
sheet.setChartPlotUpDownBars(chartId, 0, bars);

sheet.setChartPlotLeaderLineFormat(chartId, 0, drop);
sheet.setChartSeriesLeaderLineFormat(chartId, 0, drop);
```

Lifecycle APIs `removeChartDataTable()`, `removeChartPlotDropLines()`, `removeChartPlotHighLowLines()`, `removeChartPlotUpDownBars()`, `removeChartPlotLeaderLines()`, and `removeChartSeriesLeaderLines()` remove only the selected child nodes.

`removeChart()` cleans the anchor/relationship and exclusively-owned dependency
parts. `addChart()` uses an additive preserved-drawing path on loaded worksheets.
Mutable `chart()` / `charts()` access still opts into full regeneration.

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
// Optional resource guards for untrusted/very large packages.
xlpp::StreamingReadOptions readOptions;
readOptions.maxEntries = 100000;
readOptions.maxEntryBytes = 512u * 1024u * 1024u;
readOptions.maxTotalBytes = 2ull * 1024u * 1024u * 1024u;
readOptions.maxFileBytes = 4ull * 1024u * 1024u * 1024u;
readOptions.maxXmlElementBytes = 64u * 1024u * 1024u; // default
xlpp::StreamingWorkbookReader guardedReader("large.xlsx", readOptions);

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

## P0Q — Stock Charts and Data-table Text

P0Q adds first-class `Chart::Type::Stock` support for imported and newly generated high-low-close / open-high-low-close stock charts. Imported `<stockChart>` plots expose the same `Chart::Plot` auxiliary model used by line charts, including high-low lines and up/down bars. New stock charts require exactly three or four series and serialize numeric category references.

For generation, use `Chart::primaryPlot()` to configure auxiliary objects before adding the chart to a worksheet:

```cpp
Chart stock(Chart::Type::Stock);
// add 3 or 4 ChartSeries objects...
auto& plot = stock.primaryPlot();
plot.hasHighLowLines = true;
plot.highLowLinesFormat.present = true;
plot.upDownBars.present = true;
plot.upDownBars.gapWidth = 100;
```

`ChartDataTable` also exposes `textStyle` (`ChartTextStyle`) for `dTable/txPr` default text properties: bold, italic, font size, typeface and color. `setChartDataTable()` selectively patches imported charts; generated charts serialize the same text style without rebuilding unrelated DrawingML.

---

## P0R — 3D and Surface Chart Preservation Foundation

P0R extends `Chart::Type` with `Bar3D`, `Line3D`, `Area3D`, `Pie3D`, `Surface`, and `Surface3D`. Imported charts expose native plot/axis structure without flattening unsupported ChartML. `Chart::view3D()` exposes rotation, height/depth percentages, right-angle-axis mode and perspective. `floorFormat()`, `sideWallFormat()` and `backWallFormat()` expose wall thickness plus fill/line formatting.

Selective imported-chart mutations use stable chart IDs and patch only the requested chart-level owner:

```cpp
auto view = chart.view3D();
view.present = true;
view.hasRotationX = true;
view.rotationX = 30;
sheet.setChartView3D(chart.stableId(), view);

auto floor = chart.floorFormat();
floor.present = true;
floor.hasThickness = true;
floor.thickness = 24;
sheet.setChartFloorFormat(chart.stableId(), floor);
```

New `Bar3D` and `Surface3D` charts serialize three-axis structures (`catAx`, `valAx`, `serAx`). `Chart::primaryPlot()` also exposes `gapDepth`, `shape`, and `wireframe` metadata for supported 3D/surface plots. Direct OpenPyXL validation covers all six imported chart types plus generated Bar3D/Surface3D. LibreOffice preserves Bar3D/Line3D/Area3D/Pie3D but converts Surface/Surface3D to Bar3D when Calc re-saves the workbook; that conversion is a host normalization rather than an XL++ write-time loss.

---

## P0S — Projected Pie, Doughnut and Radar Expansion

P0S adds first-class `Chart::Type::PieOfPie` and `Chart::Type::BarOfPie` while preserving existing enum values for previously published chart types. Both map to OOXML `ofPieChart`; the plot model distinguishes them through the native `ofPieType` value.

`Chart::Plot` now exposes type-specific metadata for projected pie, doughnut/pie and radar charts:

```cpp
plot.projectedPie;        // gapWidth, splitType/splitPos, custSplit, secondPlotSize, serLines
plot.hasFirstSliceAngle;
plot.firstSliceAngle;
plot.hasHoleSize;
plot.holeSize;
plot.radarStyle;          // standard / marker / filled
```

Selective edits on imported charts use the stable chart ID and plot index:

```cpp
sheet.setChartPlotProjectedPieOptions(chartId, 0, options);
sheet.setChartPlotFirstSliceAngle(chartId, 0, 120);
sheet.setChartPlotDoughnutHoleSize(chartId, 0, 70);
sheet.setChartPlotRadarStyle(chartId, 0, "marker");
```

New Pie-of-Pie, Bar-of-Pie, Doughnut and Radar charts serialize the same type-specific metadata. Projected-pie validation accepts split types `auto`, `cust`, `percent`, `pos`, and `val`; first-slice angles are limited to 0–360 degrees and doughnut hole size to 10–90 percent. OpenPyXL 3.1.5 validates the direct XL++ output. LibreOffice Calc can normalize projected-pie split parameters and doughnut hole size when it re-saves the workbook, so those post-Calc values are recorded as host normalization rather than XL++ write-time loss.

---


## P0T — Chart Style, Theme and Series Cache Foundation

P0T adds first-class chart style/theme/cache metadata without changing the preservation-first ChartML pipeline. `Chart` now exposes `style()`, `themePalette()` and `styleResources()`. `ChartThemePalette` contains the workbook theme color scheme and can resolve the base RGB value of a `ChartColor::Kind::Scheme` color while preserving its transform list. `ChartStyleResources` records chart-style and chart-color-style relationship targets so untouched style resources remain connected and byte-preserved.

Each `ChartSeries` exposes cached data for its title, categories and values:

```cpp
series.titleCache();       // strCache
series.categoriesCache();  // strCache or numCache
series.valuesCache();      // numCache
```

`ChartSeriesCache` stores `present`, `numeric`, `formatCode`, `pointCount` and indexed cache points. Selective imported-chart mutations are available by stable chart ID:

```cpp
sheet.setChartStyle(chartId, "15");
sheet.setChartSeriesTitleCache(chartId, 0, titleCache);
sheet.setChartSeriesCategoryCache(chartId, 0, categoryCache);
sheet.setChartSeriesValueCache(chartId, 0, valueCache);
sheet.clearChartSeriesCaches(chartId, 0);
```

Cache setters validate explicit `pointCount` against the indexed points and reject inconsistent caches. Generated `ChartSeries` objects can use `setTitleCache()`, `setCategoriesCache()` and `setValuesCache()` before `addChart()`; the writer emits the corresponding `strCache` / `numCache` next to the formula reference.

Direct OpenPyXL validation confirms cache/style metadata produced by XL++. LibreOffice Calc re-computes chart caches from worksheet formula references and removes chart-style/color-style relationship resources when it becomes the writer; this is documented as host normalization rather than an XL++ round-trip loss.

---

## P0U — Cache Synchronization & Theme Transform Engine

P0U adds workbook-level synchronization of chart caches from worksheet A1 references. The synchronizer supports local or quoted cross-sheet single-cell and one-dimensional ranges, including sheet names containing escaped apostrophes. External-workbook references, unions, structured references and two-dimensional ranges are intentionally skipped and reported instead of guessed.

```cpp
ChartCacheSyncReport report = workbook.synchronizeChartCaches();

// Typical report fields:
report.chartsVisited;
report.seriesVisited;
report.cachesUpdated;
report.cachesCleared;
report.referencesSkipped;
report.warnings;
```

Synchronization rebuilds title `strCache`, category `strCache`/`numCache`, and value `numCache` directly from the current worksheet cell values. Blank cells are represented as sparse cache indexes while `pointCount` remains equal to the referenced range length. Numeric caches inherit an existing cache format code when available, otherwise they use the first non-General source-cell number format.

`ChartSeriesCache` now exposes validation/introspection helpers:

```cpp
cache.valid();
cache.hasDuplicateIndexes();
cache.ordered();
cache.sparse();
cache.effectivePointCount();
```

The chart theme model now also exposes theme font/effect metadata and a sequential DrawingML color-transform resolver:

```cpp
const auto& theme = chart.themePalette();
theme.fontScheme.majorLatinTypeface;
theme.fontScheme.minorLatinTypeface;
theme.effectScheme.fillStyleCount;

ChartResolvedColor resolved = chart.resolveThemeColor(color);
std::string rgb = chart.resolveThemeFinalRgb(color);
```

Supported transform operations are `alpha`, `alphaMod`, `alphaOff`, `tint`, `shade`, `lumMod`, `lumOff`, `satMod`, and `satOff`. Direct SRGB and scheme colors are resolved while preserving the original `ChartColor` and transform sequence.

LibreOffice Calc may recompute chart caches and normalize cache formatting when it re-saves a workbook; direct XL++ output retains sparse cache indexes and format codes verified by OpenPyXL.

---

## P0V — Automatic Cache Dependency Tracking & Style Resolution

P0V adds dependency-aware incremental synchronization on top of P0U. The dependency model can be inspected without rebuilding cache data:

```cpp
auto dependencies = workbook.chartCacheDependencies();
for (const auto& dep : dependencies) {
    dep.ownerSheet;
    dep.sourceSheet;
    dep.chartStableId;
    dep.seriesIndex;
    dep.kind;       // Title / Category / Value
    dep.first;
    dep.last;
    dep.supported;
}
```

Non-const worksheet cell/range access records touched cell keys, while value/formula/number-format mutations also increment a per-cell mutation revision. This keeps retained `Cell&` references detectable even after a tracker reset. Incremental synchronization uses both signals to skip unrelated cache references:

```cpp
ChartCacheSyncOptions options;
options.clearTrackedChangesAfterSync = true;
ChartCacheSyncReport report = workbook.synchronizeChangedChartCaches(options);

report.dependenciesVisited;
report.dependenciesMatched;
report.dependenciesSkippedUnchanged;
report.formulaCachePointsReused;
```

Full `synchronizeChartCaches()` behavior is unchanged. `onlyChangedCells` can also be enabled directly in `ChartCacheSyncOptions`. `clearChartCacheChangeTracking()` resets only the cache dependency tracker and does not clear the normal worksheet dirty flag.

When a referenced source cell contains a formula but has no cached worksheet value, `preserveFormulaCachedValues` (default `true`) reuses the existing chart cache point at the same index. This prevents cache data from disappearing merely because XL++ does not calculate the formula itself.

Chart style-resource inspection is also deeper:

```cpp
const auto& resources = chart.styleResources();
resources.chartStyleId;
resources.colorStyleId;
resources.colorStyleMethod;
resources.colorStyleColors;

auto resolved = resources.resolveColorStyle(chart.themePalette());
auto majorFont = chart.themePalette().resolveTypeface("+mj-lt");
auto minorFont = chart.themePalette().resolveTypeface("+mn-lt");
```

Direct XL++ output preserves chart-style/color-style relationship resources. LibreOffice Calc may remove those extension resources when it becomes the writer; that remains documented host normalization.

---

## Build

Open `XL++.sln` with Visual Studio 2026 (Platform Toolset v145), select x64
Debug or Release, then build.

```bash
# The solution builds:
#   - XLPP.lib          (static library)
#   - XLPP.UnitTests    (test runner, 155 suites)
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

## P0W — Formula dependency propagation and chart color-style application

### Incremental formula precedent propagation

`Workbook::synchronizeChangedChartCaches()` now follows simple local/cross-sheet A1 precedents of formula cells inside chart source ranges. When a tracked precedent changes but the formula cell itself has not been recalculated, XL++ preserves the existing chart cache and can request host recalculation rather than writing a stale cached worksheet value as if it were current.

`ChartCacheSyncOptions` adds `propagateFormulaDependencies`, `requestHostRecalculationForFormulaDependencies`, and `maxFormulaDependencyDepth`. `ChartCacheSyncReport` adds `formulaDependenciesVisited`, `formulaDependenciesMatched`, `staleFormulaCachesPreserved`, and `hostRecalculationRequested`.

Supported precedent syntax is intentionally conservative: simple A1 cell/range references, including quoted local worksheet names. This is dependency tracking, not formula evaluation.

### Chart color-style application

`Workbook::applyChartColorStyle(worksheetName, chartStableId, applyFill, applyLine, applyMarker)` resolves imported chart-color-style entries through the workbook theme and materializes the resulting SRGB colors into the selected imported chart series using the existing selective ChartML formatting path. It returns `ChartStyleApplyReport` with visited/styled series counts and diagnostics.

This is a foundation for style application; it does not yet reproduce the full Excel chart-style matrix/effect engine.

## P0X/P0Y — Expanded dependency grammar and theme style matrix

`Workbook::synchronizeChartCaches()` and `Workbook::synchronizeChangedChartCaches()` accept more reference forms without turning XL++ into a spreadsheet calculation engine. Formula-dependency traversal can resolve rectangular A1 ranges, whole rows/columns, workbook/local defined names, common structured table references, and statically bounded reference-form `OFFSET`/`INDEX` defined names. One-dimensional structured/defined-name sources may also be materialized directly into chart caches. Calculation-dependent dynamic references are reported and preserved for host recalculation rather than guessed.

`ChartCacheSyncReport` additionally reports `structuredReferencesVisited`, `structuredReferencesResolved`, `structuredReferencesSkipped`, `dynamicDefinedNamesVisited`, `dynamicDefinedNamesResolved`, and `dynamicDefinedNamesSkipped`, alongside the P0W formula-dependency counters and diagnostics.

`ChartThemeEffectScheme` exposes ordered materialized `fillStyles`, `lineStyles`, `effectStyles`, and `backgroundFillStyles`. The materialized vectors preserve theme matrix order; they are inspection/application primitives rather than a spreadsheet-theme evaluator.

`Workbook::applyChartThemeStyleMatrix(worksheetName, chartStableId, fillStyleIndex, lineStyleIndex, applyMarker)` applies zero-based theme `fmtScheme` fill/line entries to each series of an imported chart. `phClr` placeholders are replaced using the chart color-style palette when available and fall back through the workbook theme. `ChartStyleApplyReport` exposes `fillStylesAvailable`, `lineStylesAvailable`, and `effectStylesAvailable` in addition to series visit/style counts and diagnostics.


## P0Z — Structured-reference escaping, INDEX endpoint ranges and Office chart-style rules

Structured-reference parsing now follows Excel's apostrophe escaping for special header characters. Examples accepted by chart-cache/dependency resolution include `Table[Sales']Net]`, `Table['#Rate]`, `Table['@Code]`, and `Table[O''Brien]`. Worksheet quote parsing and structured-reference escaping use separate state, so an escaped table-header character is no longer misread as the start of a quoted worksheet name. Contiguous combinations such as `[[#Headers],[#Data],[Column]]` are resolved as one rectangular region; genuinely non-contiguous selections (for example headers plus totals with a non-empty data body) remain diagnostic rather than being flattened incorrectly.

Reference-form dynamic names can also use statically resolvable endpoints with the range operator:

```cpp
workbook.addDefinedName(DefinedName(
    "Window",
    "=INDEX('Data'!$B$2:$B$20,2,1):INDEX('Data'!$B$2:$B$20,8,1)"));
```

Each endpoint must reduce to one cell on the same worksheet. This is geometric reference resolution only; row/column arguments that themselves require formula calculation remain outside XL++'s calculation boundary.

Imported Office 2013 chart-style resources expose a first-class rule model through `ChartStyleResources::chartStyleRules`, `ChartStyleResources::rule(target)`, and `ChartStyleMarkerLayout`. `ChartStyleRule` retains `lnRef`, `fillRef`, `effectRef`, `lineWidthScale`, `fontRef`, explicit `spPr` fill/line overrides, style-color selectors, and ordered color transforms. Chart-color-style entries and their transforms preserve XML order because `styleClr="auto"` and DrawingML transform pipelines are order-sensitive.

```cpp
const auto& resources = chart.styleResources();
if (const auto* rule = resources.rule("dataPoint")) {
    rule->fillReference.index;
    rule->fillReference.styleColor;
    rule->lineWidthScale;
    rule->shapeFill;
}

auto report = workbook.applyChartStyleRules("Objects", chart.stableId());
report.rulesAvailable;
report.rulesVisited;
report.rulesApplied;
report.targetsStyled;
report.effectReferencesResolved;
```

`Workbook::applyChartStyleRules()` resolves theme/style references and materializes supported fill/line rules onto existing selective chart-edit APIs. Current targets include chart/plot area, data series and markers, legend, axes/gridlines, drop/high-low/leader/series lines, up/down bars, trendlines, error bars, data tables, floor and walls when those objects are present. Explicit `spPr` fill/line values override their matrix references, and `phClr` is materialized from the matching style reference/color-style entry. Parsed text/font rules and target effects are preserved/inspectable, but effect serialization to every ChartML target is not claimed yet.

`ChartThemeEffectStyle` now additionally exposes inner shadow, reflection and blur geometry alongside outer shadow, glow and soft edge. Effect references are resolved against the theme effect matrix and reported by `ChartStyleApplyReport`; deeper per-target effect serialization remains a follow-up refinement.

## Pivot shared caches, calculated fields and grouping (P1B)

Multiple generated PivotTables may intentionally share one physical PivotCache by assigning the same non-empty cache identity to compatible cache models:

```cpp
xlpp::PivotCache cache;
cache.setSharedCacheKey("sales-cache");
cache.setFields({"Region", "Sales"});
cache.addRecord({"East", "10"});
cache.addRecord({"West", "20"});

xlpp::PivotTable a("SalesA");
a.cache() = cache;
a.addRowField("Region");
a.addDataField("Sales", "sum");

xlpp::PivotTable b("SalesB");
b.cache() = cache; // same key + equivalent model => one physical cache part
b.addRowField("Region");
b.addDataField("Sales", "sum");
```

Using the same key with divergent source/options/fields/records/calculated formulas/grouping throws `std::invalid_argument` rather than emitting an inconsistent shared-cache graph.

Calculated fields and grouping:

```cpp
const int commission = cache.addCalculatedField("Commission", "Sales*0.1");
cache.setNumericFieldGrouping(1, 0.0, 1000.0, 100.0);
cache.setDateFieldGrouping(0, "months",
                           "2026-01-01T00:00:00",
                           "2026-12-31T23:59:59",
                           false, false);

xlpp::PivotFieldGroup group;
group.baseField = 0;
group.groupBy = "quarters";
cache.setFieldGroup(0, group);
```

`setDateFieldGrouping()` accepts `seconds`, `minutes`, `hours`, `days`, `months`, `quarters`, and `years`. `PivotField::addSubtotal()` exposes explicit subtotal flags such as `sum`, `avg`, `countA`, `max`, `min`, `product`, `count`, `stdDev`, `stdDevP`, `var`, and `varP`.

## VBA code names and extended references (P1B)

Worksheet document modules are keyed by stable worksheet VBA code names rather than by current worksheet position:

```cpp
auto& sheet = workbook.addWorksheet("Calculation");
sheet.setVbaCodeName("CalcSheet");
workbook.setVbaDocumentModuleText(
    "CalcSheet",
    "Private Sub Worksheet_Activate()\nEnd Sub");
```

Project locale/runtime metadata:

```cpp
xlpp::VbaProjectInfo info;
info.name = "AnalysisProject";
info.systemKind = 3;
info.lcid = 0x0409;
info.lcidInvoke = 0x0409;
info.codePage = 1252;
```

External VBA-project references use `VbaReferenceKind::Project` and carry absolute/relative LibIds plus the referenced project version. ActiveX/control-library references use `VbaReferenceKind::Control`:

```cpp
xlpp::VbaReference forms;
forms.name = "MSForms";
forms.kind = xlpp::VbaReferenceKind::Control;
forms.twiddledLibid = "*\\G{00000000-0000-0000-0000-000000000000}#0.0#0##";
forms.extendedName = "MSForms";
forms.libid = "*\\G{...}#2.0#0#...#Microsoft Forms 2.0 Object Library";
forms.originalTypeLib = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}";
forms.controlCookie = 1;
info.references.push_back(forms);
workbook.setVbaProjectInfo(info);
```

This is the binary-reference foundation for ActiveX/UserForm projects. P1C adds a first-class **raw** Designer Storage layer, but still does not claim semantic MS-OFORMS control authoring or arbitrary FRX property editing.


## PivotChart, filters and selective shared-cache mutation (P1C)

PivotTables can carry the SpreadsheetML side of a PivotChart link, while `Chart` carries the DrawingML `c:pivotSource` side:

```cpp
xlpp::PivotChartFormat format;
format.chartIndex = 7;
format.formatId = 0;
format.series = true;
format.pivotAreaXml = "<pivotArea type=\"normal\" dataOnly=\"1\"/>";
pivot.addChartFormat(format);
pivot.setChartFormatIndex(7);

xlpp::Chart chart(xlpp::Chart::Type::Bar);
chart.linkPivotTable("SalesPivot", 7);
```

`PivotChartFormat::pivotAreaXml` is intentionally lossless raw XML because the PivotArea selector grammar is much broader than the current high-level model. Existing PivotChart selectors therefore survive read/edit/save even when XL++ does not yet expose every selector field individually.

Pivot filters expose their common field/measure/text attributes while preserving nested advanced filter payloads:

```cpp
xlpp::PivotFilter filter;
filter.fieldIndex = 0;
filter.type = "captionContains";
filter.stringValue1 = "East";
filter.autoFilterXml =
    "<autoFilter ref=\"A1:A100\"><filters><filter val=\"East\"/></filters></autoFilter>";
pivot.addFilter(std::move(filter));
```

For an imported PivotTable whose cache has a physical package identity, common cache options can be patched without regenerating any sibling PivotTable that shares the cache:

```cpp
xlpp::PivotCacheOptionsPatch patch;
patch.refreshOnLoad = false;
patch.saveData = false;
patch.enableRefresh = true;
patch.missingItemsLimit = 17;

bool changed = workbook.updateImportedPivotCacheOptions(
    "Data", "SalesPivot", patch);
```

The patch updates the physical `pivotCacheDefinition` root and synchronizes every loaded Pivot model sharing that cache identity, while leaving the owning `pivotTable*.xml` parts untouched.

## VBA Designer modules and UserForm raw storage (P1C)

P1C distinguishes registered Designer modules from Standard/Class/Document modules and stores their binary designer state in a real recursive CFB root storage:

```cpp
xlpp::VbaDesignerStorage form;
form.name = "UserForm1";
form.streams.push_back({"f", {0x00, 0x01, 0x02}});
form.streams.push_back({"o", {0x10, 0x20}});
form.streams.push_back({"Controls/Nested/state", {0xAA, 0xBB}});

workbook.setVbaDesignerModule(
    "UserForm1",
    "Option Explicit\nPrivate Sub UserForm_Initialize()\nEnd Sub",
    form);
```

Designer source is exposed through the normal VBA module APIs with `VbaModuleType::Designer`. `VbaModule::designerClassId` retains the `PROJECT` `Package=` value and `designerBaseClass` retains the `VB_Base` identity.

Raw storages can be inspected or selectively replaced:

```cpp
auto storages = workbook.vbaDesignerStorages();
if (auto* stream = storages.front().findStream("f")) {
    // inspect bytes
}

workbook.setVbaDesignerStorage(storages.front());
workbook.removeVbaDesignerStorage("UserForm1");
```

The stream paths are relative to the designer root storage and may contain nested storage components separated by `/`. Unknown binary payloads remain byte-preserved. Removing a Designer module through `removeVbaModule()` also retires its matching root Designer Storage.

**Current boundary:** this is a preservation/editing foundation for UserForms and other registered designers, not a complete MS-OFORMS object model. XL++ does not yet decode every control/property stream, synthesize arbitrary controls, edit signatures, unlock protected VBA projects or reproduce arbitrary compiled p-code.

## Selective Pivot field editing and PivotChart validation (P1D)

An imported PivotCache field can be patched in place through its physical shared-cache identity without regenerating sibling PivotTable definitions:

```cpp
xlpp::PivotCacheFieldPatch patch;
patch.name = "Revenue";
patch.caption = "Net Revenue";
patch.formula = "Sales*1.05";
patch.numberFormatId = 4;
patch.databaseField = false;

bool changed = workbook.updateImportedPivotCacheField(
    "Data", "SalesPivot", 1, patch);
```

All loaded PivotTable models that point at the same physical cache are synchronized with the patched field metadata. Common field-item state is also modeled:

```cpp
xlpp::PivotFieldItem item;
item.cacheIndex = 0;
item.type = "data";
item.caption = "East region";
item.hidden = true;
item.showDetails = false;
pivot.rowFields().front().addItem(item);
```

Use the ownership validator before relying on imported or edited PivotChart links:

```cpp
const auto report = workbook.validatePivotChartLinks();
if (!report.ok()) {
    for (const auto& issue : report.issues) {
        // issue.worksheetName, issue.chartId,
        // issue.pivotTableName, issue.message
    }
}
```

The validator checks source resolution and chart-format identity coherence; it does not attempt to recalculate Pivot results.

## Semantic UserForm Form properties (P1D)

P1D adds a semantic layer above the raw Designer Storage for the MS-OFORMS Form stream named `f`:

```cpp
const auto form = workbook.inspectVbaUserForm("UserForm1");
if (form.valid) {
    auto caption = form.properties.caption;
    auto width   = form.properties.displayedWidth;
    auto height  = form.properties.displayedHeight;
    auto zoom    = form.properties.zoom;
}
```

Common already-materialized form properties can be patched without rewriting unrelated Designer streams:

```cpp
xlpp::VbaUserFormPropertiesPatch patch;
patch.caption = "Analysis ✓";
patch.backColor = 0x80000005u;
patch.displayedWidth = 6400;
patch.displayedHeight = 3600;
patch.logicalWidth = 8000;
patch.logicalHeight = 5000;
patch.scrollLeft = 321;
patch.scrollTop = 654;
patch.zoom = 150u;
patch.drawBuffer = 40000u;

workbook.updateVbaUserFormProperties("UserForm1", patch);
```

Caption rewrites may change encoded length and can move from the compressed single-byte representation to UTF-16. XL++ rebuilds the semantic Form block while retaining trailing FormStreamData/SiteData and sibling `o`, `vbFrame` or nested control streams unchanged.

For safety, P1D only edits properties whose corresponding Form property-mask bit is already present. It does not silently synthesize absent MS-OFORMS fields with guessed defaults.

Designer ownership can be checked explicitly:

```cpp
const auto report = workbook.validateVbaDesignerProject();
if (!report.ok()) {
    for (const auto& issue : report.issues) {
        // issue.designerName, issue.message
    }
}
```

**Current UserForm boundary after P1E:** Form-level properties and child-site metadata are modeled; individual object-stream control classes are still exposed losslessly as byte slices rather than being semantically rewritten.


## Selective imported Pivot item/filter/cache-record editing (P1E)

P1E extends preservation-safe Pivot mutation below cache-field metadata.

Patch one imported Pivot field item without regenerating the complete PivotTable:

```cpp
xlpp::PivotFieldItemPatch itemPatch;
itemPatch.caption = "Eastern region";
itemPatch.hidden = true;
itemPatch.showDetails = false;
itemPatch.formula = true;

workbook.updateImportedPivotFieldItem(
    "Data", "SalesPivot", 0, 0, itemPatch);
```

Patch an existing Pivot filter, including its nested SpreadsheetML `autoFilter` subtree:

```cpp
xlpp::PivotFilterPatch filterPatch;
filterPatch.type = "captionBeginsWith";
filterPatch.evaluationOrder = 4;
filterPatch.name = "Eastern filter";
filterPatch.stringValue1 = "Ea";
filterPatch.autoFilterXml =
    "<autoFilter ref=\"A1:A4\">"
    "<customFilters><customFilter operator=\"beginsWith\" val=\"Ea\"/>"
    "</customFilters></autoFilter>";

workbook.updateImportedPivotFilter(
    "Data", "SalesPivot", 0, filterPatch);
```

Patch one physical `pivotCacheRecords` value without regenerating its cache definition or owner PivotTables:

```cpp
xlpp::PivotCacheRecordValuePatch value;
value.type = xlpp::PivotCacheRecordValueType::Number;
value.value = "15.5";

workbook.updateImportedPivotCacheRecordValue(
    "Data", "SalesPivot", 0, 1, value);
```

Supported physical cache-record value kinds are `Missing`, `Number`, `String`, `Boolean`, `Error`, `DateTime` and `SharedItem`. For `SharedItem`, supply both the resolved logical `value` and `sharedItemIndex`.

These APIs deliberately fail instead of switching to full Pivot regeneration when the requested imported physical part/field/item/record cannot be located.

## UserForm child control sites and object-stream slicing (P1E)

P1E parses the MS-OFORMS `FormSiteData` tail of a Designer Form stream and exposes each `OleSiteConcreteControl` site:

```cpp
const auto form = workbook.inspectVbaUserFormControls("UserForm1");
if (form.valid) {
    for (const auto& control : form.controls) {
        auto name = control.name;
        auto id = control.id;
        auto tabIndex = control.tabIndex;
        auto positionTop = control.top;
        auto positionLeft = control.left;
        auto bytes = control.objectData;
    }
}
```

The site model exposes depth/type/version/property mask, Name, Tag, ID, HelpContextID, BitFlags, ObjectStreamSize, TabIndex, ClsidCacheIndex, GroupID, Position, tooltip/runtime-license/control-source/row-source values when present.

`ObjectStreamSize` values are also mapped onto the Designer Storage `o` stream. Each control receives `objectStreamOffset` and a lossless `objectData` slice. `objectStreamBytes` and `unassignedObjectStreamBytes` make incomplete or trailing object-stream data visible to callers.

Safe site metadata can be patched in place:

```cpp
xlpp::VbaUserFormControlSitePatch patch;
patch.name = "RunButton";
patch.tag = "primary";
patch.helpContextId = 88;
patch.bitFlags = 0x31;
patch.tabIndex = 7;
patch.groupId = 4;
patch.controlTipText = "Run analysis";
patch.controlSource = "C3";
patch.rowSource = "C3:C8";
patch.top = 333;
patch.left = 444;

workbook.updateVbaUserFormControlSite("UserForm1", 0, patch);
```

Variable-length site strings can grow or switch to UTF-16; XL++ updates the site record length and enclosing `FormSiteData` byte count while preserving the `o` stream and unrelated nested Designer streams byte-for-byte.

For safety, P1E only edits site properties whose `SitePropMask` bits are already materialized. Semantic parsing/writing of individual `CommandButton`, `TextBox`, `Label`, list controls and other object-stream classes remains a later layer.

## Pivot data/page field selective mutation (P1F)

P1F extends the imported-Pivot selective-edit path to data and page fields without regenerating the physical PivotCache:

```cpp
xlpp::PivotDataFieldPatch data;
data.name = "Average sales";
data.subtotal = "average";
data.showDataAs = "percentOfTotal";
data.baseField = 0;
data.baseItem = 0;
data.numberFormatId = 4;
workbook.updateImportedPivotDataField("Data", "SalesPivot", 0, data);

xlpp::PivotPageFieldPatch page;
page.item = 1;
page.hierarchy = 2;
page.name = "Fiscal year";
workbook.updateImportedPivotPageField("Data", "SalesPivot", 0, page);
```

The APIs patch only the selected imported `dataField`/`pageField` record and update the loaded model. They fail instead of silently switching to whole-Pivot regeneration when the requested owner/record cannot be resolved.

## UserForm object-specific control inspection/editing (P1F)

P1F adds built-in MSForms control classification from `ClsidCacheIndex` and the first object-specific semantic layer for `CommandButton` and `Label` Designer `o`-stream slices.

```cpp
auto info = workbook.inspectVbaUserFormControlObject("ButtonForm", 0);

xlpp::VbaUserFormControlObjectPatch patch;
patch.caption = "Run \xE2\x9C\x93";
patch.width = 2200;
patch.height = 700;
workbook.updateVbaUserFormControlObject("ButtonForm", 0, patch);
```

`VbaUserFormControlKind` identifies common cached built-ins including Form, Image, Frame, MorphData, SpinButton, CommandButton, TabStrip, Label, TextBox, ListBox, ComboBox, CheckBox, OptionButton, ToggleButton, ScrollBar and MultiPage. Custom-class cache indexes are surfaced as `CustomClass`.

For `CommandButton` and `Label`, XL++ can inspect/preserve common header/property-mask data, colors, caption, size and selected style/accelerator fields. Caption growth and compressed-to-UTF-16 conversion update the control `cbControl` length and the owning site's `ObjectStreamSize`; unrelated trailing object bytes and sibling Designer streams remain lossless.

`TextBox`/MorphData and other families are classified and common-header validated in P1F, but semantic object-property mutation is intentionally limited until their family-specific binary layouts are implemented.



## P1J binding-safe worksheet lifecycle

The C ABI now exposes `xlpp_workbook_rename_sheet()` and catches exceptions from `xlpp_workbook_remove_sheet()` so the last-sheet invariant never allows a C++ exception to cross the C boundary. Python exposes `rename_worksheet()`, and C# exposes `RenameWorksheet()`.


## P1S chart generation additions

`ChartSeries` adds `bubbleSizeReference()` / `setBubbleSizeReference()` and a dedicated bubble-size cache. `Chart` adds `scatterStyle()` / `setScatterStyle()` plus generated multi-plot helpers `addPlot()`, `addSeriesToPlot()` and `clearPlots()`.

```cpp
xlpp::Chart combo(xlpp::Chart::Type::Bar);
combo.addPlot(xlpp::Chart::Type::Bar, xlpp::Chart::Grouping::Clustered, false);
combo.addSeriesToPlot(0, primary);
combo.addPlot(xlpp::Chart::Type::Line, xlpp::Chart::Grouping::Standard, true);
combo.addSeriesToPlot(1, secondary);
```

Scatter/Bubble generated series use the XY-family OOXML representation. Bubble series require an explicit bubble-size reference/cache when generated.

## P1S typed Pivot cache values

`PivotCacheValueKind` describes the physical SpreadsheetML cache value kind: `Missing`, `Number`, `String`, `Boolean`, `Error`, or `DateTime`.

```cpp
cache.setTypedRecords(
    {{"00123", "123", "2026-08-14T09:30:00", "true", ""}},
    {{xlpp::PivotCacheValueKind::String,
      xlpp::PivotCacheValueKind::Number,
      xlpp::PivotCacheValueKind::DateTime,
      xlpp::PivotCacheValueKind::Boolean,
      xlpp::PivotCacheValueKind::Missing}});
```

Use `addTypedRecord()` or `setRecordValue(..., kind)` when exact physical cache types matter. The legacy mutable `records()` accessor invalidates typed-kind metadata because direct string mutation does not specify a physical value kind. Use `hasTypedRecordKinds()`, `recordKinds()` and `recordKind()` for inspection.

## P1T mixed workbook sheet model and Chartsheets

Legacy `Workbook::sheetCount()`, `sheetNames()` and worksheet indexing remain worksheet-only. Use the mixed-tab APIs when workbook order must include chart-only sheets:

```cpp
xlpp::Workbook wb;
auto& data = wb.addWorksheet("Data");
xlpp::Chart chart(xlpp::Chart::Type::Bar);
chart.setTitle("Dashboard");
wb.addChartsheet("Chart View", std::move(chart));
wb.moveWorkbookSheet(1, 0);

for (const auto& tab : wb.workbookSheets()) {
    // tab.kind is Worksheet or Chartsheet; tab.kindIndex indexes that collection.
}
```

Chartsheet APIs: `addChartsheet`, `chartsheet`, `renameChartsheet`, `removeChartsheet`, `chartsheetCount`, `chartsheets`, `workbookSheetCount`, `workbookSheetNames`, `workbookSheets`, `moveWorkbookSheet`. `LoadOptions::maxChartsheets` bounds materialized chart-only sheets. Streaming readers expose `chartsheetNames()` and `workbookSheetNames()`.

## P1T Pivot logical cache-item identity

Prefer value-bound field items when the cache may be edited:

```cpp
auto& field = pivot.addRowField("Category");
field.hideCacheValue("B", xlpp::PivotCacheValueKind::String);
```

`PivotFieldItem::bindCacheValue()` and `PivotField::{addCacheValueItem,hideCacheValue,showCacheValue}` bind semantics to `(kind,value)`. On save, XL++ resolves the current physical shared-item index and writes the repaired `x` ordinal. Legacy `cacheIndex` remains supported for callers that intentionally address the physical cache table.

## P1T C ABI additions

The C ABI adds mixed-tab and Chartsheet functions (`xlpp_workbook_tab_*`, `xlpp_workbook_*chartsheet*`, `xlpp_chartsheet_chart`) and chart-generation helpers for scatter style, generated plots/secondary axes, series-to-plot insertion and bubble-size references. Legacy C worksheet functions keep their original worksheet-only meaning.


## P1U template and Chartsheet production APIs

### Mixed-tab visibility and active state

```cpp
xlpp::Workbook wb;
wb.addWorksheet("Data");
wb.addChartsheet("Dashboard", chart);
wb.addWorksheet("Hidden Data");

wb.setActiveWorkbookSheet("Dashboard");
wb.setWorkbookSheetVisibility(2, xlpp::WorkbookSheetVisibility::Hidden);

auto active = wb.activeWorkbookSheetIndex();
auto state = wb.workbookSheetVisibility(2);
```

`WorkbookSheetVisibility` has `Visible`, `Hidden` and `VeryHidden`. A workbook cannot hide its final visible tab. `workbookSheets()` reports kind, kind-local index, visibility and active state while legacy `sheetCount()` / `sheetNames()` remain worksheet-only.

### Chartsheet metadata

```cpp
auto& cs = wb.addChartsheet("Dashboard", chart);
cs.properties().setCodeName("ChartDashboard");
cs.properties().setTabColor("FF336699");
cs.view().setZoomScale(125);
cs.view().setZoomToFit(false);
cs.protection().setContent(true);
cs.protection().setObjects(true);
cs.protection().setPassword("pw");
cs.pageSetup().setOrientation(xlpp::PageOrientation::Landscape);
cs.pageSetup().setScale(85);
cs.headerFooter().setOddHeader("&CReport");
cs.headerFooter().setFirstHeader("&LFirst page");
```

Mutable sheet-metadata access marks only the Chartsheet part dirty. Mutable `chart()` access marks the chart subtree dirty. For preservation-backed imported Chartsheets this distinction lets sheet metadata change while original drawing/chart bytes remain untouched.

### Templates

```cpp
wb.setTemplate(true);
wb.save("dashboard.xltx");

wb.setVbaModuleText("Module1", "Sub Hello()\r\nEnd Sub\r\n");
wb.save("dashboard.xltm");
```

Template identity is recovered automatically on load. `.xltm` identity is the combination of template mode and a VBA project.

### C ABI additions

P1U adds `xlpp_workbook_tab_visibility`, `xlpp_workbook_set_tab_visibility`, `xlpp_workbook_active_tab`, `xlpp_workbook_set_active_tab`, `xlpp_workbook_set_template` and `xlpp_workbook_is_template`.


## P1V advanced Chartsheet page/view APIs

### Advanced page setup

```cpp
auto& setup = chartsheet.pageSetup();
setup.setPaperHeight("210mm");
setup.setPaperWidth("297mm");
setup.setPageOrder(xlpp::PageOrder::OverThenDown);
setup.setUsePrinterDefaults(false);
setup.setCellComments(xlpp::PageCellComments::AtEnd);
setup.setErrors(xlpp::PageErrorDisplay::Dash);
setup.setHorizontalDpi(600);
setup.setVerticalDpi(600);
setup.setCopies(3);
```

`PageSetup` now carries these settings for both Worksheets and Chartsheets. Imported printer-settings relationship IDs can be retained through `relationshipId()` / `setRelationshipId()`. Generated callers should not invent a printer-settings relationship ID without also owning the corresponding package part; model validation reports that situation.

### Modern Chartsheet protection metadata

```cpp
auto& protection = chartsheet.protection();
protection.setContent(true);
protection.setObjects(true);
protection.setAlgorithmName("SHA-512");
protection.setHashValue("AQIDBA==");
protection.setSaltValue("BQYHCA==");
protection.setSpinCount(100000);
```

These APIs preserve the OOXML protection descriptor; they do not calculate a new password hash from plaintext.

### Custom Chartsheet views

```cpp
auto& view = chartsheet.customViews().emplace_back();
view.setGuid("{11111111-2222-3333-4444-555555555555}");
view.setScale(90);
view.setState(xlpp::CustomChartsheetViewState::Hidden);
view.setZoomToFit(false);
view.pageSetup().setPaperHeight("210mm");
view.headerFooter().setOddHeader("&CAlternate view");
```

Custom views retain GUID, scale, visibility state and zoom-to-fit state plus optional nested `PageMargins`, `PageSetup` and `HeaderFooter`. Untouched imported `customSheetViews` XML is retained losslessly; mutable `customViews()` access opts into semantic regeneration. Duplicate or empty GUIDs are rejected by save/model validation.

### I/O decomposition

Chartsheet XML parsing/serialization and chart-only drawing relationship helpers live in `src/XLPP/Workbook/WorkbookChartsheetIO.cpp`. This reduces the amount of the monolithic workbook serializer that must be recompiled for Chartsheet-only changes and gives future parser/writer hardening a narrower boundary.


## P1X-A internal package-writer note

P1X-A does not add a public API. The `.xltx/.xltm` Chartsheet writer has been moved to an internal `WorkbookChartsheetPackage` subsystem. Relationship-ID collision repair is now relationship-owner aware so generated drawing/printer-settings IDs cannot rewrite unrelated preserved Chartsheet owner nodes.
