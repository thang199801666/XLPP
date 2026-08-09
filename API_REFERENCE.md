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
wb.removeWorksheet("Sheet1");               // remove; dependent local refs become #REF!
wb.copyWorksheet(sheet, "Copy");            // copy existing sheet
auto renamed = wb.renameWorksheet("Sheet1", "Input Data"); // dependency-aware rename

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
auto* local = wb.definedName("SalesData", 0);                 // scoped lookup
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
opts.atomicWrite = true;                     // default: stage + atomic replace for path saves
opts.durableWrite = true;                    // default: flush file + directory/write-through barrier
opts.validateBeforeSave = true;              // default: validate modeled invariants
wb.save("output.xlsx", opts);
```

### Model validation

```cpp
auto validation = wb.validate();
if (!validation.ok()) {
    for (const auto& issue : validation.issues)
        std::cerr << issue.code << ": " << issue.message << "\n";
}
```

Path-based `load()` provides a strong exception guarantee: if parsing/decryption fails, the caller's existing workbook remains unchanged. Path-based `save()` uses same-directory staging and atomic replacement by default. `SaveOptions::durableWrite` also defaults to true, flushing the completed file and using a POSIX parent-directory `fsync` or Windows write-through replacement where applicable. Set either option to false only for an intentional performance/semantics tradeoff.

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

### In-process formula calculation (P0W)

```cpp
xlpp::CalculationOptions options;
options.recursiveDependencies = true;
options.updateCachedValues = true;
options.evaluateVolatileFunctions = true;
options.maxDepth = 512;

auto report = workbook.calculateFormulas(options);
report.formulaCellsEvaluated;
report.cachedValuesUpdated;
report.circularReferences;
report.unsupportedFormulas;
report.warnings;
```

`Workbook::calculateFormulas()` updates formula cached values but never rewrites
formula text. The P0W engine covers common aggregate/statistical, logical/error,
math/trigonometric, text, date/time, criteria, lookup and financial functions;
it does not yet implement dynamic arrays/spill, structured references, external
workbook references, `INDIRECT`/`OFFSET`, or iterative circular calculation.

`SaveOptions::calculateFormulasBeforeSave = true` evaluates a private workbook
copy before serialization. When chart-cache synchronization is also enabled,
formula calculation runs first.

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
xlpp::StreamingReaderOptions limits;
limits.maxFileBytes = 8ull * 1024 * 1024 * 1024;
limits.maxEntryBytes = 2ull * 1024 * 1024 * 1024;
limits.maxTotalBytes = 16ull * 1024 * 1024 * 1024;
limits.maxEntries = 100000;
limits.validateCellReferences = true;
xlpp::StreamingWorkbookReader reader("large.xlsx", limits);
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

## VBA / Macro Projects

XL++ supports two VBA ownership modes: source projects generated by XL++, and opaque external `vbaProject.bin` projects that are preserved without destructive rewriting.

```cpp
Workbook wb;
auto& ws = wb.addWorksheet("Data");
ws.setVbaCodeName("DataSheet");

VbaProjectProperties props;
props.name = "MyMacros";
props.description = "XL++ generated VBA";
props.helpFile = "help.chm";
props.helpContextId = 10;
props.constants = "FeatureX = 1:DebugMode = 0";
wb.setVbaProjectProperties(props);

wb.setVbaModuleText("Utilities", "Public Sub Run()\nEnd Sub\n");
wb.setVbaClassModuleText("Worker", "Public Function Value() As Long\nValue = 1\nEnd Function\n",
                         true, true);
wb.setVbaDocumentModuleText("ThisWorkbook",
                            "Private Sub Workbook_Open()\nEnd Sub\n");
wb.setVbaDocumentModuleText("DataSheet",
                            "Private Sub Worksheet_Activate()\nEnd Sub\n");

for (const auto& module : wb.vbaModules()) {
    module.name;
    module.source;
    module.type;          // Standard / Document / Class
    module.readOnly;
    module.privateModule;
}

auto binary = wb.vbaProjectBytes();
wb.saveVbaProject("vbaProject.bin");
bool signedProject = wb.hasVbaSignature();
bool editable = wb.vbaSourceEditable();
```

`Worksheet::vbaCodeName()` is independent from the worksheet display name and is serialized as `sheetPr/@codeName`. Generated code names are stable across normal workbook topology changes.

`setVbaModuleText`, `setVbaClassModuleText`, `setVbaDocumentModuleText`, `setVbaModule`, `removeVbaModule` and `setVbaProjectProperties` rebuild the source-owned project. They reject externally supplied projects and projects carrying a VBA signature because rebuilding could discard or invalidate opaque VBA state.

`addVbaProject()` and `setVbaProject()` intentionally establish external/binary ownership. The attached bytes can be inspected/exported and are preserved by unrelated workbook edits, but source mutation is not enabled.

Current source generation does not create UserForms/FRX designer state, custom type-library references, VBA project passwords/locks or digital signatures.

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

## P0X — 90% General-purpose Editing Core

### Dependency graph and structural transactions

```cpp
auto graph = workbook.dependencyGraph();
auto users = graph.dependentsOf("Data!A1:A10");

xlpp::StructuralEditOptions editOptions;
editOptions.transactional = true;
editOptions.failOnInvalidReference = true;
auto edit = workbook.insertRows("Data", 3, 2, editOptions);
```

P0X structural transactions rewrite supported formula/name/table/filter/CF/DV/print/chart/Pivot/drawing/hyperlink dependencies and roll back when configured invalid references are introduced.

### Modern formula semantics

`CalculationOptions` now includes iterative calculation controls and an optional C++ external-reference resolver. Dynamic arrays/spills, structured table references, `LET`, `INDIRECT`, `OFFSET` and common array-shaping families are supported within the scope documented in `docs/P0X_90_PERCENT_ENGINE.md`.

```cpp
xlpp::CalculationOptions calc;
calc.iterativeCalculation = true;
calc.maxIterations = 100;
calc.maxChange = 1e-9;
auto report = workbook.calculateFormulas(calc);
```

### Standard or Agile password encryption

```cpp
xlpp::SaveOptions save;
save.encryptionPassword = "secret";
save.encryptionMode = xlpp::OfficeEncryptionMode::Standard;
save.encryptionKeyBits = 256;
workbook.save("secure.xlsx", save);
```

P0X supports Agile AES-256/SHA-512 read/write and Standard AES-128/192/256 + SHA-1 read/write. Certificate/private-key encryption and alternate Agile profiles remain outside the current core.

### Pivot semantic editing

Imported PivotTables remain byte-preserved while untouched. Mutable access opts into regeneration of the supported semantic model, including common source/location, fields, layout/style, cache lifecycle, data-field and report-filter properties.


## P0Y — Core Reliability and Topology

### Dependency-aware worksheet rename

```cpp
xlpp::WorksheetRenameOptions renameOptions;
renameOptions.recalculateFormulas = false;
renameOptions.synchronizeChartCaches = true;
auto report = workbook.renameWorksheet("Raw Data", "Input Data", renameOptions);
```

Qualified formula/name/chart/Pivot/CF/DV/internal-hyperlink references are rewritten without altering formula string literals or external-workbook references. Removing a worksheet invalidates supported local dependencies to `#REF!` and compacts local defined-name scopes. Because worksheet storage is currently a `std::deque`, erasing a worksheet may invalidate existing worksheet handles; reacquire them after removal.

### Core reliability build gates

```bash
cmake -S . -B build-asan -DXLPP_ENABLE_SANITIZERS=ON
cmake --build build-asan
ctest --test-dir build-asan --output-on-failure

cmake -S . -B build-warn -DXLPP_ENABLE_STRICT_WARNINGS=ON
cmake --build build-warn
```

Continuous parser/package fuzzing (Clang/libFuzzer):

```bash
CC=clang CXX=clang++ cmake -S . -B build-fuzz \
  -DXLPP_BUILD_FUZZERS=ON -DXLPP_BUILD_TESTS=OFF \
  -DXLPP_BUILD_SAMPLES=OFF -DXLPP_BUILD_TOOLS=OFF
cmake --build build-fuzz --target XLPP_WorkbookLoadFuzzer
./build-fuzz/tests/fuzz/XLPP_WorkbookLoadFuzzer corpus/ -max_total_time=60
```

P0Z-G release validation builds the complete native core with this strict-warning profile at **0 warnings / 0 errors** on the release GCC toolchain; safety-critical diagnostics remain promoted to errors.

## Build

Open `XL++.sln` with Visual Studio 2022+ (Platform Toolset v145), select x64
Debug or Release, then build.

```bash
# The solution builds:
#   - XLPP.lib          (static library)
#   - XLPP.UnitTests    (test runner; current P0X suite count is recorded in PACKAGE_STATUS.md)
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


---

## P0V — Dependency-aware Chart Cache Synchronization

P0V extends `Workbook::synchronizeChartCaches()` with exact source-dependency snapshots. A snapshot includes the normalized chart reference, source worksheet, cell coordinates, stored cell values, formula text, number format and workbook date epoch.

```cpp
xlpp::ChartCacheSyncOptions options;
options.changedReferencesOnly = true;

auto report = workbook.synchronizeChartCaches(options);
report.referencesChecked;
report.referencesUnchanged;
report.dependenciesRegistered;
report.dependenciesChanged;
report.cachesUpdated;
```

The first changed-only synchronization has no prior snapshot, so all supported references are synchronized and registered. Later calls skip `strCache`/`numCache` rebuilds when the dependency snapshot is unchanged. Editing a cell outside every tracked chart range does not invalidate any chart dependency.

```cpp
workbook.trackedChartCacheDependencyCount();
workbook.resetChartCacheDependencyTracking();
```

A save can opt into cache synchronization without mutating the caller workbook. XL++ creates a private copy, synchronizes that copy, and serializes it:

```cpp
xlpp::SaveOptions saveOptions;
saveOptions.synchronizeChartCaches = true;
saveOptions.synchronizeChangedChartCachesOnly = true;
workbook.save("output.xlsx", saveOptions);
```

`SaveOptions::synchronizeChartCaches` defaults to `false`, preserving the P0U save behavior. Set `synchronizeChangedChartCachesOnly=false` to force a full supported-reference rebuild even when dependency snapshots are available.

### P0V-A scope

This batch tracks direct chart A1 dependencies. P0W subsequently added formula calculation, but the chart snapshot tracker is still not a general transitive dependency graph and structural row/column edits still do not rewrite every dependent reference. Those remain P1 work.

---

## P0W — Password-to-open OOXML Encryption

```cpp
xlpp::SaveOptions save;
save.encryptionPassword = "secret";
save.encryptionSpinCount = 100000;
workbook.save("secure.xlsx", save);

xlpp::LoadOptions load;
load.password = "secret";
load.verifyEncryptionIntegrity = true;
workbook.load("secure.xlsx", load);

auto info = xlpp::inspectOfficeEncryption("secure.xlsx");
```

XL++ writes Office Agile AES-256-CBC/SHA-512 encryption with HMAC integrity and reads both that profile and Standard AES/SHA-1 (128/192/256-bit key paths). The CFB implementation includes DataSpaces, miniFAT/FAT/DIFAT and large-file support. Saving a loaded encrypted workbook with a different password rotates the password; saving it without `encryptionPassword` removes file-open encryption. Standard-encryption output, alternate Agile profiles and certificate key encryptors remain unsupported.

---

## P0Z-I / v1.12.0 — Formula, Enterprise Inspection and Capability APIs

### Dependency-driven recalculation

```cpp
xlpp::CalculationOptions options;
options.changedCells.push_back({"Data", "A1"});
auto report = workbook.calculateFormulas(options);

std::cout << report.dirtyRoots << " roots\n";
std::cout << report.dirtyFormulaCellsSelected << " formulas selected\n";
```

When `changedCells` is non-empty, XL++ selects formula cells transitively dependent on those roots instead of visiting every formula cell. An explicitly dirty formula cell is also selected as a calculation root.

Typed workbook calculation mode:

```cpp
workbook.calcProperties().setCalculationMode(xlpp::CalculationMode::Automatic);
workbook.calcProperties().setCalculationMode(xlpp::CalculationMode::AutomaticExceptDataTables);
workbook.calcProperties().setCalculationMode(xlpp::CalculationMode::Manual);
```

### Advanced AutoFilter

```cpp
auto& column = sheet.autoFilter().column(0);

xlpp::Top10Filter top;
top.top = true;
top.percent = false;
top.value = 10;
column.setTop10Filter(top);

xlpp::DynamicFilter dynamic;
dynamic.type = xlpp::DynamicFilterType::ThisMonth;
sheet.autoFilter().column(1).setDynamicFilter(dynamic);

xlpp::ColorFilter color;
color.dxfId = 3;
color.cellColor = true;
sheet.autoFilter().column(2).setColorFilter(color);

xlpp::IconFilter icon;
icon.iconSet = "4Arrows";
icon.iconId = 2;
sheet.autoFilter().column(3).setIconFilter(icon);

xlpp::DateGroupItem month;
month.year = 2026;
month.month = 8;
month.grouping = xlpp::DateTimeGrouping::Month;
sheet.autoFilter().column(4).addDateGroup(month);
```

### External Data inspection

```cpp
const auto external = workbook.inspectExternalData();
for (const auto& connection : external.connections)
    std::cout << connection.id << ": " << connection.name << "\n";

for (const auto& query : external.queryTables)
    std::cout << query.name << " -> connection " << query.connectionId << "\n";
```

The inspection model is deliberately preservation-first. Reading metadata does not imply that XL++ will regenerate opaque connection or Power Query payloads.

### Data Model / OLAP inspection

```cpp
const auto model = workbook.inspectDataModel();
if (model.present) {
    std::cout << "model parts: " << model.modelParts.size() << "\n";
    std::cout << "OLAP pivot caches: " << model.olapPivotCacheParts.size() << "\n";
}
```

XL++ currently inventories these proprietary/enterprise parts and preserves them where possible; semantic Data Model authoring is not claimed.

### Version and C ABI capability negotiation

C++:

```cpp
#include <XLPP/Version.h>
static_assert(xlpp::versionMajor == 1);
std::cout << xlpp::versionString << "\n";
std::cout << xlpp::cAbiVersion << "\n";
```

C:

```c
printf("XL++ %s, ABI %llu\n", xlpp_version(),
       (unsigned long long)xlpp_c_abi_version());
uint64_t caps = xlpp_capabilities();
if (caps & XLPP_CAP_DIRTY_RECALC) {
    /* runtime supports the v1.12 dirty-recalculation capability */
}
```

C ABI feature bits are additive: existing bits do not change meaning in later compatible releases.
