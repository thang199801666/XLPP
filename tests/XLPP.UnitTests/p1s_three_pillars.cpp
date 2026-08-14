#include "Packaging/ZipArchive.h"
#include <XLPP/XLPP.h>

#include <chrono>
#include <cstdlib>
#include <filesystem>
#include <iostream>
#include <string>
#include <vector>

namespace {
int failures = 0;

void check(bool ok, const char* message) {
    if (!ok) {
        ++failures;
        std::cerr << "FAIL: " << message << '\n';
    }
}

template <class T, class U>
void checkEqual(const T& actual, const U& expected, const char* message) {
    check(actual == expected, message);
}

bool keepArtifacts() { return std::getenv("XLPP_KEEP_P1S_ARTIFACTS") != nullptr; }

void cleanupOrReport(const std::filesystem::path& path) {
    if (keepArtifacts()) std::cout << "P1S_ARTIFACT=" << path.string() << '\n';
    else std::filesystem::remove(path);
}

std::filesystem::path tempPath(const char* stem, const char* ext = ".xlsx") {
    const auto tick = std::chrono::high_resolution_clock::now().time_since_epoch().count();
    return std::filesystem::temp_directory_path() /
        (std::string(stem) + "_" + std::to_string(tick) + ext);
}

void testTemplateRoundTrip() {
    const auto first = tempPath("xlpp_p1s_template", ".xltx");
    const auto second = tempPath("xlpp_p1s_template_roundtrip", ".xltx");
    xlpp::Workbook wb;
    wb.setTemplate(true);
    auto& sheet = wb.addWorksheet("TemplateData");
    sheet.cell("A1").setValue("template-sentinel");
    wb.save(first);

    const auto archive = xlpp::internal::ZipArchive::open(first);
    const auto types = archive.get("[Content_Types].xml");
    check(types.find("application/vnd.openxmlformats-officedocument.spreadsheetml.template.main+xml") != std::string::npos,
          "XLTX workbook content type is emitted");

    xlpp::Workbook loaded;
    loaded.load(first);
    check(loaded.isTemplate(), "XLTX template identity is detected on load");
    const auto* loadedSheet = static_cast<const xlpp::Workbook&>(loaded).worksheet("TemplateData");
    const auto* templateCell = loadedSheet ? loadedSheet->tryCell("A1") : nullptr;
    check(templateCell && templateCell->stringValueOr({}) == "template-sentinel",
          "Template cell content round-trips");
    loaded.save(second);
    const auto secondArchive = xlpp::internal::ZipArchive::open(second);
    check(secondArchive.get("[Content_Types].xml").find("application/vnd.openxmlformats-officedocument.spreadsheetml.template.main+xml") != std::string::npos,
          "XLTX template identity survives load/save round-trip");
    cleanupOrReport(first);
    cleanupOrReport(second);

    const auto macroTemplate = tempPath("xlpp_p1s_macro_template", ".xltm");
    xlpp::Workbook macro;
    macro.setTemplate(true);
    macro.addWorksheet("MacroTemplate");
    macro.setVbaModuleText("Module1", "Sub Hello()\r\nEnd Sub\r\n");
    macro.save(macroTemplate);
    const auto macroArchive = xlpp::internal::ZipArchive::open(macroTemplate);
    check(macroArchive.get("[Content_Types].xml").find("application/vnd.ms-excel.template.macroEnabled.main+xml") != std::string::npos,
          "XLTM macro-enabled template content type is emitted");
    check(macroArchive.contains("xl/vbaProject.bin"), "XLTM generated template includes VBA project bytes");
    xlpp::Workbook macroLoaded;
    macroLoaded.load(macroTemplate);
    check(macroLoaded.isTemplate() && macroLoaded.hasVbaProject(),
          "XLTM template and VBA identities survive load");
    cleanupOrReport(macroTemplate);
}

xlpp::ChartSeries makeSeries(std::string title, std::string x, std::string y) {
    xlpp::ChartSeries series(std::move(title));
    series.setCategoriesReference(std::move(x));
    series.setValuesReference(std::move(y));
    return series;
}

void testScatterBubbleAndComboGeneration() {
    const auto path = tempPath("xlpp_p1s_charts");
    xlpp::Workbook wb;
    auto& sheet = wb.addWorksheet("Data");
    sheet.append({std::string("X"), std::string("Y1"), std::string("Y2"), std::string("Size")});
    for (int i = 1; i <= 5; ++i)
        sheet.append({static_cast<double>(i), static_cast<double>(i * 2), static_cast<double>(i * 3), static_cast<double>(i + 4)});

    xlpp::Chart scatter(xlpp::Chart::Type::Scatter);
    scatter.setTitle("Scatter");
    scatter.setScatterStyle("smoothMarker");
    scatter.addSeries(makeSeries("Y1", "'Data'!$A$2:$A$6", "'Data'!$B$2:$B$6"));
    sheet.addChart(scatter);

    xlpp::Chart bubble(xlpp::Chart::Type::Bubble);
    bubble.setTitle("Bubble");
    auto bubbleSeries = makeSeries("Y2", "'Data'!$A$2:$A$6", "'Data'!$C$2:$C$6");
    bubbleSeries.setBubbleSizeReference("'Data'!$D$2:$D$6");
    bubble.addSeries(std::move(bubbleSeries));
    sheet.addChart(bubble);

    xlpp::Chart combo(xlpp::Chart::Type::Bar);
    combo.setTitle("Combo");
    combo.setXAxisTitle("X");
    combo.setYAxisTitle("Primary");
    combo.addPlot(xlpp::Chart::Type::Bar, xlpp::Chart::Grouping::Clustered, false);
    combo.addSeriesToPlot(0, makeSeries("Primary", "'Data'!$A$2:$A$6", "'Data'!$B$2:$B$6"));
    combo.addPlot(xlpp::Chart::Type::Line, xlpp::Chart::Grouping::Standard, true);
    combo.addSeriesToPlot(1, makeSeries("Secondary", "'Data'!$A$2:$A$6", "'Data'!$C$2:$C$6"));
    sheet.addChart(combo);

    wb.save(path);
    const auto archive = xlpp::internal::ZipArchive::open(path);
    const auto scatterXml = archive.get("xl/charts/chart1.xml");
    check(scatterXml.find("<c:scatterStyle val=\"smoothMarker\"/>") != std::string::npos,
          "Generated Scatter chart preserves scatterStyle");
    check(scatterXml.find("<c:xVal><c:numRef>") != std::string::npos &&
          scatterXml.find("<c:yVal><c:numRef>") != std::string::npos &&
          scatterXml.find("<c:cat>") == std::string::npos,
          "Generated Scatter chart uses xVal/yVal schema");
    check(scatterXml.find("<c:valAx><c:axId val=\"1\"") != std::string::npos,
          "Generated Scatter chart uses a value X axis");

    const auto bubbleXml = archive.get("xl/charts/chart2.xml");
    check(bubbleXml.find("<c:xVal><c:numRef>") != std::string::npos &&
          bubbleXml.find("<c:yVal><c:numRef>") != std::string::npos &&
          bubbleXml.find("<c:bubbleSize><c:numRef>") != std::string::npos,
          "Generated Bubble chart emits X/Y/size references");

    const auto comboXml = archive.get("xl/charts/chart3.xml");
    check(comboXml.find("<c:barChart>") != std::string::npos && comboXml.find("<c:lineChart>") != std::string::npos,
          "Generated combo chart emits independent bar and line plot groups");
    check(comboXml.find("<c:axId val=\"200\"/>") != std::string::npos &&
          comboXml.find("<c:valAx><c:axId val=\"200\"") != std::string::npos &&
          comboXml.find("<c:axPos val=\"r\"/>") != std::string::npos,
          "Generated combo chart emits a secondary value axis");

    xlpp::Workbook loaded;
    loaded.load(path);
    const auto* s = static_cast<const xlpp::Workbook&>(loaded).worksheet("Data");
    check(s && s->chartCount() == 3, "Scatter/Bubble/combo charts reload into the chart model");
    if (s && s->chartCount() == 3) {
        checkEqual(s->chart(0).scatterStyle(), std::string("smoothMarker"), "Scatter style round-trips");
        checkEqual(s->chart(1).series().front().bubbleSizeReference(), std::string("'Data'!$D$2:$D$6"),
                   "Bubble-size reference round-trips");
        check(s->chart(2).combined() && s->chart(2).plots().size() == 2,
              "Generated combo chart reloads as two plots");
        check(s->chart(2).plots().size() == 2 && s->chart(2).plots()[1].usesSecondaryAxes,
              "Secondary-axis plot identity round-trips");
    }
    cleanupOrReport(path);
}

void testTypedPivotCacheRoundTrip() {
    const auto path = tempPath("xlpp_p1s_typed_pivot");
    xlpp::Workbook wb;
    auto& sheet = wb.addWorksheet("Data");
    sheet.append({std::string("Code"), std::string("Amount"), std::string("When"), std::string("Flag"), std::string("Optional")});
    sheet.append({std::string("00123"), 123.0, xlpp::DateTime{2026, 8, 14, 9, 30, 0}, true, xlpp::CellValue{}});

    xlpp::PivotTable pivot("TypedPivot");
    pivot.setLocation("H2");
    pivot.cache().setSourceData("'Data'!$A$1:$E$2");
    pivot.cache().setFields({"Code", "Amount", "When", "Flag", "Optional"});
    pivot.cache().addTypedRecord(
        {"00123", "123", "2026-08-14T09:30:00", "true", ""},
        {xlpp::PivotCacheValueKind::String, xlpp::PivotCacheValueKind::Number,
         xlpp::PivotCacheValueKind::DateTime, xlpp::PivotCacheValueKind::Boolean,
         xlpp::PivotCacheValueKind::Missing});
    pivot.addRowField("Code");
    pivot.addDataField("Amount", "sum");
    sheet.addPivotTable(std::move(pivot));
    wb.save(path);

    const auto archive = xlpp::internal::ZipArchive::open(path);
    const auto definition = archive.get("xl/pivotCache/pivotCacheDefinition1.xml");
    const auto records = archive.get("xl/pivotCache/pivotCacheRecords1.xml");
    check(definition.find("<s v=\"00123\"/>") != std::string::npos,
          "Pivot shared items preserve numeric-looking strings as strings");
    check(records.find("<n v=\"123\"/>") != std::string::npos,
          "Pivot cache records preserve numeric values as numbers");
    check(definition.find("<d v=\"2026-08-14T09:30:00\"/>") != std::string::npos &&
          definition.find("containsDate=\"1\"") != std::string::npos,
          "Pivot cache definition emits typed DateTime shared items");
    check(definition.find("<b v=\"1\"/>") != std::string::npos,
          "Pivot shared items preserve Boolean physical type");
    check(definition.find("<m/>") != std::string::npos,
          "Pivot shared items preserve missing values");

    xlpp::Workbook loaded;
    loaded.load(path);
    const auto* loadedSheet = static_cast<const xlpp::Workbook&>(loaded).worksheet("Data");
    check(loadedSheet && loadedSheet->pivotTables().size() == 1, "Typed PivotTable reloads");
    if (loadedSheet && loadedSheet->pivotTables().size() == 1) {
        const auto& cache = loadedSheet->pivotTables().front().cache();
        check(cache.hasTypedRecordKinds(), "Pivot physical value kinds survive load");
        if (cache.hasTypedRecordKinds()) {
            checkEqual(cache.recordKind(0, 0), xlpp::PivotCacheValueKind::String, "Code remains string");
            checkEqual(cache.recordKind(0, 1), xlpp::PivotCacheValueKind::Number, "Amount remains number");
            checkEqual(cache.recordKind(0, 2), xlpp::PivotCacheValueKind::DateTime, "When remains date/time");
            checkEqual(cache.recordKind(0, 3), xlpp::PivotCacheValueKind::Boolean, "Flag remains Boolean");
            checkEqual(cache.recordKind(0, 4), xlpp::PivotCacheValueKind::Missing, "Optional remains missing");
        }
    }
    cleanupOrReport(path);
}
} // namespace

int main() {
    try {
        testTemplateRoundTrip();
        testScatterBubbleAndComboGeneration();
        testTypedPivotCacheRoundTrip();
    } catch (const std::exception& e) {
        ++failures;
        std::cerr << "UNCAUGHT: " << e.what() << '\n';
    }
    if (failures == 0) {
        std::cout << "P1S three-pillar regression: PASS\n";
        return 0;
    }
    std::cerr << "P1S three-pillar regression: " << failures << " failure(s)\n";
    return 1;
}
