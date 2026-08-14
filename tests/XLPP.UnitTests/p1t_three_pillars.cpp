#include "Packaging/RelationshipGraph.h"
#include "Packaging/ZipArchive.h"
#include <XLPP/XLPP.h>

#include <chrono>
#include <cstdlib>
#include <filesystem>
#include <iostream>
#include <string>
#include <vector>

#ifndef XLPP_TEST_SOURCE_DIR
#define XLPP_TEST_SOURCE_DIR "."
#endif

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

bool keepArtifacts() { return std::getenv("XLPP_KEEP_P1T_ARTIFACTS") != nullptr; }

std::filesystem::path tempPath(const char* stem, const char* ext = ".xlsx") {
    const auto tick = std::chrono::high_resolution_clock::now().time_since_epoch().count();
    return std::filesystem::temp_directory_path() /
        (std::string(stem) + "_" + std::to_string(tick) + ext);
}

void cleanupOrReport(const std::filesystem::path& path) {
    if (keepArtifacts()) std::cout << "P1T_ARTIFACT=" << path.string() << '\n';
    else std::filesystem::remove(path);
}

xlpp::ChartSeries series(std::string title, std::string categories, std::string values) {
    xlpp::ChartSeries s(std::move(title));
    s.setCategoriesReference(std::move(categories));
    s.setValuesReference(std::move(values));
    return s;
}

xlpp::Chart makeComboChart() {
    xlpp::Chart combo(xlpp::Chart::Type::Bar);
    combo.setTitle("P1T Chartsheet Combo");
    combo.setXAxisTitle("Category");
    combo.setYAxisTitle("Primary");
    combo.addPlot(xlpp::Chart::Type::Bar, xlpp::Chart::Grouping::Clustered, false);
    combo.addSeriesToPlot(0, series("Primary", "'Data'!$A$2:$A$5", "'Data'!$B$2:$B$5"));
    combo.addPlot(xlpp::Chart::Type::Line, xlpp::Chart::Grouping::Standard, true);
    combo.addSeriesToPlot(1, series("Secondary", "'Data'!$A$2:$A$5", "'Data'!$C$2:$C$5"));
    return combo;
}

void testFirstClassChartsheetAndWorkbookOrder() {
    const auto path = tempPath("xlpp_p1t_chartsheet");
    const auto roundtrip = tempPath("xlpp_p1t_chartsheet_roundtrip");

    xlpp::Workbook wb;
    auto& data = wb.addWorksheet("Data");
    data.append({std::string("Category"), std::string("Primary"), std::string("Secondary")});
    for (int i = 1; i <= 4; ++i)
        data.append({std::string("C") + std::to_string(i), static_cast<double>(i * 10), static_cast<double>(i * 100)});
    wb.addWorksheet("Notes").cell("A1").setValue("sheet-order-sentinel");
    wb.addChartsheet("Chart View", makeComboChart());
    wb.moveWorkbookSheet(2, 1); // Data, Chart View, Notes

    checkEqual(wb.sheetCount(), std::size_t{2}, "Legacy sheetCount remains worksheet-only");
    checkEqual(wb.chartsheetCount(), std::size_t{1}, "Chartsheet count is first-class");
    checkEqual(wb.workbookSheetCount(), std::size_t{3}, "Mixed workbook sheet count includes chartsheets");
    checkEqual(wb.workbookSheetNames(), std::vector<std::string>({"Data", "Chart View", "Notes"}),
               "Mixed workbook tab order is explicit and movable");
    const auto descriptors = wb.workbookSheets();
    check(descriptors.size() == 3 && descriptors[1].kind == xlpp::WorkbookSheetKind::Chartsheet,
          "Workbook sheet descriptor identifies chart-only sheet kind");

    wb.save(path);
    const auto archive = xlpp::internal::ZipArchive::open(path);
    check(archive.contains("xl/chartsheets/sheet1.xml"), "Generated chartsheet part exists");
    check(archive.contains("xl/chartsheets/_rels/sheet1.xml.rels"), "Generated chartsheet relationship part exists");
    check(archive.get("[Content_Types].xml").find("chartsheet+xml") != std::string::npos,
          "Content types declares chartsheet part");
    const auto workbookXml = archive.get("xl/workbook.xml");
    const auto pData = workbookXml.find("name=\"Data\"");
    const auto pChart = workbookXml.find("name=\"Chart View\"");
    const auto pNotes = workbookXml.find("name=\"Notes\"");
    check(pData != std::string::npos && pChart != std::string::npos && pNotes != std::string::npos
              && pData < pChart && pChart < pNotes,
          "Workbook XML preserves mixed worksheet/chartsheet tab order");
    check(archive.get("xl/_rels/workbook.xml.rels").find(
              "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartsheet") != std::string::npos,
          "Workbook relationship graph uses the chartsheet relationship type");
    const auto graph = xlpp::internal::RelationshipGraph::fromArchive(archive);
    check(graph.validate().ok(), "Generated chartsheet package graph validates");
    check(graph.objectInventory().worksheets == 2 && graph.objectInventory().chartsheets == 1
              && graph.objectInventory().charts == 1 && graph.objectInventory().drawings == 1,
          "Package inventory counts first-class chartsheet ownership");

    xlpp::Workbook loaded;
    loaded.load(path);
    checkEqual(loaded.sheetNames(), std::vector<std::string>({"Data", "Notes"}),
               "Legacy worksheet APIs exclude chartsheets after load");
    checkEqual(loaded.workbookSheetNames(), std::vector<std::string>({"Data", "Chart View", "Notes"}),
               "Mixed tab order survives load");
    const auto* cs = static_cast<const xlpp::Workbook&>(loaded).chartsheet("Chart View");
    check(cs && cs->hasChart(), "Generated chartsheet reloads with a materialized chart");
    if (cs && cs->hasChart()) {
        check(cs->chart().combined() && cs->chart().plots().size() == 2,
              "Chartsheet combo chart reloads as two plot groups");
        check(cs->chart().plots().size() == 2 && cs->chart().plots()[1].usesSecondaryAxes,
              "Chartsheet combo secondary-axis identity survives load");
    }

    xlpp::StreamingWorkbookReader streaming(path);
    checkEqual(streaming.worksheetNames(), std::vector<std::string>({"Data", "Notes"}),
               "Streaming reader continues to expose row-streamable worksheets only");
    checkEqual(streaming.chartsheetNames(), std::vector<std::string>({"Chart View"}),
               "Streaming reader reports chartsheets without treating them as row sources");
    checkEqual(streaming.workbookSheetNames(), std::vector<std::string>({"Data", "Chart View", "Notes"}),
               "Streaming reader preserves mixed workbook tab order");

    // Mutation opts an imported/generated chartsheet into regeneration.
    auto* mutableCs = loaded.chartsheet("Chart View");
    check(mutableCs != nullptr, "Mutable chartsheet lookup succeeds");
    if (mutableCs) mutableCs->chart().setTitle("P1T Mutated Chartsheet");
    loaded.save(roundtrip);
    xlpp::Workbook reloaded;
    reloaded.load(roundtrip);
    const auto* reloadedCs = static_cast<const xlpp::Workbook&>(reloaded).chartsheet("Chart View");
    check(reloadedCs && reloadedCs->hasChart() && reloadedCs->chart().title() == "P1T Mutated Chartsheet",
          "Chartsheet chart mutation regenerates and round-trips");

    cleanupOrReport(path);
    cleanupOrReport(roundtrip);
}

void testOpenpyxlChartsheetImportPreservation() {
    const auto fixture = std::filesystem::path(XLPP_TEST_SOURCE_DIR) / "fixtures" / "openpyxl_chartsheet_fixture.xlsx";
    const auto roundtrip = tempPath("xlpp_p1t_openpyxl_chartsheet_roundtrip");
    check(std::filesystem::exists(fixture), "Openpyxl chartsheet fixture is present");
    if (!std::filesystem::exists(fixture)) return;

    xlpp::Workbook wb;
    wb.load(fixture);
    checkEqual(wb.workbookSheetNames(), std::vector<std::string>({"Data", "Chart View"}),
               "Openpyxl-created chartsheet tab order imports");
    checkEqual(wb.sheetCount(), std::size_t{1}, "Openpyxl fixture has one row worksheet");
    checkEqual(wb.chartsheetCount(), std::size_t{1}, "Openpyxl fixture chartsheet is recognized");
    const auto* cs = static_cast<const xlpp::Workbook&>(wb).chartsheet("Chart View");
    check(cs && cs->imported(), "Imported chartsheet remains preservation-backed");
    check(cs && cs->hasChart() && cs->chart().type() == xlpp::Chart::Type::Bar,
          "Openpyxl-created chartsheet chart is materialized for inspection");

    wb.save(roundtrip);
    const auto sourceZip = xlpp::internal::ZipArchive::open(fixture);
    const auto savedZip = xlpp::internal::ZipArchive::open(roundtrip);
    check(savedZip.contains("xl/chartsheets/sheet1.xml"), "Untouched imported chartsheet part remains present");
    check(savedZip.get("xl/chartsheets/sheet1.xml") == sourceZip.get("xl/chartsheets/sheet1.xml"),
          "Untouched imported chartsheet XML remains byte-preserved");
    cleanupOrReport(roundtrip);
}


void testChartsheetLifecycleCleanup() {
    const auto original = tempPath("xlpp_p1t_chartsheet_lifecycle_original");
    const auto renamed = tempPath("xlpp_p1t_chartsheet_lifecycle_renamed");
    const auto removed = tempPath("xlpp_p1t_chartsheet_lifecycle_removed");

    xlpp::Workbook wb;
    wb.addWorksheet("Data").cell("A1").setValue(1.0);
    xlpp::Chart chart(xlpp::Chart::Type::Line);
    chart.setTitle("Lifecycle");
    chart.addSeries(series("S", "'Data'!$A$1:$A$1", "'Data'!$A$1:$A$1"));
    wb.addChartsheet("Chart View", std::move(chart));
    wb.save(original);

    xlpp::Workbook loaded;
    loaded.load(original);
    check(loaded.renameChartsheet("Chart View", "Dashboard"), "Chartsheet rename succeeds");
    checkEqual(loaded.workbookSheetNames(), std::vector<std::string>({"Data", "Dashboard"}),
               "Chartsheet rename updates mixed workbook tab model");
    loaded.save(renamed);
    auto renamedZip = xlpp::internal::ZipArchive::open(renamed);
    auto renamedGraph = xlpp::internal::RelationshipGraph::fromArchive(renamedZip);
    check(renamedGraph.validate().ok() && renamedGraph.objectInventory().chartsheets == 1,
          "Renamed imported chartsheet retains a valid package graph");

    check(loaded.removeChartsheet("Dashboard"), "Chartsheet removal succeeds while a worksheet remains");
    check(loaded.chartsheetCount() == 0 && loaded.workbookSheetCount() == 1,
          "Chartsheet removal updates workbook sheet topology");
    loaded.save(removed);
    auto removedZip = xlpp::internal::ZipArchive::open(removed);
    auto removedGraph = xlpp::internal::RelationshipGraph::fromArchive(removedZip);
    check(removedGraph.validate().ok(), "Chartsheet removal leaves no broken relationship graph");
    check(removedGraph.objectInventory().chartsheets == 0 && removedGraph.objectInventory().charts == 0
              && removedGraph.objectInventory().drawings == 0,
          "Chartsheet removal retires exclusive chart/drawing package closure");

    cleanupOrReport(original);
    cleanupOrReport(renamed);
    cleanupOrReport(removed);
}

void testChartsheetResourceLimit() {
    const auto path = tempPath("xlpp_p1t_chartsheet_limit");
    xlpp::Workbook wb;
    wb.addWorksheet("Data");
    wb.addChartsheet("Chart A", xlpp::Chart(xlpp::Chart::Type::Bar));
    wb.addChartsheet("Chart B", xlpp::Chart(xlpp::Chart::Type::Line));
    wb.save(path);

    xlpp::Workbook limited;
    xlpp::LoadOptions options;
    options.maxChartsheets = 1;
    bool rejected = false;
    try { limited.load(path, options); }
    catch (const std::exception&) { rejected = true; }
    check(rejected, "Materialized load enforces maxChartsheets model limit");
    cleanupOrReport(path);
}

void testPivotLogicalItemIdentityRepair() {
    const auto first = tempPath("xlpp_p1t_pivot_identity_before");
    const auto second = tempPath("xlpp_p1t_pivot_identity_after");

    xlpp::Workbook wb;
    auto& sheet = wb.addWorksheet("Data");
    sheet.append({std::string("Category"), std::string("Amount")});
    sheet.append({std::string("A"), 10.0});
    sheet.append({std::string("B"), 20.0});
    sheet.append({std::string("C"), 30.0});

    xlpp::PivotTable pivot("StableItems");
    pivot.setLocation("E2");
    pivot.cache().setSourceData("'Data'!$A$1:$B$4");
    pivot.cache().setFields({"Category", "Amount"});
    pivot.cache().addTypedRecord({"A", "10"}, {xlpp::PivotCacheValueKind::String, xlpp::PivotCacheValueKind::Number});
    pivot.cache().addTypedRecord({"B", "20"}, {xlpp::PivotCacheValueKind::String, xlpp::PivotCacheValueKind::Number});
    pivot.cache().addTypedRecord({"C", "30"}, {xlpp::PivotCacheValueKind::String, xlpp::PivotCacheValueKind::Number});
    auto& rowField = pivot.addRowField("Category");
    rowField.hideCacheValue("B", xlpp::PivotCacheValueKind::String);
    pivot.addDataField("Amount", "sum");
    sheet.addPivotTable(std::move(pivot));
    wb.save(first);

    // Reorder first-seen shared-item identity from [A,B,C] to [B,C,A]. The
    // hidden logical value B must move from physical x=1 to x=0 on save.
    auto& table = sheet.pivotTables().front();
    table.cache().setRecordValue(0, 0, "B", xlpp::PivotCacheValueKind::String);
    table.cache().setRecordValue(1, 0, "C", xlpp::PivotCacheValueKind::String);
    table.cache().setRecordValue(2, 0, "A", xlpp::PivotCacheValueKind::String);
    wb.save(second);

    const auto archive = xlpp::internal::ZipArchive::open(second);
    const auto definition = archive.get("xl/pivotCache/pivotCacheDefinition1.xml");
    const auto tableXml = archive.get("xl/pivotTables/pivotTable1.xml");
    const auto bPos = definition.find("<s v=\"B\"/>");
    const auto cPos = definition.find("<s v=\"C\"/>");
    const auto aPos = definition.find("<s v=\"A\"/>");
    check(bPos != std::string::npos && cPos != std::string::npos && aPos != std::string::npos
              && bPos < cPos && cPos < aPos,
          "Pivot shared items compact to the new first-seen physical order");
    check(tableXml.find("<item x=\"0\" h=\"1\"") != std::string::npos,
          "Logical hidden Pivot item is rebound to its new physical shared-item index");

    xlpp::Workbook loaded;
    loaded.load(second);
    const auto* loadedSheet = static_cast<const xlpp::Workbook&>(loaded).worksheet("Data");
    check(loadedSheet && loadedSheet->pivotTables().size() == 1, "Repaired PivotTable reloads");
    if (loadedSheet && loadedSheet->pivotTables().size() == 1) {
        const auto& fields = loadedSheet->pivotTables().front().rowFields();
        check(fields.size() == 1 && !fields.front().items().empty(), "Loaded Pivot row field retains item metadata");
        if (fields.size() == 1 && !fields.front().items().empty()) {
            const auto& item = fields.front().items().front();
            check(item.hidden && item.cacheIndex == 0, "Loaded hidden Pivot item uses repaired physical index");
            check(item.hasCacheValue && item.cacheValue == "B" && item.cacheValueKind == xlpp::PivotCacheValueKind::String,
                  "Loaded Pivot item binds to stable logical typed value identity");
        }
    }

    cleanupOrReport(first);
    cleanupOrReport(second);
}

} // namespace

int main() {
    try {
        testFirstClassChartsheetAndWorkbookOrder();
        testOpenpyxlChartsheetImportPreservation();
        testChartsheetLifecycleCleanup();
        testChartsheetResourceLimit();
        testPivotLogicalItemIdentityRepair();
    } catch (const std::exception& e) {
        ++failures;
        std::cerr << "UNCAUGHT: " << e.what() << '\n';
    }
    if (failures == 0) {
        std::cout << "P1T three-pillar regression: PASS\n";
        return 0;
    }
    std::cerr << "P1T three-pillar regression: " << failures << " failure(s)\n";
    return 1;
}
