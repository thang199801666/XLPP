#include "RegressionTests.h"
#include <XLPP/XLPP.h>
#include <filesystem>

void testFormulaDirtyRecalculation(TestContext& test) {
    xlpp::Workbook wb;
    auto& ws = wb.addWorksheet("Calc");
    ws.cell("A1").setValue(1.0);
    ws.cell("B1").setFormula("A1*2");
    ws.cell("C1").setFormula("B1+1");
    ws.cell("D1").setFormula("40+2");
    auto initial = wb.calculateFormulas();
    test.checkTrue(initial.success(), "Initial formula calculation succeeds");
    test.checkEqual(ws.cell("B1").numericValueOr(-1.0), 2.0, "Initial dependent B1");
    test.checkEqual(ws.cell("C1").numericValueOr(-1.0), 3.0, "Initial transitive dependent C1");
    test.checkEqual(ws.cell("D1").numericValueOr(-1.0), 42.0, "Independent formula is initially calculated");

    ws.cell("A1").setValue(5.0);
    xlpp::CalculationOptions options;
    options.changedCells.push_back({"Calc", "A1"});
    const auto report = wb.calculateFormulas(options);
    test.checkTrue(report.success(), "Dirty-root calculation succeeds");
    test.checkEqual(report.dirtyRoots, std::size_t{1}, "Dirty-root count");
    test.checkEqual(report.dirtyFormulaCellsSelected, std::size_t{2}, "Dependency fan-out selects B1 and C1 only");
    test.checkEqual(ws.cell("B1").numericValueOr(-1.0), 10.0, "Dirty recalculation updates direct dependent");
    test.checkEqual(ws.cell("C1").numericValueOr(-1.0), 11.0, "Dirty recalculation updates transitive dependent");
    test.checkEqual(ws.cell("D1").numericValueOr(-1.0), 42.0, "Dirty recalculation leaves unrelated formula cache intact");

    wb.calcProperties().setCalculationMode(xlpp::CalculationMode::AutomaticExceptDataTables);
    test.checkEqual(wb.calcProperties().calcMode(), std::string("autoNoTable"), "Typed calculation mode maps to OOXML calcMode");
}

void testAdvancedAutoFilterRoundTrip(TestContext& test) {
    const auto path = std::filesystem::temp_directory_path() / "xlpp_advanced_autofilter_phase29.xlsx";
    xlpp::Workbook wb;
    auto& ws = wb.addWorksheet("Filters");
    ws.autoFilter().setReference("A1:E20");

    xlpp::Top10Filter top; top.top = false; top.percent = true; top.value = 25; top.filterValue = 7.5;
    ws.autoFilter().column(0).setTop10Filter(top);
    xlpp::DynamicFilter dyn; dyn.type = xlpp::DynamicFilterType::ThisMonth; dyn.value = 45000.0; dyn.maxValue = 45031.0;
    ws.autoFilter().column(1).setDynamicFilter(dyn);
    xlpp::ColorFilter color; color.dxfId = 3; color.cellColor = false;
    ws.autoFilter().column(2).setColorFilter(color);
    xlpp::IconFilter icon; icon.iconSet = "4Arrows"; icon.iconId = 2;
    ws.autoFilter().column(3).setIconFilter(icon);
    xlpp::DateGroupItem group; group.year = 2026; group.month = 8; group.grouping = xlpp::DateTimeGrouping::Month;
    ws.autoFilter().column(4).addDateGroup(group);
    wb.save(path);

    xlpp::Workbook loaded; loaded.load(path);
    const auto& filter = loaded[0].autoFilter();
    test.checkTrue(filter.tryColumn(0)->top10Filter().has_value(), "Top10 filter reloads");
    test.checkTrue(filter.tryColumn(0)->top10Filter()->percent, "Top10 percent semantics reload");
    test.checkEqual(filter.tryColumn(0)->top10Filter()->value, 25.0, "Top10 value reload");
    test.checkTrue(filter.tryColumn(1)->dynamicFilter().has_value(), "Dynamic filter reloads");
    test.checkTrue(filter.tryColumn(1)->dynamicFilter()->type == xlpp::DynamicFilterType::ThisMonth, "Dynamic type reloads");
    test.checkTrue(filter.tryColumn(2)->colorFilter().has_value(), "Color filter reloads");
    test.checkEqual(filter.tryColumn(2)->colorFilter()->dxfId, std::size_t{3}, "Color dxfId reloads");
    test.checkTrue(filter.tryColumn(3)->iconFilter().has_value(), "Icon filter reloads");
    test.checkEqual(filter.tryColumn(3)->iconFilter()->iconSet, std::string("4Arrows"), "Icon set reloads");
    test.checkEqual(filter.tryColumn(4)->dateGroups().size(), std::size_t{1}, "Date group reloads");
    test.checkEqual(filter.tryColumn(4)->dateGroups()[0].year, 2026, "Date group year reloads");
    std::filesystem::remove(path);
}

void testExternalDataAndDataModelInspection(TestContext& test) {
    xlpp::Workbook wb;
    wb.addWorksheet("Data");
    wb.preservedParts().push_back({"xl/connections.xml",
        "<connections><connection id=\"7\" name=\"Warehouse\" type=\"5\" refreshOnLoad=\"1\" background=\"1\"/></connections>",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.connections+xml", "xml", "application/xml", true});
    wb.preservedParts().push_back({"xl/queryTables/queryTable1.xml",
        "<queryTable name=\"Orders\" connectionId=\"7\" refreshOnLoad=\"1\"><webPr/></queryTable>",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.queryTable+xml", "xml", "application/xml", true});
    wb.preservedParts().push_back({"xl/externalLinks/externalLink1.xml",
        "<externalLink><externalBook><sheetNames><sheetName val=\"Input\"/></sheetNames><definedNames><definedName name=\"Rate\"/></definedNames></externalBook></externalLink>",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml", "xml", "application/xml", true});
    wb.preservedParts().push_back({"xl/model/model.bin", std::string("MODEL", 5),
        "application/vnd.ms-excel.model", "bin", "application/octet-stream", true});
    wb.preservedParts().push_back({"xl/pivotCache/pivotCacheDefinition9.xml",
        "<pivotCacheDefinition><cacheSource type=\"external\"/><olapPr/></pivotCacheDefinition>",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml", "xml", "application/xml", true});

    const auto external = wb.inspectExternalData();
    test.checkEqual(external.connections.size(), std::size_t{1}, "Connections are inspectable without regeneration");
    test.checkEqual(external.connections[0].name, std::string("Warehouse"), "Connection metadata parsed");
    test.checkEqual(external.queryTables.size(), std::size_t{1}, "Query table metadata parsed");
    test.checkEqual(external.externalWorkbooks.size(), std::size_t{1}, "External workbook link metadata parsed");
    test.checkEqual(external.externalWorkbooks[0].sheetNames[0], std::string("Input"), "External sheet names parsed");
    const auto model = wb.inspectDataModel();
    test.checkTrue(model.present, "Data Model part is detected");
    test.checkTrue(model.hasOlapPivotCaches, "OLAP pivot cache is detected as preservation-only");
    test.checkEqual(model.modelParts.size(), std::size_t{1}, "Data Model part inventory");
}
