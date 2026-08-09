#include <XLPP/XLPP.h>
#include "Package/Zip/ZipArchive.h"
#include "Package/Zip/ZipArchiveReader.h"
#include "Package/Opc/RelationshipGraph.h"
#include "Platform/MappedFile.h"
#include "Streaming/SharedStringsReader.h"
#include "Core/Threading/ThreadPool.h"
#include "Package/Xml/SimdScan.h"
#include "Package/Xml/XmlScanner.h"
#include "Package/Xml/XmlUtilities.h"
#include "VBA/VbaProjectBinary.h"
#include "Encryption/OfficeEncryption.h"
#include <array>
#include <chrono>
#include <cmath>
#include <filesystem>
#include <fstream>
#include <functional>
#include <iostream>
#include <set>
#include <sstream>
#include <stdexcept>
#include <string>
#include <variant>
#include <vector>

#include "../TestFramework.h"
#include "RegressionTests.h"

void testChartPartPreservation(TestContext& test) {
    const auto dir = std::filesystem::temp_directory_path();
    const auto base = dir / "xlpp_chart_preserve_base.xlsx";
    const auto staged = dir / "xlpp_chart_preserve_staged.xlsx";
    const auto out = dir / "xlpp_chart_preserve_out.xlsx";
    {
        xlpp::Workbook w;
        w.addWorksheet("Sheet1").cell("A1").setValue("chart host");
        w.save(base);
    }
    {
        auto z = xlpp::internal::ZipArchive::open(base);
        const std::string chartXml = "<c:chartSpace xmlns:c=\"urn:fixture\"><c:extLst><c:ext uri=\"custom\"/></c:extLst></c:chartSpace>";
        z.add("xl/charts/chart1.xml", chartXml);
        z.add("xl/drawings/drawing1.xml", "<drawing/>" );
        z.add("xl/drawings/_rels/drawing1.xml.rels", "<Relationships><Relationship Id=\"rIdChart1\" Target=\"../charts/chart1.xml\"/></Relationships>");
        z.add("xl/worksheets/_rels/sheet1.xml.rels", "<Relationships><Relationship Id=\"rIdDrawing\" Target=\"../drawings/drawing1.xml\"/></Relationships>");
        auto ct = z.get("[Content_Types].xml");
        const auto marker = std::string("<Override PartName=\"/xl/charts/chart1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.drawingml.chart+xml\"/>");
        ct.insert(ct.rfind("</Types>"), marker);
         z.replace("[Content_Types].xml", ct);
        z.save(staged);
    }
    {
        xlpp::Workbook w;
        w.load(staged);
        const auto it = std::find_if(w.preservedParts().begin(), w.preservedParts().end(),
            [](const xlpp::PreservedPart& part) { return part.name == "xl/charts/chart1.xml"; });
        test.checkTrue(it != w.preservedParts().end(), "Chart part is preserved on load");
        w.save(out);
    }
    {
        auto z = xlpp::internal::ZipArchive::open(out);
        test.checkEqual(z.get("xl/charts/chart1.xml"), std::string("<c:chartSpace xmlns:c=\"urn:fixture\"><c:extLst><c:ext uri=\"custom\"/></c:extLst></c:chartSpace>"), "Chart XML survives load-save");
        test.checkEqual(z.get("xl/drawings/_rels/drawing1.xml.rels"), std::string("<Relationships><Relationship Id=\"rIdChart1\" Target=\"../charts/chart1.xml\"/></Relationships>"), "Chart relationships survive load-save");
    }
    std::filesystem::remove(base);
    std::filesystem::remove(staged);
    std::filesystem::remove(out);
}

void testIndependentImageChartFixtureRoundTrip(TestContext& test) {
    const auto fixtureRoot = std::filesystem::path(XLPP_TEST_SOURCE_DIR) / "fixtures";
    const std::array<std::pair<const char*, std::filesystem::path>, 2> fixtures{{
        {"OpenPyXL", fixtureRoot / "openpyxl" / "image_chart.xlsx"},
        {"LibreOffice", fixtureRoot / "libreoffice" / "image_chart.xlsx"}
    }};

    for (const auto& [producer, sourcePath] : fixtures) {
        test.checkTrue(std::filesystem::exists(sourcePath), std::string(producer) + " fixture exists");
        const auto before = xlpp::internal::ZipArchive::open(sourcePath);
        const auto beforeGraph = xlpp::internal::RelationshipGraph::fromArchive(before);
        const auto beforeValidation = beforeGraph.validate();
        test.checkTrue(beforeValidation.ok(), std::string(producer) + " fixture object graph is valid");
        test.checkEqual(beforeGraph.objectInventory().worksheets, std::size_t{1}, std::string(producer) + " worksheet count");
        test.checkEqual(beforeGraph.objectInventory().drawings, std::size_t{1}, std::string(producer) + " drawing count");
        test.checkEqual(beforeGraph.objectInventory().images, std::size_t{1}, std::string(producer) + " visible image count");
        test.checkEqual(beforeGraph.objectInventory().charts, std::size_t{1}, std::string(producer) + " visible chart count");

        const auto outputPath = std::filesystem::temp_directory_path() /
            (std::string("xlpp_external_") + producer + "_roundtrip.xlsx");
        xlpp::Workbook workbook;
        workbook.load(sourcePath);
        const auto hasBrokenReferenceWarning = std::any_of(
            workbook.diagnostics().warnings.begin(), workbook.diagnostics().warnings.end(), [](const auto& warning) {
                return warning.find("Broken owner reference") != std::string::npos;
            });
        test.checkTrue(!hasBrokenReferenceWarning, std::string(producer) + " load has no owner-reference warning");
        auto* sheet = workbook.worksheet("Objects");
        test.checkTrue(sheet != nullptr, std::string(producer) + " worksheet loads by name");
        sheet->cell("K21").setValue(std::string("unrelated edit"));
        workbook.save(outputPath);

        const auto after = xlpp::internal::ZipArchive::open(outputPath);
        const auto diff = xlpp::internal::comparePackages(before, after);
        test.checkTrue(diff.afterValidation.ok(), std::string(producer) + " round-trip object graph is valid");
        test.checkTrue(diff.objectCountRegressions.empty(), std::string(producer) + " round-trip has no object-count regression");
        test.checkEqual(diff.afterObjects.drawings, diff.beforeObjects.drawings, std::string(producer) + " drawing count preserved");
        test.checkEqual(diff.afterObjects.images, diff.beforeObjects.images, std::string(producer) + " image count preserved");
        test.checkEqual(diff.afterObjects.charts, diff.beforeObjects.charts, std::string(producer) + " chart count preserved");
        for (const auto& part : {"xl/drawings/drawing1.xml", "xl/drawings/_rels/drawing1.xml.rels",
                                 "xl/charts/chart1.xml", "xl/media/image1.png"}) {
            test.checkTrue(after.contains(part), std::string(producer) + " keeps part " + part);
            test.checkEqual(after.get(part), before.get(part), std::string(producer) + " keeps untouched bytes for " + part);
        }
        std::filesystem::remove(outputPath);
    }
}

void testChartAndPivotPackage(TestContext& test) {
    const auto path = std::filesystem::temp_directory_path() / "xlpp_m21_chart_pivot.xlsx";
    xlpp::Workbook wb;
    auto& sheet = wb.addWorksheet("Charts");
    sheet.append({std::string("Quarter"), std::string("Units")});
    sheet.append({std::string("Q1"), 10.0});
    sheet.append({std::string("Q2"), 20.0});

    xlpp::Chart chart(xlpp::Chart::Type::Bar);
    chart.setTitle("Sales");
    chart.setXAxisTitle("Quarter");
    chart.setYAxisTitle("Units");
    chart.setShowLegend(true);
    chart.setLegendPosition("b");
    chart.setStyle("10");
    auto& series = chart.addSeries(xlpp::ChartSeries("Units"));
    series.reference("Charts", "$B$2:$B$3");
    series.categories("Charts", "$A$2:$A$3");
    sheet.addChart(chart);

    xlpp::PivotTable pivot("SalesPivot");
    pivot.setLocation("D1");
    pivot.cache().setCacheId(1);
    pivot.cache().setSourceData("'Charts'!$A$1:$B$3");
    pivot.cache().setFields({"Quarter", "Units"});
    pivot.cache().addRecord({"Q1", "10"});
    pivot.cache().addRecord({"Q2", "20"});
    pivot.addRowField("Quarter");
    pivot.rowFields().back().setFieldIndex(0);
    pivot.addColumnField("Units");
    pivot.columnFields().back().setFieldIndex(1);
    pivot.addDataField(1);
    sheet.addPivotTable(std::move(pivot));

    wb.save(path);
    test.checkTrue(std::filesystem::exists(path), "Chart/pivot workbook saved");

    auto z = xlpp::internal::ZipArchive::open(path);
    test.checkTrue(z.contains("xl/charts/chart1.xml"), "Chart part written");
    const auto chartXml = z.get("xl/charts/chart1.xml");
    test.checkTrue(chartXml.find("barChart") != std::string::npos, "Bar chart type in part");
    test.checkTrue(chartXml.find("Sales") != std::string::npos, "Chart title in part");
    test.checkTrue(chartXml.find("<c:overlay val=\"0\"/>") != std::string::npos, "Chart title overlay in part");
    test.checkTrue(chartXml.find("<c:v>Units</c:v>") != std::string::npos, "Series title in part");
    test.checkTrue(chartXml.find("<c:style val=\"10\"/>") != std::string::npos, "Chart style in part");
    test.checkTrue(chartXml.find("<c:legendPos val=\"b\"/>") != std::string::npos, "Legend position in part");
    test.checkTrue(z.contains("xl/pivotTables/pivotTable1.xml"), "Pivot part written");
    test.checkTrue(z.contains("xl/pivotCache/pivotCacheDefinition1.xml"), "Pivot cache written");
    test.checkTrue(z.contains("xl/pivotCache/_rels/pivotCacheDefinition1.xml.rels"), "Pivot cache relationship part written");
    test.checkTrue(z.contains("xl/pivotCache/pivotCacheRecords1.xml"), "Pivot cache records part written");
    const auto pivotCacheRels = z.get("xl/pivotCache/_rels/pivotCacheDefinition1.xml.rels");
    test.checkTrue(pivotCacheRels.find("pivotCacheRecords") != std::string::npos, "Pivot cache records relationship written");
    const auto pivotCacheXml = z.get("xl/pivotCache/pivotCacheDefinition1.xml");
    test.checkTrue(pivotCacheXml.find("recordCount=\"2\"") != std::string::npos, "Pivot cache record count written");
    test.checkTrue(pivotCacheXml.find("<cacheFields count=\"2\">") != std::string::npos, "Pivot cache fields written");
    test.checkTrue(pivotCacheXml.find("name=\"Quarter\"") != std::string::npos, "Pivot first field written");
    test.checkTrue(pivotCacheXml.find("<sharedItems count=\"2\">") != std::string::npos, "Pivot shared items written");
    const auto pivotRecordsXml = z.get("xl/pivotCache/pivotCacheRecords1.xml");
    test.checkTrue(pivotRecordsXml.find("<pivotCacheRecords") != std::string::npos &&
                   pivotRecordsXml.find("count=\"2\"") != std::string::npos, "Pivot cache records written");
    test.checkTrue(pivotCacheXml.find("<s v=\"Q1\"/>") != std::string::npos, "Pivot shared string written");
    test.checkTrue(pivotCacheXml.find("<n v=\"10\"/>") != std::string::npos, "Pivot shared number written");
    test.checkTrue(pivotRecordsXml.find("<x v=\"0\"/><x v=\"0\"/>") != std::string::npos, "Pivot cache indexes written");
    const auto pivotTableXml = z.get("xl/pivotTables/pivotTable1.xml");
    test.checkTrue(pivotTableXml.find("<rowFields count=\"1\"><field x=\"0\"/>") != std::string::npos, "Pivot row field index written");
    test.checkTrue(pivotTableXml.find("<colFields count=\"1\"><field x=\"1\"/>") != std::string::npos, "Pivot column field index written");
    test.checkTrue(pivotTableXml.find("<dataField name=\"Sum of Units\" fld=\"1\"") != std::string::npos, "Pivot data field index written");
    test.checkTrue(z.contains("xl/drawings/drawing1.xml"), "Drawing part written for chart");

    {
        xlpp::Workbook loaded;
        loaded.load(path);
        test.checkTrue(loaded.worksheet("Charts") != nullptr, "Chart workbook reloads");
        const auto preservedPivot = std::find_if(loaded.preservedParts().begin(), loaded.preservedParts().end(),
            [](const xlpp::PreservedPart& part) { return part.name == "xl/pivotTables/pivotTable1.xml"; });
        test.checkTrue(preservedPivot != loaded.preservedParts().end(), "Pivot table part is preserved on load");
        const auto preservedCache = std::find_if(loaded.preservedParts().begin(), loaded.preservedParts().end(),
            [](const xlpp::PreservedPart& part) { return part.name == "xl/pivotCache/pivotCacheDefinition1.xml"; });
        test.checkTrue(preservedCache != loaded.preservedParts().end(), "Pivot cache part is preserved on load");
        const auto preservedOut = std::filesystem::temp_directory_path() / "xlpp_m21_chart_pivot_preserved.xlsx";
        loaded.save(preservedOut);
        auto preservedZip = xlpp::internal::ZipArchive::open(preservedOut);
        test.checkTrue(preservedZip.get("xl/pivotTables/pivotTable1.xml") == z.get("xl/pivotTables/pivotTable1.xml"), "Pivot table XML survives load-save");
        test.checkTrue(preservedZip.get("xl/pivotCache/pivotCacheDefinition1.xml") == z.get("xl/pivotCache/pivotCacheDefinition1.xml"), "Pivot cache XML survives load-save");
        std::filesystem::remove(preservedOut);
    }
    std::filesystem::remove(path);
}

void testImportedChartInspectionAndSelectiveMutation(TestContext& test) {
    const struct FixtureCase {
        const char* producer;
        const char* relativePath;
        xlpp::DrawingAnchorType anchorType;
        const char* preservedMarker;
    } cases[] = {
        {"OpenPyXL", "fixtures/openpyxl/image_chart.xlsx", xlpp::DrawingAnchorType::OneCell, "prstDash"},
        {"LibreOffice", "fixtures/libreoffice/image_chart.xlsx", xlpp::DrawingAnchorType::TwoCell, "c15:showLeaderLines"}
    };

    for (const auto& fixture : cases) {
        const auto source = std::filesystem::path(XLPP_TEST_SOURCE_DIR) / fixture.relativePath;
        const auto before = xlpp::internal::ZipArchive::open(source);
        const auto originalChartXml = before.get("xl/charts/chart1.xml");
        const auto originalDrawingRels = before.get("xl/drawings/_rels/drawing1.xml.rels");
        const auto originalImageBytes = before.get("xl/media/image1.png");
        const auto originalSheetFormats = xlpp::internal::tags(before.get("xl/worksheets/sheet1.xml"), "sheetFormatPr");

        xlpp::Workbook workbook;
        workbook.load(source);
        auto* sheet = workbook.worksheet("Objects");
        test.checkTrue(sheet != nullptr, std::string(fixture.producer) + " chart fixture worksheet loads");
        const auto& charts = static_cast<const xlpp::Worksheet&>(*sheet).charts();
        test.checkEqual(charts.size(), std::size_t{1}, std::string(fixture.producer) + " chart reader exposes one chart");
        if (charts.empty()) continue;
        const auto& chart = charts.front();
        test.checkTrue(chart.imported(), std::string(fixture.producer) + " chart is marked imported");
        test.checkTrue(!chart.stableId().empty(), std::string(fixture.producer) + " chart has stable drawing-object ID");
        test.checkEqual(chart.sourceDrawingPart(), std::string("xl/drawings/drawing1.xml"), std::string(fixture.producer) + " chart source drawing part");
        test.checkEqual(chart.sourceChartPart(), std::string("xl/charts/chart1.xml"), std::string(fixture.producer) + " chart source chart part");
        test.checkEqual(chart.sourceRelationshipId(), std::string("rId1"), std::string(fixture.producer) + " chart source relationship ID");
        test.checkEqual(chart.title(), std::string("OpenPyXL chart"), std::string(fixture.producer) + " chart title parsed across namespace styles");
        test.checkEqual(chart.xAxisTitle(), std::string("Category"), std::string(fixture.producer) + " category-axis title parsed");
        test.checkEqual(chart.yAxisTitle(), std::string("Amount"), std::string(fixture.producer) + " value-axis title parsed");
        test.checkEqual(static_cast<int>(chart.type()), static_cast<int>(xlpp::Chart::Type::Bar), std::string(fixture.producer) + " chart type parsed");
        test.checkEqual(static_cast<int>(chart.anchorInfo().type), static_cast<int>(fixture.anchorType), std::string(fixture.producer) + " chart anchor type parsed");
        test.checkEqual(chart.anchorInfo().from.row, std::size_t{8}, std::string(fixture.producer) + " chart anchor row parsed");
        test.checkEqual(chart.anchorInfo().from.column, std::size_t{4}, std::string(fixture.producer) + " chart anchor column parsed");
        test.checkNear(static_cast<double>(chart.width()), 302.0, 1.0, std::string(fixture.producer) + " chart width parsed from owning anchor");
        test.checkNear(static_cast<double>(chart.height()), 189.0, 1.0, std::string(fixture.producer) + " chart height parsed from owning anchor");
        test.checkEqual(chart.series().size(), std::size_t{1}, std::string(fixture.producer) + " chart series parsed");
        test.checkTrue(!chart.series().front().categoriesReference().empty(), std::string(fixture.producer) + " category formula parsed");
        test.checkTrue(!chart.series().front().valuesReference().empty(), std::string(fixture.producer) + " value formula parsed");

        const auto stableId = chart.stableId();
        test.checkTrue(sheet->chartByStableId(stableId) != nullptr, std::string(fixture.producer) + " chart lookup by stable ID");
        test.checkTrue(sheet->chartByStableId("missing-chart") == nullptr, std::string(fixture.producer) + " missing stable chart ID returns null");
        test.checkTrue(!sheet->setChartSeriesReferences(stableId, 9, "Objects!$A$2:$A$3", "Objects!$B$2:$B$3"), std::string(fixture.producer) + " invalid chart series index is rejected");
        test.checkTrue(!sheet->resizeChart(stableId, 0.0, 200.0), std::string(fixture.producer) + " invalid chart resize is rejected");
        test.checkTrue(!sheet->moveChartAbsolute(stableId, 1, 1), std::string(fixture.producer) + " cell-anchored chart rejects absolute move");

        test.checkTrue(sheet->setChartTitle(stableId, "XL++ <safe> & chart"), std::string(fixture.producer) + " imported chart title edits selectively");
        test.checkTrue(sheet->setChartSeriesReferences(stableId, 0, "'Objects'!$A$2:$A$3", "'Objects'!$B$2:$B$3"), std::string(fixture.producer) + " imported chart series formulas edit selectively");
        test.checkTrue(sheet->moveChart(stableId, "H6"), std::string(fixture.producer) + " imported chart moves by stable ID");
        test.checkTrue(sheet->resizeChart(stableId, 320.0, 200.0), std::string(fixture.producer) + " imported chart resizes by stable ID");
        sheet->cell("J20").setValue("chart-edit-regression");

        const auto output = std::filesystem::temp_directory_path() /
            (std::string("xlpp_p0h_") + fixture.producer + "_chart_mutation.xlsx");
        workbook.save(output);
        const auto after = xlpp::internal::ZipArchive::open(output);
        const auto editedChartXml = after.get("xl/charts/chart1.xml");
        test.checkTrue(editedChartXml.find("XL++ &lt;safe&gt; &amp; chart") != std::string::npos,
                       std::string(fixture.producer) + " chart title is XML-escaped in selective patch");
        test.checkTrue(editedChartXml.find("&apos;Objects&apos;!$A$2:$A$3") != std::string::npos ||
                       editedChartXml.find("'Objects'!$A$2:$A$3") != std::string::npos,
                       std::string(fixture.producer) + " category formula updated");
        test.checkTrue(editedChartXml.find("&apos;Objects&apos;!$B$2:$B$3") != std::string::npos ||
                       editedChartXml.find("'Objects'!$B$2:$B$3") != std::string::npos,
                       std::string(fixture.producer) + " value formula updated");
        test.checkTrue(editedChartXml.find(fixture.preservedMarker) != std::string::npos,
                       std::string(fixture.producer) + " unsupported chart formatting/extensions survive selective edits");
        test.checkEqual(after.get("xl/drawings/_rels/drawing1.xml.rels"), originalDrawingRels,
                        std::string(fixture.producer) + " drawing relationships remain byte-identical");
        test.checkEqual(after.get("xl/media/image1.png"), originalImageBytes,
                        std::string(fixture.producer) + " sibling image bytes remain byte-identical");
        const auto afterSheetFormats = xlpp::internal::tags(after.get("xl/worksheets/sheet1.xml"), "sheetFormatPr");
        test.checkTrue(!originalSheetFormats.empty() && !afterSheetFormats.empty(),
                       std::string(fixture.producer) + " source sheet-format metrics remain present");
        if (!originalSheetFormats.empty() && !afterSheetFormats.empty())
            test.checkEqual(afterSheetFormats.front(), originalSheetFormats.front(),
                            std::string(fixture.producer) + " source sheet-format metrics remain byte-identical for drawing geometry");

        const auto graph = xlpp::internal::RelationshipGraph::fromArchive(after);
        const auto validation = graph.validate();
        test.checkTrue(validation.relationshipSyntaxErrors.empty(), std::string(fixture.producer) + " selective chart output has no relationship syntax errors");
        test.checkTrue(validation.danglingRelationships.empty(), std::string(fixture.producer) + " selective chart output has no dangling relationships");
        test.checkTrue(validation.orphanedParts.empty(), std::string(fixture.producer) + " selective chart output has no orphaned parts");
        test.checkTrue(validation.ownerReferenceErrors.empty(), std::string(fixture.producer) + " selective chart output has no owner-reference errors");
        test.checkEqual(graph.objectInventory().charts, std::size_t{1}, std::string(fixture.producer) + " selective chart output retains visible chart count");
        test.checkEqual(graph.objectInventory().images, std::size_t{1}, std::string(fixture.producer) + " selective chart output retains sibling image count");

        xlpp::Workbook reloaded;
        reloaded.load(output);
        const auto* reloadedSheet = reloaded.worksheet("Objects");
        const auto& reloadedChart = static_cast<const xlpp::Worksheet&>(*reloadedSheet).charts().front();
        test.checkEqual(reloadedChart.title(), std::string("XL++ <safe> & chart"), std::string(fixture.producer) + " selective title survives reload");
        test.checkEqual(reloadedChart.series().front().categoriesReference(), std::string("'Objects'!$A$2:$A$3"), std::string(fixture.producer) + " selective category formula survives reload");
        test.checkEqual(reloadedChart.series().front().valuesReference(), std::string("'Objects'!$B$2:$B$3"), std::string(fixture.producer) + " selective value formula survives reload");
        test.checkEqual(reloadedChart.anchorInfo().from.row, std::size_t{6}, std::string(fixture.producer) + " selective chart row survives reload");
        test.checkEqual(reloadedChart.anchorInfo().from.column, std::size_t{8}, std::string(fixture.producer) + " selective chart column survives reload");
        test.checkNear(static_cast<double>(reloadedChart.width()), 320.0, 1.0, std::string(fixture.producer) + " selective chart width survives reload");
        test.checkNear(static_cast<double>(reloadedChart.height()), 200.0, 1.0, std::string(fixture.producer) + " selective chart height survives reload");
        std::filesystem::remove(output);
        (void)originalChartXml;
    }
}

void testImportedScatterChartDeepSelectiveEditing(TestContext& test) {
    const auto source = std::filesystem::path(XLPP_TEST_SOURCE_DIR) / "fixtures/openpyxl/scatter_multiseries.xlsx";
    xlpp::Workbook workbook;
    workbook.load(source);
    auto* sheet = workbook.worksheet("Scatter");
    test.checkTrue(sheet != nullptr, "OpenPyXL scatter fixture worksheet loads");
    const auto& charts = static_cast<const xlpp::Worksheet&>(*sheet).charts();
    test.checkEqual(charts.size(), std::size_t{1}, "Scatter fixture exposes one chart");
    if (charts.empty()) return;
    const auto& chart = charts.front();
    test.checkEqual(static_cast<int>(chart.type()), static_cast<int>(xlpp::Chart::Type::Scatter), "Scatter chart type parsed");
    test.checkEqual(chart.xAxisTitle(), std::string("X Axis"), "Scatter X title comes from first value axis");
    test.checkEqual(chart.yAxisTitle(), std::string("Y Axis"), "Scatter Y title comes from second value axis");
    test.checkEqual(chart.series().size(), std::size_t{2}, "Scatter multi-series count parsed");
    test.checkEqual(chart.series()[0].title(), std::string("Alpha"), "Scatter first series title parsed");
    test.checkEqual(chart.series()[1].title(), std::string("Beta"), "Scatter second series title parsed");
    test.checkTrue(chart.series()[1].categoriesReference().find("$A$2:$A$5") != std::string::npos, "Scatter second X reference parsed");
    test.checkTrue(chart.series()[1].valuesReference().find("$C$2:$C$5") != std::string::npos, "Scatter second Y reference parsed");

    const auto stableId = chart.stableId();
    test.checkTrue(!sheet->setChartLegend(stableId, true, "invalid"), "Invalid legend position rejected");
    test.checkTrue(sheet->setChartXAxisTitle(stableId, "Horizontal <X>"), "Scatter X-axis title selective edit");
    test.checkTrue(sheet->setChartYAxisTitle(stableId, "Vertical & Y"), "Scatter Y-axis title selective edit");
    test.checkTrue(sheet->setChartLegend(stableId, true, "b"), "Scatter legend selective edit");
    test.checkTrue(sheet->setChartSeriesTitle(stableId, 1, "Gamma & Delta"), "Scatter second series title selective edit");
    test.checkTrue(!sheet->setChartSeriesTitle(stableId, 5, "bad"), "Scatter invalid series-title index rejected");
    test.checkTrue(sheet->setChartSeriesReferences(stableId, 1, "'Scatter'!$A$2:$A$4", "'Scatter'!$C$2:$C$4"),
                   "Scatter xVal/yVal references selective edit");

    const auto output = std::filesystem::temp_directory_path() / "xlpp_p0i_scatter_selective.xlsx";
    workbook.save(output);
    const auto archive = xlpp::internal::ZipArchive::open(output);
    const auto xml = archive.get("xl/charts/chart1.xml");
    test.checkTrue(xml.find("Horizontal &lt;X&gt;") != std::string::npos, "Scatter X title XML escaped");
    test.checkTrue(xml.find("Vertical &amp; Y") != std::string::npos, "Scatter Y title XML escaped");
    test.checkTrue(xml.find("Gamma &amp; Delta") != std::string::npos, "Scatter series title XML escaped");
    test.checkTrue(xml.find("legendPos val=\"b\"") != std::string::npos, "Scatter legend position patched");
    test.checkTrue(xml.find("prstDash") != std::string::npos, "Scatter unsupported series line formatting preserved");

    const auto graph = xlpp::internal::RelationshipGraph::fromArchive(archive);
    const auto validation = graph.validate();
    test.checkTrue(validation.ok(), "Scatter selective output package graph validates");
    test.checkEqual(graph.objectInventory().charts, std::size_t{1}, "Scatter selective output chart count stable");
    test.checkEqual(graph.objectInventory().images, std::size_t{1}, "Scatter sibling image retained");

    xlpp::Workbook reloaded;
    reloaded.load(output);
    const auto& reloadedChart = static_cast<const xlpp::Worksheet&>(*reloaded.worksheet("Scatter")).charts().front();
    test.checkEqual(reloadedChart.xAxisTitle(), std::string("Horizontal <X>"), "Scatter X title survives reload");
    test.checkEqual(reloadedChart.yAxisTitle(), std::string("Vertical & Y"), "Scatter Y title survives reload");
    test.checkEqual(reloadedChart.legendPosition(), std::string("b"), "Scatter legend position survives reload");
    test.checkEqual(reloadedChart.series()[1].title(), std::string("Gamma & Delta"), "Scatter series title survives reload");
    test.checkTrue(reloadedChart.series()[1].categoriesReference().find("$A$2:$A$4") != std::string::npos, "Scatter edited X reference survives reload");
    test.checkTrue(reloadedChart.series()[1].valuesReference().find("$C$2:$C$4") != std::string::npos, "Scatter edited Y reference survives reload");

    const auto stableReloaded = reloadedChart.stableId();
    auto* reloadedSheet = reloaded.worksheet("Scatter");
    test.checkTrue(reloadedSheet->setChartLegend(stableReloaded, false), "Imported chart legend can be hidden selectively");
    const auto noLegend = std::filesystem::temp_directory_path() / "xlpp_p0i_scatter_no_legend.xlsx";
    reloaded.save(noLegend);
    const auto noLegendXml = xlpp::internal::ZipArchive::open(noLegend).get("xl/charts/chart1.xml");
    test.checkTrue(noLegendXml.find("<legend") == std::string::npos && noLegendXml.find("<c:legend") == std::string::npos, "Selective legend hide removes legend node");
    std::filesystem::remove(output);
    std::filesystem::remove(noLegend);
}

void testImportedCombinedChartAxisStructure(TestContext& test) {
    const auto source = std::filesystem::path(XLPP_TEST_SOURCE_DIR) / "fixtures/openpyxl/combined_secondary_axes.xlsx";
    xlpp::Workbook workbook;
    workbook.load(source);
    auto* sheet = workbook.worksheet("Data");
    test.checkTrue(sheet != nullptr, "Combined chart fixture worksheet loads");
    if (!sheet) return;
    const auto& charts = static_cast<const xlpp::Worksheet&>(*sheet).charts();
    test.checkEqual(charts.size(), std::size_t{1}, "Combined fixture exposes one chart object");
    if (charts.empty()) return;
    const auto& chart = charts.front();
    test.checkTrue(chart.combined(), "Bar + line chart is recognized as combined");
    test.checkTrue(chart.hasSecondaryAxes(), "Combined chart secondary axis is recognized");
    test.checkEqual(chart.plots().size(), std::size_t{2}, "Combined chart exposes two plots");
    test.checkEqual(chart.axes().size(), std::size_t{3}, "Combined chart exposes three axis definitions");
    test.checkEqual(chart.primaryXAxisId(), std::uint64_t{10}, "Primary X/category axis ID parsed from first plot");
    test.checkEqual(chart.primaryYAxisId(), std::uint64_t{100}, "Primary Y/value axis ID parsed from first plot");
    test.checkEqual(chart.xAxisTitle(), std::string("Month"), "Primary X axis title resolved by axId");
    test.checkEqual(chart.yAxisTitle(), std::string("Sales"), "Primary Y axis title resolved by axId");
    test.checkEqual(chart.series().size(), std::size_t{2}, "Combined chart series from both plots are loaded");

    test.checkEqual(static_cast<int>(chart.plots()[0].type), static_cast<int>(xlpp::Chart::Type::Bar), "First plot is bar");
    test.checkEqual(static_cast<int>(chart.plots()[1].type), static_cast<int>(xlpp::Chart::Type::Line), "Second plot is line");
    test.checkEqual(chart.plots()[0].firstSeries, std::size_t{0}, "Primary plot first series index");
    test.checkEqual(chart.plots()[0].seriesCount, std::size_t{1}, "Primary plot series count");
    test.checkEqual(chart.plots()[1].firstSeries, std::size_t{1}, "Secondary plot first series index");
    test.checkEqual(chart.plots()[1].seriesCount, std::size_t{1}, "Secondary plot series count");
    test.checkTrue(!chart.plots()[0].usesSecondaryAxes, "Primary plot does not use secondary axes");
    test.checkTrue(chart.plots()[1].usesSecondaryAxes, "Line plot is linked to secondary value axis");

    const auto* categoryAxis = chart.axisById(10);
    const auto* primaryValueAxis = chart.axisById(100);
    const auto* secondaryValueAxis = chart.axisById(200);
    test.checkTrue(categoryAxis != nullptr && primaryValueAxis != nullptr && secondaryValueAxis != nullptr,
                   "Axis lookup by native axId succeeds");
    if (categoryAxis && primaryValueAxis && secondaryValueAxis) {
        test.checkEqual(static_cast<int>(categoryAxis->kind), static_cast<int>(xlpp::Chart::AxisKind::Category), "Axis 10 is category axis");
        test.checkEqual(categoryAxis->crossAxisId, std::uint64_t{100}, "Category axis crossAx parsed");
        test.checkTrue(!categoryAxis->secondary, "Shared category axis remains primary");
        test.checkEqual(primaryValueAxis->crossAxisId, std::uint64_t{10}, "Primary value axis crossAx parsed");
        test.checkTrue(!primaryValueAxis->secondary, "Primary value axis classified primary");
        test.checkEqual(secondaryValueAxis->crossAxisId, std::uint64_t{10}, "Secondary value axis crossAx parsed");
        test.checkEqual(secondaryValueAxis->position, std::string("r"), "Secondary axis position parsed");
        test.checkEqual(secondaryValueAxis->title, std::string("Margin"), "Secondary axis title parsed");
        test.checkTrue(secondaryValueAxis->secondary, "Axis 200 classified secondary");
    }
    test.checkTrue(chart.axisById(9999) == nullptr, "Unknown axis ID lookup returns null");

    const auto stableId = chart.stableId();
    test.checkTrue(!sheet->setChartAxisTitle(stableId, 9999, "invalid"), "Unknown axis ID edit is rejected");
    test.checkTrue(sheet->setChartXAxisTitle(stableId, "Primary Category"), "Primary X title edits by native axis ID");
    test.checkTrue(sheet->setChartYAxisTitle(stableId, "Primary Sales"), "Primary Y title edits by native axis ID");
    test.checkTrue(sheet->setChartAxisTitle(stableId, 200, "Secondary Margin"), "Secondary axis title edits by native axis ID");
    test.checkTrue(sheet->setChartSeriesTitle(stableId, 1, "Margin series"), "Secondary plot series title edits selectively");
    sheet->cell("J20").setValue("combined-axis-regression");

    const auto output = std::filesystem::temp_directory_path() / "xlpp_p0j_combined_secondary_axes.xlsx";
    workbook.save(output);
    const auto archive = xlpp::internal::ZipArchive::open(output);
    const auto xml = archive.get("xl/charts/chart1.xml");
    test.checkTrue(xml.find("Primary Category") != std::string::npos, "Primary category title written");
    test.checkTrue(xml.find("Primary Sales") != std::string::npos, "Primary value title written");
    test.checkTrue(xml.find("Secondary Margin") != std::string::npos, "Secondary value title written");
    test.checkTrue(xml.find("Margin series") != std::string::npos, "Secondary plot series title written");
    test.checkTrue(xml.find("axId val=\"10\"") != std::string::npos &&
                   xml.find("axId val=\"100\"") != std::string::npos &&
                   xml.find("axId val=\"200\"") != std::string::npos,
                   "Selective edits preserve all native axis IDs");
    test.checkTrue(xml.find("crosses val=\"max\"") != std::string::npos, "Unsupported secondary-axis crosses metadata preserved");
    test.checkTrue(xml.find("<barChart") != std::string::npos && xml.find("<lineChart") != std::string::npos,
                   "Combined plot structure preserved");
    const auto graph = xlpp::internal::RelationshipGraph::fromArchive(archive);
    test.checkTrue(graph.validate().ok(), "Combined chart selective output package graph validates");
    test.checkEqual(graph.objectInventory().charts, std::size_t{1}, "Combined chart remains one visible drawing chart");

    xlpp::Workbook reloaded;
    reloaded.load(output);
    const auto* reloadedSheet = reloaded.worksheet("Data");
    test.checkTrue(reloadedSheet != nullptr, "Combined chart output reloads");
    if (reloadedSheet) {
        const auto& reloadedCharts = static_cast<const xlpp::Worksheet&>(*reloadedSheet).charts();
        test.checkEqual(reloadedCharts.size(), std::size_t{1}, "Combined chart count survives reload");
        if (!reloadedCharts.empty()) {
            const auto& reloadedChart = reloadedCharts.front();
            test.checkTrue(reloadedChart.combined(), "Combined plot identity survives reload");
            test.checkTrue(reloadedChart.hasSecondaryAxes(), "Secondary-axis identity survives reload");
            test.checkEqual(reloadedChart.primaryXAxisId(), std::uint64_t{10}, "Primary X axis ID survives reload");
            test.checkEqual(reloadedChart.primaryYAxisId(), std::uint64_t{100}, "Primary Y axis ID survives reload");
            test.checkEqual(reloadedChart.xAxisTitle(), std::string("Primary Category"), "Primary X title survives reload");
            test.checkEqual(reloadedChart.yAxisTitle(), std::string("Primary Sales"), "Primary Y title survives reload");
            const auto* secondary = reloadedChart.axisById(200);
            test.checkTrue(secondary != nullptr && secondary->title == "Secondary Margin" && secondary->secondary,
                           "Secondary axis title/classification survives reload");
        }
    }
    std::filesystem::remove(output);
}

void testImportedChartLabelsTrendlinesAndErrorBars(TestContext& test) {
    const auto source = std::filesystem::path(XLPP_TEST_SOURCE_DIR) / "fixtures/openpyxl/chart_labels_trendline_errorbars.xlsx";
    const auto before = xlpp::internal::ZipArchive::open(source);
    const auto originalDrawingRels = before.get("xl/drawings/_rels/drawing1.xml.rels");
    const auto originalImage = before.get("xl/media/image1.png");

    xlpp::Workbook workbook;
    workbook.load(source);
    auto* sheet = workbook.worksheet("Advanced");
    test.checkTrue(sheet != nullptr, "Advanced chart fixture worksheet loads");
    if (!sheet) return;
    const auto& charts = static_cast<const xlpp::Worksheet&>(*sheet).charts();
    test.checkEqual(charts.size(), std::size_t{1}, "Advanced fixture exposes one chart");
    if (charts.empty()) return;
    const auto& chart = charts.front();
    test.checkEqual(chart.plots().size(), std::size_t{1}, "Advanced scatter exposes one plot");
    test.checkTrue(chart.plots()[0].dataLabels.present, "Plot data labels are parsed");
    test.checkTrue(chart.plots()[0].dataLabels.showValue, "Data-label showVal parsed");
    test.checkTrue(chart.plots()[0].dataLabels.showSeriesName, "Data-label showSerName parsed");
    test.checkEqual(chart.plots()[0].dataLabels.position, std::string("t"), "Data-label position parsed");
    test.checkEqual(chart.series().size(), std::size_t{2}, "Advanced scatter series count parsed");
    test.checkEqual(chart.series()[0].trendlines().size(), std::size_t{1}, "First series trendline parsed");
    test.checkEqual(static_cast<int>(chart.series()[0].trendlines()[0].type),
                    static_cast<int>(xlpp::ChartSeries::TrendlineType::Linear), "Linear trendline type parsed");
    test.checkTrue(chart.series()[0].trendlines()[0].displayEquation, "Trendline equation flag parsed");
    test.checkTrue(chart.series()[0].trendlines()[0].displayRSquared, "Trendline R-squared flag parsed");
    test.checkEqual(chart.series()[1].trendlines().size(), std::size_t{1}, "Second series polynomial trendline parsed");
    test.checkTrue(chart.series()[1].dataLabels().present, "Series-level data labels are parsed");
    test.checkTrue(chart.series()[1].dataLabels().showCategoryName, "Series-level category-name flag parsed");
    test.checkEqual(chart.series()[1].dataLabels().position, std::string("r"), "Series-level data-label position parsed");
    test.checkEqual(chart.series()[1].trendlines()[0].order, 2, "Polynomial trendline order parsed");
    test.checkEqual(chart.series()[0].errorBars().size(), std::size_t{1}, "Series Y error bars parsed");
    test.checkEqual(static_cast<int>(chart.series()[0].errorBars()[0].direction),
                    static_cast<int>(xlpp::ChartSeries::ErrorBarDirection::Y), "Error-bar direction parsed");
    test.checkEqual(static_cast<int>(chart.series()[0].errorBars()[0].valueType),
                    static_cast<int>(xlpp::ChartSeries::ErrorValueType::FixedValue), "Error-bar value type parsed");
    test.checkNear(chart.series()[0].errorBars()[0].value, 1.5, 1e-9, "Fixed error-bar value parsed");

    const auto stableId = chart.stableId();
    auto labels = chart.plots()[0].dataLabels;
    labels.showCategoryName = true;
    labels.showSeriesName = false;
    labels.position = "b";
    labels.separator = " | ";
    test.checkTrue(sheet->setChartPlotDataLabels(stableId, 0, labels), "Plot data labels selectively edit");
    test.checkTrue(!sheet->setChartPlotDataLabels(stableId, 3, labels), "Invalid plot-index data-label edit rejected");

    auto seriesLabels = chart.series()[1].dataLabels();
    seriesLabels.showValue = true;
    seriesLabels.position = "l";
    seriesLabels.separator = " / ";
    test.checkTrue(sheet->setChartSeriesDataLabels(stableId, 1, seriesLabels), "Series-level data labels selectively edit");
    test.checkTrue(!sheet->setChartSeriesDataLabels(stableId, 8, seriesLabels), "Invalid series data-label index rejected");

    auto trendline = chart.series()[0].trendlines()[0];
    trendline.type = xlpp::ChartSeries::TrendlineType::Polynomial;
    trendline.order = 3;
    trendline.forward = 1.0;
    trendline.displayEquation = false;
    trendline.displayRSquared = true;
    test.checkTrue(sheet->setChartSeriesTrendline(stableId, 0, 0, trendline), "Existing trendline selectively edits");
    test.checkTrue(sheet->removeChartSeriesTrendline(stableId, 1, 0), "Existing second trendline selectively removes");
    xlpp::ChartSeries::Trendline exponential;
    exponential.type = xlpp::ChartSeries::TrendlineType::Exponential;
    exponential.displayEquation = true;
    test.checkTrue(sheet->addChartSeriesTrendline(stableId, 1, exponential), "Trendline selectively appends");

    auto yBars = chart.series()[0].errorBars()[0];
    yBars.barType = xlpp::ChartSeries::ErrorBarType::Plus;
    yBars.value = 2.25;
    yBars.noEndCap = true;
    test.checkTrue(sheet->setChartSeriesErrorBars(stableId, 0, yBars), "Existing Y error bars selectively edit");
    xlpp::ChartSeries::ErrorBars xBars;
    xBars.direction = xlpp::ChartSeries::ErrorBarDirection::X;
    xBars.barType = xlpp::ChartSeries::ErrorBarType::Both;
    xBars.valueType = xlpp::ChartSeries::ErrorValueType::Percentage;
    xBars.value = 5.0;
    test.checkTrue(sheet->setChartSeriesErrorBars(stableId, 1, xBars), "New X error bars selectively append");
    auto customBars = xBars;
    customBars.valueType = xlpp::ChartSeries::ErrorValueType::Custom;
    test.checkTrue(!sheet->setChartSeriesErrorBars(stableId, 1, customBars), "Custom error-bar write rejected until range refs are modeled");
    sheet->cell("J20").setValue("p0k-regression");

    const auto output = std::filesystem::temp_directory_path() / "xlpp_p0k_chart_features.xlsx";
    workbook.save(output);
    const auto after = xlpp::internal::ZipArchive::open(output);
    const auto xml = after.get("xl/charts/chart1.xml");
    test.checkTrue(xml.find("dLblPos val=\"b\"") != std::string::npos, "Data-label position selectively patched");
    test.checkTrue(xml.find("showCatName val=\"1\"") != std::string::npos, "Data-label category-name flag added");
    test.checkTrue(xml.find("showSerName val=\"0\"") != std::string::npos, "Data-label series-name flag disabled");
    test.checkTrue(xml.find("separator> | </") != std::string::npos, "Data-label separator selectively added");
    test.checkTrue(xml.find("dLblPos val=\"l\"") != std::string::npos && xml.find("separator> / </") != std::string::npos,
                   "Series-level data labels selectively patched");
    test.checkTrue(xml.find("trendlineType val=\"poly\"") != std::string::npos && xml.find("order val=\"3\"") != std::string::npos,
                   "Polynomial trendline selective patch written");
    test.checkTrue(xml.find("trendlineType val=\"exp\"") != std::string::npos, "Replacement exponential trendline written");
    test.checkTrue(xml.find("errBarType val=\"plus\"") != std::string::npos && xml.find("val val=\"2.25\"") != std::string::npos,
                   "Existing fixed Y error bars selectively patched");
    test.checkTrue(xml.find("errDir val=\"x\"") != std::string::npos && xml.find("errValType val=\"percentage\"") != std::string::npos,
                   "New X percentage error bars written");
    test.checkTrue(xml.find("prstDash") != std::string::npos, "Unsupported series formatting survives chart feature edits");
    test.checkEqual(after.get("xl/drawings/_rels/drawing1.xml.rels"), originalDrawingRels,
                    "Chart feature edits keep drawing relationships byte-identical");
    test.checkEqual(after.get("xl/media/image1.png"), originalImage, "Chart feature edits keep sibling image byte-identical");
    const auto graph = xlpp::internal::RelationshipGraph::fromArchive(after);
    test.checkTrue(graph.validate().ok(), "Chart labels/trendline/error-bar output package graph validates");
    test.checkEqual(graph.objectInventory().charts, std::size_t{1}, "Advanced chart remains visible after selective edits");
    test.checkEqual(graph.objectInventory().images, std::size_t{1}, "Sibling image remains visible after selective edits");

    xlpp::Workbook reloaded;
    reloaded.load(output);
    const auto* reloadSheet = reloaded.worksheet("Advanced");
    test.checkTrue(reloadSheet != nullptr, "Advanced selective output reloads");
    if (reloadSheet) {
        const auto& reloadChart = static_cast<const xlpp::Worksheet&>(*reloadSheet).charts().front();
        test.checkTrue(reloadChart.plots()[0].dataLabels.showCategoryName && !reloadChart.plots()[0].dataLabels.showSeriesName,
                       "Data-label flags survive reload");
        test.checkEqual(reloadChart.plots()[0].dataLabels.position, std::string("b"), "Data-label position survives reload");
        test.checkTrue(reloadChart.series()[1].dataLabels().present && reloadChart.series()[1].dataLabels().showValue,
                       "Series-level data labels survive reload");
        test.checkEqual(reloadChart.series()[1].dataLabels().position, std::string("l"), "Series-level data-label position survives reload");
        test.checkEqual(reloadChart.series()[0].trendlines().size(), std::size_t{1}, "Edited first trendline count survives reload");
        test.checkEqual(static_cast<int>(reloadChart.series()[0].trendlines()[0].type),
                        static_cast<int>(xlpp::ChartSeries::TrendlineType::Polynomial), "Edited trendline type survives reload");
        test.checkEqual(reloadChart.series()[0].trendlines()[0].order, 3, "Edited polynomial order survives reload");
        test.checkEqual(reloadChart.series()[1].trendlines().size(), std::size_t{1}, "Replacement second trendline count survives reload");
        test.checkEqual(static_cast<int>(reloadChart.series()[1].trendlines()[0].type),
                        static_cast<int>(xlpp::ChartSeries::TrendlineType::Exponential), "Replacement trendline survives reload");
        test.checkEqual(reloadChart.series()[0].errorBars().size(), std::size_t{1}, "Edited Y error-bar count survives reload");
        test.checkNear(reloadChart.series()[0].errorBars()[0].value, 2.25, 1e-9, "Edited error-bar value survives reload");
        test.checkEqual(reloadChart.series()[1].errorBars().size(), std::size_t{1}, "Appended X error bars survive reload");
        test.checkEqual(static_cast<int>(reloadChart.series()[1].errorBars()[0].direction),
                        static_cast<int>(xlpp::ChartSeries::ErrorBarDirection::X), "Appended X error-bar direction survives reload");

        const auto reloadId = reloadChart.stableId();
        auto* mutableReload = reloaded.worksheet("Advanced");
        test.checkTrue(mutableReload->removeChartSeriesErrorBars(reloadId, 0, xlpp::ChartSeries::ErrorBarDirection::Y),
                       "Selective error-bar removal supported");
        test.checkTrue(mutableReload->removeChartSeriesTrendline(reloadId, 1, 0), "Selective trendline removal supported after reload");
        const auto removedOutput = std::filesystem::temp_directory_path() / "xlpp_p0k_chart_features_removed.xlsx";
        reloaded.save(removedOutput);
        const auto removedXml = xlpp::internal::ZipArchive::open(removedOutput).get("xl/charts/chart1.xml");
        test.checkTrue(removedXml.find("errDir val=\"y\"") == std::string::npos, "Selective Y error bars removed from XML");
        test.checkTrue(removedXml.find("trendlineType val=\"exp\"") == std::string::npos, "Selective exponential trendline removed from XML");
        std::filesystem::remove(removedOutput);
    }
    std::filesystem::remove(output);
}

void testImportedChartPerPointCustomErrorsAndFormatting(TestContext& test) {
    const auto source = std::filesystem::path(XLPP_TEST_SOURCE_DIR) / "fixtures/openpyxl/per_point_custom_errors_formatting.xlsx";
    const auto before = xlpp::internal::ZipArchive::open(source);
    const auto originalDrawingRels = before.get("xl/drawings/_rels/drawing1.xml.rels");
    const auto originalImage = before.get("xl/media/image1.png");

    xlpp::Workbook workbook;
    workbook.load(source);
    auto* sheet = workbook.worksheet("P0L");
    test.checkTrue(sheet != nullptr, "P0L OpenPyXL fixture loads");
    if (!sheet) return;
    const auto& charts = static_cast<const xlpp::Worksheet&>(*sheet).charts();
    test.checkEqual(charts.size(), std::size_t{1}, "P0L fixture exposes one imported chart");
    if (charts.empty() || charts.front().series().empty()) return;
    const auto& chart = charts.front();
    const auto& series = chart.series().front();

    test.checkTrue(series.dataLabels().present && series.dataLabels().showValue, "Series aggregate data-label defaults parsed independently");
    test.checkEqual(series.dataLabels().points.size(), std::size_t{2}, "Per-point data labels parsed");
    if (series.dataLabels().points.size() >= 2) {
        test.checkEqual(series.dataLabels().points[0].index, std::size_t{1}, "First point-label index parsed");
        test.checkTrue(series.dataLabels().points[0].showSeriesName && !series.dataLabels().points[0].showValue,
                       "First point-label overrides parsed");
        test.checkEqual(series.dataLabels().points[0].position, std::string("t"), "First point-label position parsed");
        test.checkEqual(series.dataLabels().points[1].index, std::size_t{3}, "Second point-label index parsed");
        test.checkTrue(series.dataLabels().points[1].showCategoryName && series.dataLabels().points[1].showValue,
                       "Second point-label overrides parsed");
    }

    test.checkTrue(series.lineFormat().present, "Series line formatting parsed");
    test.checkEqual(static_cast<int>(series.lineFormat().color.kind), static_cast<int>(xlpp::ChartColor::Kind::SRgb), "Series line color kind parsed");
    test.checkEqual(series.lineFormat().color.value, std::string("FF0000"), "Series line color parsed");
    test.checkNear(series.lineFormat().widthPoints, 2.0, 1e-9, "Series line width parsed in points");
    test.checkEqual(series.lineFormat().dash, std::string("dash"), "Series dash style parsed");
    test.checkTrue(series.markerFormat().present, "Marker formatting parsed");
    test.checkEqual(series.markerFormat().symbol, std::string("diamond"), "Marker symbol parsed");
    test.checkEqual(series.markerFormat().size, 9, "Marker size parsed");
    test.checkEqual(series.markerFormat().fill.color.value, std::string("00FF00"), "Marker fill parsed");
    test.checkEqual(series.markerFormat().line.color.value, std::string("0000FF"), "Marker outline parsed");

    test.checkEqual(series.trendlines().size(), std::size_t{1}, "Formatted trendline parsed");
    if (!series.trendlines().empty()) {
        test.checkTrue(series.trendlines()[0].lineFormat.present, "Trendline line formatting parsed");
        test.checkEqual(series.trendlines()[0].lineFormat.color.value, std::string("AA00AA"), "Trendline line color parsed");
        test.checkEqual(series.trendlines()[0].lineFormat.dash, std::string("dot"), "Trendline dash parsed");
        test.checkNear(series.trendlines()[0].lineFormat.widthPoints, 1.5, 1e-9, "Trendline width parsed");
    }
    test.checkEqual(series.errorBars().size(), std::size_t{1}, "Custom error bars parsed");
    if (!series.errorBars().empty()) {
        const auto& bars = series.errorBars()[0];
        test.checkEqual(static_cast<int>(bars.valueType), static_cast<int>(xlpp::ChartSeries::ErrorValueType::Custom), "Custom error-bar value type parsed");
        test.checkEqual(bars.plusReference, std::string("'P0L'!$C$2:$C$6"), "Custom plus reference parsed");
        test.checkEqual(bars.minusReference, std::string("'P0L'!$D$2:$D$6"), "Custom minus reference parsed");
        test.checkTrue(bars.lineFormat.present, "Error-bar line formatting parsed");
        test.checkEqual(bars.lineFormat.color.value, std::string("444444"), "Error-bar line color parsed");
        test.checkEqual(bars.lineFormat.dash, std::string("sysDot"), "Error-bar dash parsed");
    }

    const auto stableId = chart.stableId();
    auto aggregateLabels = series.dataLabels();
    aggregateLabels.showValue = false;
    aggregateLabels.showCategoryName = true;
    aggregateLabels.position = "ctr";
    aggregateLabels.separator = " / ";
    test.checkTrue(sheet->setChartSeriesDataLabels(stableId, 0, aggregateLabels),
                   "Aggregate series labels selectively edit without touching point labels");

    auto point1 = series.dataLabels().points.front();
    point1.showValue = true;
    point1.showSeriesName = false;
    point1.position = "b";
    point1.separator = " | ";
    test.checkTrue(sheet->setChartSeriesDataLabelPoint(stableId, 0, point1), "Existing point label selectively edits");
    test.checkTrue(sheet->removeChartSeriesDataLabelPoint(stableId, 0, 3), "Existing point label selectively removes");
    xlpp::ChartDataLabelPoint point4;
    point4.index = 4;
    point4.showCategoryName = true;
    point4.position = "l";
    test.checkTrue(sheet->setChartSeriesDataLabelPoint(stableId, 0, point4), "New point label selectively appends");
    xlpp::ChartDataLabelPoint plotPoint;
    plotPoint.index = 2;
    plotPoint.showValue = true;
    plotPoint.position = "r";
    test.checkTrue(sheet->setChartPlotDataLabelPoint(stableId, 0, plotPoint), "Plot-level point label selectively appends");

    auto customBars = series.errorBars().front();
    customBars.noEndCap = true;
    customBars.plusReference = "'P0L'!$D$2:$D$6";
    customBars.minusReference = "'P0L'!$C$2:$C$6";
    test.checkTrue(sheet->setChartSeriesErrorBars(stableId, 0, customBars), "Custom error-bar ranges selectively edit");
    auto invalidCustom = customBars;
    invalidCustom.plusReference.clear();
    test.checkTrue(!sheet->setChartSeriesErrorBars(stableId, 0, invalidCustom), "Custom error bars require plus/minus references");

    xlpp::ChartLineFormat seriesLine;
    seriesLine.present = true;
    seriesLine.color = {xlpp::ChartColor::Kind::SRgb, "112233"};
    seriesLine.widthPoints = 2.5;
    seriesLine.dash = "lgDash";
    test.checkTrue(sheet->setChartSeriesLineFormat(stableId, 0, seriesLine), "Series line selectively formats");
    xlpp::ChartFillFormat seriesFill;
    seriesFill.present = true;
    seriesFill.color = {xlpp::ChartColor::Kind::SRgb, "ABCDEF"};
    test.checkTrue(sheet->setChartSeriesFillFormat(stableId, 0, seriesFill), "Series fill selectively formats");
    auto marker = series.markerFormat();
    marker.symbol = "triangle";
    marker.size = 11;
    marker.fill.present = true;
    marker.fill.color = {xlpp::ChartColor::Kind::SRgb, "123456"};
    marker.line.present = true;
    marker.line.color = {xlpp::ChartColor::Kind::SRgb, "654321"};
    marker.line.widthPoints = 1.25;
    marker.line.dash = "solid";
    test.checkTrue(sheet->setChartSeriesMarkerFormat(stableId, 0, marker), "Marker selectively formats");

    xlpp::ChartLineFormat trendLine;
    trendLine.present = true;
    trendLine.color = {xlpp::ChartColor::Kind::SRgb, "00AAAA"};
    trendLine.widthPoints = 2.0;
    trendLine.dash = "dashDot";
    test.checkTrue(sheet->setChartSeriesTrendlineLineFormat(stableId, 0, 0, trendLine), "Trendline line selectively formats");
    xlpp::ChartLineFormat errorLine;
    errorLine.present = true;
    errorLine.color = {xlpp::ChartColor::Kind::SRgb, "333333"};
    errorLine.widthPoints = 1.25;
    errorLine.dash = "dash";
    test.checkTrue(sheet->setChartSeriesErrorBarsLineFormat(stableId, 0, xlpp::ChartSeries::ErrorBarDirection::Y, errorLine),
                   "Error-bar line selectively formats");

    xlpp::ChartSeries::Trendline addedTrendline;
    addedTrendline.type = xlpp::ChartSeries::TrendlineType::Polynomial;
    addedTrendline.order = 3;
    addedTrendline.displayEquation = true;
    addedTrendline.lineFormat.present = true;
    addedTrendline.lineFormat.color = {xlpp::ChartColor::Kind::SRgb, "CC5500"};
    addedTrendline.lineFormat.widthPoints = 1.75;
    addedTrendline.lineFormat.dash = "sysDash";
    test.checkTrue(sheet->addChartSeriesTrendline(stableId, 0, addedTrendline),
                   "New formatted trendline selectively appends");

    xlpp::ChartSeries::ErrorBars xBars;
    xBars.direction = xlpp::ChartSeries::ErrorBarDirection::X;
    xBars.barType = xlpp::ChartSeries::ErrorBarType::Both;
    xBars.valueType = xlpp::ChartSeries::ErrorValueType::Custom;
    xBars.plusReference = "'P0L'!$C$2:$C$6";
    xBars.minusReference = "'P0L'!$D$2:$D$6";
    xBars.lineFormat.present = true;
    xBars.lineFormat.color = {xlpp::ChartColor::Kind::SRgb, "0055CC"};
    xBars.lineFormat.widthPoints = 1.5;
    xBars.lineFormat.dash = "sysDashDot";
    test.checkTrue(sheet->setChartSeriesErrorBars(stableId, 0, xBars),
                   "New formatted custom X error bars selectively append");

    sheet->cell("J20").setValue("p0l-regression");
    const auto output = std::filesystem::temp_directory_path() / "xlpp_p0l_chart_features.xlsx";
    workbook.save(output);
    const auto after = xlpp::internal::ZipArchive::open(output);
    const auto xml = after.get("xl/charts/chart1.xml");
    test.checkTrue(xml.find("idx val=\"1\"") != std::string::npos && xml.find("dLblPos val=\"b\"") != std::string::npos,
                   "Edited point-label XML written");
    test.checkTrue(xml.find("dLblPos val=\"ctr\"") != std::string::npos && xml.find("<showCatName val=\"1\"") != std::string::npos,
                   "Aggregate data-label XML written separately from point labels");
    test.checkTrue(xml.find("idx val=\"3\"") == std::string::npos, "Removed point label absent from XML");
    test.checkTrue(xml.find("idx val=\"4\"") != std::string::npos, "Appended point label written");
    test.checkTrue(xml.find("&apos;P0L&apos;!$D$2:$D$6") != std::string::npos && xml.find("&apos;P0L&apos;!$C$2:$C$6") != std::string::npos,
                   "Custom plus/minus error-bar references remain in ChartML");
    test.checkTrue(xml.find("112233") != std::string::npos && xml.find("ABCDEF") != std::string::npos,
                   "Series line/fill colors written");
    test.checkTrue(xml.find("triangle") != std::string::npos && xml.find("123456") != std::string::npos && xml.find("654321") != std::string::npos,
                   "Marker formatting written");
    test.checkTrue(xml.find("00AAAA") != std::string::npos && xml.find("333333") != std::string::npos,
                   "Trendline and error-bar line colors written");
    test.checkTrue(xml.find("CC5500") != std::string::npos && xml.find("0055CC") != std::string::npos,
                   "Formatting for newly appended trendline and custom X error bars written");
    test.checkEqual(after.get("xl/drawings/_rels/drawing1.xml.rels"), originalDrawingRels,
                    "P0L chart edits keep drawing relationships byte-identical");
    test.checkEqual(after.get("xl/media/image1.png"), originalImage, "P0L chart edits keep sibling image byte-identical");
    const auto graph = xlpp::internal::RelationshipGraph::fromArchive(after);
    test.checkTrue(graph.validate().ok(), "P0L output package graph validates");
    test.checkEqual(graph.objectInventory().charts, std::size_t{1}, "P0L chart remains visible");
    test.checkEqual(graph.objectInventory().images, std::size_t{1}, "P0L sibling image remains visible");

    xlpp::Workbook reloaded;
    reloaded.load(output);
    const auto* reloadSheet = reloaded.worksheet("P0L");
    test.checkTrue(reloadSheet != nullptr, "P0L selective output reloads");
    if (reloadSheet) {
        const auto& reloadSeries = static_cast<const xlpp::Worksheet&>(*reloadSheet).charts().front().series().front();
        test.checkTrue(reloadSeries.dataLabels().present && !reloadSeries.dataLabels().showValue && reloadSeries.dataLabels().showCategoryName,
                       "Aggregate data-label flags survive reload without inheriting point overrides");
        test.checkEqual(reloadSeries.dataLabels().position, std::string("ctr"), "Aggregate data-label position survives reload");
        test.checkEqual(reloadSeries.dataLabels().separator, std::string(" / "), "Aggregate data-label separator survives reload");
        test.checkEqual(reloadSeries.dataLabels().points.size(), std::size_t{2}, "Point-label count survives reload after remove/add");
        const auto pointIndex1 = std::find_if(reloadSeries.dataLabels().points.begin(), reloadSeries.dataLabels().points.end(), [](const auto& point) { return point.index == 1; });
        const auto pointIndex4 = std::find_if(reloadSeries.dataLabels().points.begin(), reloadSeries.dataLabels().points.end(), [](const auto& point) { return point.index == 4; });
        test.checkTrue(pointIndex1 != reloadSeries.dataLabels().points.end() && pointIndex1->showValue && !pointIndex1->showSeriesName,
                       "Edited point label survives reload");
        test.checkTrue(pointIndex4 != reloadSeries.dataLabels().points.end() && pointIndex4->showCategoryName,
                       "Appended point label survives reload");
        test.checkEqual(reloadSeries.errorBars().front().plusReference, std::string("'P0L'!$D$2:$D$6"), "Edited custom plus reference survives reload");
        test.checkEqual(reloadSeries.errorBars().front().minusReference, std::string("'P0L'!$C$2:$C$6"), "Edited custom minus reference survives reload");
        test.checkEqual(reloadSeries.lineFormat().color.value, std::string("112233"), "Edited series line color survives reload");
        test.checkNear(reloadSeries.lineFormat().widthPoints, 2.5, 1e-9, "Edited series line width survives reload");
        test.checkEqual(reloadSeries.fillFormat().color.value, std::string("ABCDEF"), "Edited series fill survives reload");
        test.checkEqual(reloadSeries.markerFormat().symbol, std::string("triangle"), "Edited marker symbol survives reload");
        test.checkEqual(reloadSeries.markerFormat().fill.color.value, std::string("123456"), "Edited marker fill survives reload");
        test.checkEqual(reloadSeries.trendlines().front().lineFormat.color.value, std::string("00AAAA"), "Edited trendline formatting survives reload");
        test.checkEqual(reloadSeries.trendlines().size(), std::size_t{2}, "New formatted trendline survives reload");
        if (reloadSeries.trendlines().size() >= 2) {
            test.checkEqual(static_cast<int>(reloadSeries.trendlines()[1].type), static_cast<int>(xlpp::ChartSeries::TrendlineType::Polynomial),
                            "New polynomial trendline type survives reload");
            test.checkEqual(reloadSeries.trendlines()[1].order, 3, "New polynomial trendline order survives reload");
            test.checkEqual(reloadSeries.trendlines()[1].lineFormat.color.value, std::string("CC5500"),
                            "New trendline line formatting survives reload");
        }
        const auto yBars = std::find_if(reloadSeries.errorBars().begin(), reloadSeries.errorBars().end(), [](const auto& bars) {
            return bars.direction == xlpp::ChartSeries::ErrorBarDirection::Y;
        });
        const auto xBarsReloaded = std::find_if(reloadSeries.errorBars().begin(), reloadSeries.errorBars().end(), [](const auto& bars) {
            return bars.direction == xlpp::ChartSeries::ErrorBarDirection::X;
        });
        test.checkTrue(yBars != reloadSeries.errorBars().end() && yBars->lineFormat.color.value == "333333",
                       "Edited Y error-bar formatting survives reload");
        test.checkTrue(xBarsReloaded != reloadSeries.errorBars().end(), "New custom X error bars survive reload");
        if (xBarsReloaded != reloadSeries.errorBars().end()) {
            test.checkEqual(xBarsReloaded->plusReference, std::string("'P0L'!$C$2:$C$6"), "New X plus reference survives reload");
            test.checkEqual(xBarsReloaded->minusReference, std::string("'P0L'!$D$2:$D$6"), "New X minus reference survives reload");
            test.checkEqual(xBarsReloaded->lineFormat.color.value, std::string("0055CC"), "New X error-bar formatting survives reload");
        }
    }
    std::filesystem::remove(output);
}

void testImportedChartDataPointRichTextAndAdvancedFormatting(TestContext& test) {
    const auto source = std::filesystem::path(XLPP_TEST_SOURCE_DIR) / "fixtures/openpyxl/data_point_rich_text_advanced_formatting.xlsx";
    const auto before = xlpp::internal::ZipArchive::open(source);
    const auto originalDrawingRels = before.get("xl/drawings/_rels/drawing1.xml.rels");
    const auto originalImage = before.get("xl/media/image1.png");

    xlpp::Workbook workbook;
    workbook.load(source);
    auto* sheet = workbook.worksheet("P0L");
    test.checkTrue(sheet != nullptr, "P0M advanced ChartML fixture loads");
    if (!sheet) return;
    const auto& charts = static_cast<const xlpp::Worksheet&>(*sheet).charts();
    test.checkEqual(charts.size(), std::size_t{1}, "P0M fixture exposes one chart");
    if (charts.empty() || charts.front().series().empty()) return;
    const auto& chart = charts.front();
    const auto& series = chart.series().front();

    test.checkTrue(chart.titleRichText().present, "Rich chart title detected");
    test.checkEqual(chart.titleRichText().runs.size(), std::size_t{2}, "Rich chart title runs parsed");
    test.checkEqual(chart.titleRichText().plainText(), std::string("P0M advanced"), "Rich chart title plain text concatenated");
    if (chart.titleRichText().runs.size() >= 2) {
        const auto& first = chart.titleRichText().runs[0];
        const auto& second = chart.titleRichText().runs[1];
        test.checkTrue(first.bold && !first.italic, "First title run bold metadata parsed");
        test.checkNear(first.fontSizePoints, 16.0, 1e-9, "First title run font size parsed");
        test.checkEqual(first.typeface, std::string("Aptos"), "First title run typeface parsed");
        test.checkEqual(static_cast<int>(first.color.kind), static_cast<int>(xlpp::ChartColor::Kind::Scheme), "First title run scheme color parsed");
        test.checkEqual(first.color.value, std::string("accent1"), "First title run scheme value parsed");
        test.checkEqual(first.color.transforms.size(), std::size_t{1}, "Title color transform parsed");
        if (!first.color.transforms.empty())
            test.checkEqual(first.color.transforms.front().value, 20000, "Title tint transform value parsed");
        test.checkTrue(second.italic, "Second title run italic metadata parsed");
        test.checkEqual(second.color.value, std::string("336699"), "Second title run RGB parsed");
    }

    test.checkTrue(series.fillFormat().present, "Series advanced fill parsed");
    test.checkEqual(static_cast<int>(series.fillFormat().kind), static_cast<int>(xlpp::ChartFillFormat::Kind::Gradient), "Series gradient fill kind parsed");
    test.checkEqual(series.fillFormat().gradientStops.size(), std::size_t{2}, "Gradient stops parsed");
    test.checkNear(series.fillFormat().gradientAngleDegrees, 45.0, 1e-9, "Gradient angle parsed");
    if (series.fillFormat().gradientStops.size() >= 2) {
        test.checkEqual(series.fillFormat().gradientStops[0].position, 0, "First gradient stop position parsed");
        test.checkEqual(series.fillFormat().gradientStops[0].color.value, std::string("FF0000"), "First gradient stop color parsed");
        test.checkEqual(series.fillFormat().gradientStops[0].color.transforms.size(), std::size_t{1}, "Gradient alpha transform parsed");
        test.checkEqual(series.fillFormat().gradientStops[1].color.value, std::string("accent3"), "Second gradient scheme color parsed");
    }

    test.checkEqual(series.lineFormat().cap, std::string("rnd"), "Advanced line cap parsed");
    test.checkEqual(series.lineFormat().compound, std::string("dbl"), "Advanced compound line parsed");
    test.checkEqual(series.lineFormat().join, std::string("round"), "Advanced line join parsed");
    test.checkEqual(series.lineFormat().customDash.size(), std::size_t{2}, "Custom dash sequence parsed");
    test.checkEqual(series.lineFormat().color.value, std::string("accent2"), "Series scheme line color parsed");
    test.checkEqual(series.lineFormat().color.transforms.size(), std::size_t{2}, "Series line color transforms parsed");

    test.checkEqual(series.dataPoints().size(), std::size_t{1}, "Per-data-point style parsed");
    const auto* point2 = series.dataPoint(2);
    test.checkTrue(point2 != nullptr, "dPt index 2 accessible by index");
    if (point2) {
        test.checkEqual(static_cast<int>(point2->fill.kind), static_cast<int>(xlpp::ChartFillFormat::Kind::Pattern), "dPt pattern fill parsed");
        test.checkEqual(point2->fill.pattern, std::string("pct20"), "dPt pattern type parsed");
        test.checkEqual(point2->fill.foregroundColor.value, std::string("00AA00"), "dPt foreground color parsed");
        test.checkEqual(point2->fill.backgroundColor.value, std::string("accent4"), "dPt background scheme color parsed");
        test.checkEqual(point2->line.cap, std::string("sq"), "dPt line cap parsed");
        test.checkEqual(point2->line.join, std::string("bevel"), "dPt line join parsed");
        test.checkEqual(point2->line.color.transforms.size(), std::size_t{1}, "dPt line alpha parsed");
        test.checkEqual(point2->marker.symbol, std::string("square"), "dPt marker symbol parsed");
        test.checkEqual(point2->marker.size, 7, "dPt marker size parsed");
    }

    const auto label1 = std::find_if(series.dataLabels().points.begin(), series.dataLabels().points.end(),
                                     [](const auto& point) { return point.index == 1; });
    test.checkTrue(label1 != series.dataLabels().points.end(), "Rich point label found");
    if (label1 != series.dataLabels().points.end()) {
        test.checkTrue(label1->richText.present, "Point-label rich text parsed");
        test.checkEqual(label1->richText.plainText(), std::string("Point One"), "Point-label rich text value parsed");
        test.checkEqual(label1->richText.runs.size(), std::size_t{1}, "Point-label rich run parsed");
        if (!label1->richText.runs.empty()) {
            test.checkTrue(label1->richText.runs.front().italic, "Point-label italic run metadata parsed");
            test.checkEqual(label1->richText.runs.front().color.value, std::string("accent5"), "Point-label rich color parsed");
        }
    }

    const auto stableId = chart.stableId();
    xlpp::ChartRichText newTitle;
    newTitle.present = true;
    xlpp::ChartTextRun titleA;
    titleA.text = "XL++ ";
    titleA.bold = true;
    titleA.fontSizePoints = 18.0;
    titleA.typeface = "Aptos Display";
    titleA.color = {xlpp::ChartColor::Kind::Scheme, "accent6"};
    titleA.color.transforms.push_back({xlpp::ChartColorTransform::Kind::Tint, 15000});
    xlpp::ChartTextRun titleB;
    titleB.text = "P0M";
    titleB.italic = true;
    titleB.fontSizePoints = 14.0;
    titleB.color = {xlpp::ChartColor::Kind::SRgb, "224466"};
    titleB.color.transforms.push_back({xlpp::ChartColorTransform::Kind::Alpha, 90000});
    newTitle.runs = {titleA, titleB};
    test.checkTrue(sheet->setChartTitleRichText(stableId, newTitle), "Rich chart title selectively edits");

    xlpp::ChartRichText labelRich;
    labelRich.present = true;
    xlpp::ChartTextRun labelRun;
    labelRun.text = "Edited point";
    labelRun.bold = true;
    labelRun.fontSizePoints = 11.0;
    labelRun.color = {xlpp::ChartColor::Kind::Scheme, "accent2"};
    labelRun.color.transforms.push_back({xlpp::ChartColorTransform::Kind::Shade, 65000});
    labelRich.runs.push_back(labelRun);
    test.checkTrue(sheet->setChartSeriesDataLabelPointRichText(stableId, 0, 1, labelRich),
                   "Point-label rich text selectively edits");

    xlpp::ChartLineFormat advancedLine = series.lineFormat();
    advancedLine.present = true;
    advancedLine.color = {xlpp::ChartColor::Kind::Scheme, "accent4"};
    advancedLine.color.transforms = {{xlpp::ChartColorTransform::Kind::LumMod, 80000},
                                     {xlpp::ChartColorTransform::Kind::LumOff, 10000}};
    advancedLine.widthPoints = 3.0;
    advancedLine.dash.clear();
    advancedLine.customDash = {{3.0, 1.0}, {1.0, 1.5}};
    advancedLine.cap = "sq";
    advancedLine.compound = "thickThin";
    advancedLine.join = "miter";
    test.checkTrue(sheet->setChartSeriesLineFormat(stableId, 0, advancedLine), "Advanced series line selectively edits");

    xlpp::ChartFillFormat gradient;
    gradient.present = true;
    gradient.kind = xlpp::ChartFillFormat::Kind::Gradient;
    gradient.gradientAngleDegrees = 90.0;
    xlpp::ChartGradientStop gs1;
    gs1.position = 0;
    gs1.color = {xlpp::ChartColor::Kind::Scheme, "accent1"};
    gs1.color.transforms.push_back({xlpp::ChartColorTransform::Kind::Tint, 25000});
    xlpp::ChartGradientStop gs2;
    gs2.position = 100000;
    gs2.color = {xlpp::ChartColor::Kind::SRgb, "112244"};
    gs2.color.transforms.push_back({xlpp::ChartColorTransform::Kind::Alpha, 75000});
    gradient.gradientStops = {gs1, gs2};
    test.checkTrue(sheet->setChartSeriesFillFormat(stableId, 0, gradient), "Advanced series gradient selectively edits");

    xlpp::ChartDataPointFormat editedPoint;
    editedPoint.index = 2;
    editedPoint.fill.present = true;
    editedPoint.fill.kind = xlpp::ChartFillFormat::Kind::Pattern;
    editedPoint.fill.pattern = "diagCross";
    editedPoint.fill.foregroundColor = {xlpp::ChartColor::Kind::Scheme, "accent5"};
    editedPoint.fill.foregroundColor.transforms.push_back({xlpp::ChartColorTransform::Kind::Tint, 12000});
    editedPoint.fill.backgroundColor = {xlpp::ChartColor::Kind::SRgb, "F0F0F0"};
    editedPoint.line.present = true;
    editedPoint.line.color = {xlpp::ChartColor::Kind::SRgb, "101010"};
    editedPoint.line.widthPoints = 1.75;
    editedPoint.line.cap = "rnd";
    editedPoint.line.join = "miter";
    editedPoint.line.customDash = {{2.0, 1.0}};
    editedPoint.marker.present = true;
    editedPoint.marker.symbol = "diamond";
    editedPoint.marker.size = 10;
    editedPoint.marker.fill.present = true;
    editedPoint.marker.fill.kind = xlpp::ChartFillFormat::Kind::Solid;
    editedPoint.marker.fill.color = {xlpp::ChartColor::Kind::SRgb, "00BBBB"};
    test.checkTrue(sheet->setChartSeriesDataPointFormat(stableId, 0, editedPoint), "Existing dPt selectively edits");

    xlpp::ChartDataPointFormat newPoint;
    newPoint.index = 4;
    newPoint.fill.present = true;
    newPoint.fill.kind = xlpp::ChartFillFormat::Kind::Gradient;
    newPoint.fill.gradientAngleDegrees = 30.0;
    xlpp::ChartGradientStop np1;
    np1.position = 0;
    np1.color = {xlpp::ChartColor::Kind::SRgb, "AA0000"};
    xlpp::ChartGradientStop np2;
    np2.position = 100000;
    np2.color = {xlpp::ChartColor::Kind::Scheme, "accent6"};
    newPoint.fill.gradientStops = {np1, np2};
    newPoint.line.present = true;
    newPoint.line.color = {xlpp::ChartColor::Kind::Scheme, "accent3"};
    newPoint.line.cap = "flat";
    newPoint.line.join = "round";
    newPoint.marker.present = true;
    newPoint.marker.symbol = "triangle";
    newPoint.marker.size = 8;
    test.checkTrue(sheet->setChartSeriesDataPointFormat(stableId, 0, newPoint), "New dPt selectively appends");

    sheet->cell("K21").setValue("p0m-regression");
    const auto output = std::filesystem::temp_directory_path() / "xlpp_p0m_chart_features.xlsx";
    workbook.save(output);
    const auto after = xlpp::internal::ZipArchive::open(output);
    const auto xml = after.get("xl/charts/chart1.xml");
    test.checkTrue(xml.find("Aptos Display") != std::string::npos && xml.find("XL++ ") != std::string::npos,
                   "Rich title runs written to ChartML");
    test.checkTrue(xml.find("Edited point") != std::string::npos, "Rich point-label text written");
    test.checkTrue(xml.find("idx val=\"2\"") != std::string::npos && xml.find("diagCross") != std::string::npos,
                   "Existing dPt advanced formatting written");
    test.checkTrue(xml.find("idx val=\"4\"") != std::string::npos && xml.find("triangle") != std::string::npos,
                   "New dPt written");
    test.checkTrue(xml.find("custDash") != std::string::npos && xml.find("cmpd=\"thickThin\"") != std::string::npos,
                   "Advanced custom dash and compound line written");
    test.checkTrue(xml.find("gradFill") != std::string::npos && xml.find("pattFill") != std::string::npos,
                   "Gradient and pattern fills remain in ChartML");
    test.checkTrue(xml.find("lumMod") != std::string::npos && xml.find("alpha") != std::string::npos && xml.find("tint") != std::string::npos,
                   "Color transforms written");
    test.checkEqual(after.get("xl/drawings/_rels/drawing1.xml.rels"), originalDrawingRels,
                    "P0M chart edits keep drawing relationships byte-identical");
    test.checkEqual(after.get("xl/media/image1.png"), originalImage, "P0M chart edits keep sibling image byte-identical");
    const auto graph = xlpp::internal::RelationshipGraph::fromArchive(after);
    test.checkTrue(graph.validate().ok(), "P0M output package graph validates");

    xlpp::Workbook reloaded;
    reloaded.load(output);
    const auto* reloadSheet = reloaded.worksheet("P0L");
    test.checkTrue(reloadSheet != nullptr, "P0M output reloads");
    if (reloadSheet) {
        const auto& reloadChart = static_cast<const xlpp::Worksheet&>(*reloadSheet).charts().front();
        const auto& reloadSeries = reloadChart.series().front();
        test.checkEqual(reloadChart.titleRichText().plainText(), std::string("XL++ P0M"), "Edited rich chart title survives reload");
        test.checkEqual(reloadChart.titleRichText().runs.size(), std::size_t{2}, "Edited title run count survives reload");
        test.checkEqual(reloadSeries.lineFormat().cap, std::string("sq"), "Edited line cap survives reload");
        test.checkEqual(reloadSeries.lineFormat().compound, std::string("thickThin"), "Edited compound line survives reload");
        test.checkEqual(reloadSeries.lineFormat().customDash.size(), std::size_t{2}, "Edited custom dash survives reload");
        test.checkEqual(static_cast<int>(reloadSeries.fillFormat().kind), static_cast<int>(xlpp::ChartFillFormat::Kind::Gradient),
                        "Edited gradient fill survives reload");
        test.checkNear(reloadSeries.fillFormat().gradientAngleDegrees, 90.0, 1e-9, "Edited gradient angle survives reload");
        const auto* reloadPoint2 = reloadSeries.dataPoint(2);
        const auto* reloadPoint4 = reloadSeries.dataPoint(4);
        test.checkTrue(reloadPoint2 != nullptr && reloadPoint4 != nullptr, "Edited and appended dPt survive reload");
        if (reloadPoint2) {
            test.checkEqual(reloadPoint2->fill.pattern, std::string("diagCross"), "Edited dPt pattern survives reload");
            test.checkEqual(reloadPoint2->marker.symbol, std::string("diamond"), "Edited dPt marker survives reload");
        }
        if (reloadPoint4)
            test.checkEqual(static_cast<int>(reloadPoint4->fill.kind), static_cast<int>(xlpp::ChartFillFormat::Kind::Gradient),
                            "Appended dPt gradient survives reload");
        const auto reloadLabel = std::find_if(reloadSeries.dataLabels().points.begin(), reloadSeries.dataLabels().points.end(),
                                              [](const auto& point) { return point.index == 1; });
        test.checkTrue(reloadLabel != reloadSeries.dataLabels().points.end() && reloadLabel->richText.plainText() == "Edited point",
                       "Edited rich point-label survives reload");

        auto* mutableSheet = reloaded.worksheet("P0L");
        const auto reloadStableId = reloadChart.stableId();
        test.checkTrue(mutableSheet->removeChartSeriesDataPointFormat(reloadStableId, 0, 2), "Imported dPt selectively removes");
        const auto removedOutput = std::filesystem::temp_directory_path() / "xlpp_p0m_chart_features_removed.xlsx";
        reloaded.save(removedOutput);
        xlpp::Workbook removedReload;
        removedReload.load(removedOutput);
        const auto* removedSheet = removedReload.worksheet("P0L");
        test.checkTrue(removedSheet != nullptr, "dPt removal output reloads");
        if (removedSheet) {
            const auto& removedSeries = static_cast<const xlpp::Worksheet&>(*removedSheet).charts().front().series().front();
            test.checkTrue(removedSeries.dataPoint(2) == nullptr, "Removed dPt absent after reload");
            test.checkTrue(removedSeries.dataPoint(4) != nullptr, "Sibling dPt preserved after removal");
        }
        std::filesystem::remove(removedOutput);
    }
    std::filesystem::remove(output);
}

void testImportedChartLayoutAxisLegendFormatting(TestContext& test) {
    const auto source = std::filesystem::path(XLPP_TEST_SOURCE_DIR) / "fixtures/openpyxl/chart_layout_axis_legend_formatting.xlsx";
    const auto before = xlpp::internal::ZipArchive::open(source);
    const auto originalDrawingRels = before.get("xl/drawings/_rels/drawing1.xml.rels");
    const auto originalImage = before.get("xl/media/image1.png");

    xlpp::Workbook workbook;
    workbook.load(source);
    auto* sheet = workbook.worksheet("P0L");
    test.checkTrue(sheet != nullptr, "P0N layout/axis/legend fixture loads");
    if (!sheet) return;
    const auto& charts = static_cast<const xlpp::Worksheet&>(*sheet).charts();
    test.checkEqual(charts.size(), std::size_t{1}, "P0N fixture exposes one chart");
    if (charts.empty()) return;
    const auto& chart = charts.front();
    test.checkTrue(chart.plotAreaLayout().present, "Plot-area manual layout parsed");
    test.checkEqual(chart.plotAreaLayout().target, std::string("inner"), "Plot-area layout target parsed");
    test.checkNear(chart.plotAreaLayout().x, 0.1, 1e-12, "Plot-area layout X parsed");
    test.checkNear(chart.plotAreaLayout().width, 0.72, 1e-12, "Plot-area layout width parsed");

    const auto* axis = chart.axisById(10);
    test.checkTrue(axis != nullptr, "P0N primary axis located by native axId");
    if (!axis) return;
    test.checkEqual(axis->numberFormat, std::string("0.00"), "Axis number format parsed");
    test.checkTrue(!axis->numberFormatSourceLinked, "Axis sourceLinked=false parsed");
    test.checkEqual(axis->majorTickMark, std::string("out"), "Axis major tick parsed");
    test.checkEqual(axis->minorTickMark, std::string("in"), "Axis minor tick parsed");
    test.checkEqual(axis->tickLabelPosition, std::string("low"), "Axis tick-label position parsed");
    test.checkTrue(axis->hasMajorUnit && axis->hasMinorUnit, "Axis major/minor units detected");
    test.checkNear(axis->majorUnit, 2.5, 1e-12, "Axis major unit parsed");
    test.checkNear(axis->minorUnit, 0.5, 1e-12, "Axis minor unit parsed");
    test.checkEqual(axis->crosses, std::string("max"), "Axis crosses parsed");
    test.checkEqual(axis->crossBetween, std::string("midCat"), "Axis crossBetween parsed");
    test.checkTrue(axis->titleRichText.present && axis->titleRichText.runs.size() == 2, "Axis rich title parsed");
    test.checkEqual(axis->titleRichText.plainText(), std::string("Input axis"), "Axis rich title text concatenated");
    test.checkTrue(axis->lineFormat.present, "Axis line formatting parsed");
    test.checkEqual(axis->lineFormat.color.value, std::string("accent1"), "Axis line scheme color parsed");
    test.checkEqual(axis->majorGridlineFormat.color.value, std::string("D0D0D0"), "Major gridline format parsed");
    test.checkEqual(axis->minorGridlineFormat.color.value, std::string("accent3"), "Minor gridline format parsed");

    test.checkTrue(chart.legendFormat().present, "Legend formatting model populated");
    test.checkTrue(chart.legendFormat().overlay, "Legend overlay parsed");
    test.checkEqual(chart.legendPosition(), std::string("b"), "Legend position parsed");
    test.checkTrue(chart.legendFormat().layout.present, "Legend manual layout parsed");
    test.checkEqual(chart.legendFormat().layout.target, std::string("outer"), "Legend layout target parsed");
    test.checkEqual(chart.legendFormat().fill.color.value, std::string("accent6"), "Legend fill parsed");
    test.checkEqual(chart.legendFormat().line.color.value, std::string("222222"), "Legend line parsed");

    const auto stableId = chart.stableId();
    xlpp::ChartRichText richTitle; richTitle.present = true;
    xlpp::ChartTextRun axisRun; axisRun.text = "XL++ Axis"; axisRun.bold = true; axisRun.fontSizePoints = 13.0; axisRun.color = {xlpp::ChartColor::Kind::SRgb, "123456"};
    richTitle.runs.push_back(axisRun);
    test.checkTrue(sheet->setChartAxisTitleRichText(stableId, 10, richTitle), "Axis rich title selectively edits");
    test.checkTrue(sheet->setChartAxisNumberFormat(stableId, 10, "0.0000", false), "Axis number format selectively edits");
    test.checkTrue(sheet->setChartAxisTicks(stableId, 10, "cross", "none", "high"), "Axis tick settings selectively edit");
    test.checkTrue(sheet->setChartAxisUnits(stableId, 10, 5.0, 1.0), "Axis units selectively edit");
    test.checkTrue(sheet->setChartAxisCrossing(stableId, 10, "autoZero", "between"), "Axis crossing selectively edits");

    xlpp::ChartLineFormat axisLine; axisLine.present = true; axisLine.color = {xlpp::ChartColor::Kind::SRgb, "880000"}; axisLine.widthPoints = 2.0; axisLine.dash = "solid";
    test.checkTrue(sheet->setChartAxisLineFormat(stableId, 10, axisLine), "Axis line selectively edits");
    xlpp::ChartLineFormat majorGrid; majorGrid.present = true; majorGrid.color = {xlpp::ChartColor::Kind::SRgb, "00AA00"}; majorGrid.widthPoints = 1.25; majorGrid.dash = "dash";
    test.checkTrue(sheet->setChartAxisGridlineFormat(stableId, 10, true, majorGrid), "Major gridline selectively edits");
    xlpp::ChartLineFormat minorGrid; minorGrid.present = true; minorGrid.noFill = true;
    test.checkTrue(sheet->setChartAxisGridlineFormat(stableId, 10, false, minorGrid), "Minor gridline selectively edits");

    xlpp::ChartManualLayout plotLayout; plotLayout.present = true; plotLayout.target = "inner"; plotLayout.xMode = plotLayout.yMode = plotLayout.widthMode = plotLayout.heightMode = "factor"; plotLayout.hasX = plotLayout.hasY = plotLayout.hasWidth = plotLayout.hasHeight = true; plotLayout.x = 0.15; plotLayout.y = 0.16; plotLayout.width = 0.68; plotLayout.height = 0.62;
    test.checkTrue(sheet->setChartPlotAreaLayout(stableId, plotLayout), "Plot-area manual layout selectively edits");
    xlpp::ChartManualLayout legendLayout; legendLayout.present = true; legendLayout.target = "outer"; legendLayout.xMode = legendLayout.yMode = legendLayout.widthMode = legendLayout.heightMode = "factor"; legendLayout.hasX = legendLayout.hasY = legendLayout.hasWidth = legendLayout.hasHeight = true; legendLayout.x = 0.25; legendLayout.y = 0.8; legendLayout.width = 0.5; legendLayout.height = 0.14;
    test.checkTrue(sheet->setChartLegend(stableId, true, "tr"), "Legend position selectively edits");
    test.checkTrue(sheet->setChartLegendLayout(stableId, legendLayout), "Legend layout selectively edits");
    test.checkTrue(sheet->setChartLegendOverlay(stableId, false), "Legend overlay selectively edits");
    xlpp::ChartLineFormat legendLine; legendLine.present = true; legendLine.color = {xlpp::ChartColor::Kind::SRgb, "000088"}; legendLine.widthPoints = 1.5;
    test.checkTrue(sheet->setChartLegendLineFormat(stableId, legendLine), "Legend line selectively edits");
    xlpp::ChartFillFormat legendFill; legendFill.present = true; legendFill.kind = xlpp::ChartFillFormat::Kind::Solid; legendFill.color = {xlpp::ChartColor::Kind::SRgb, "EEEEEE"};
    test.checkTrue(sheet->setChartLegendFillFormat(stableId, legendFill), "Legend fill selectively edits");

    sheet->cell("K22").setValue("p0n-regression");
    const auto output = std::filesystem::temp_directory_path() / "xlpp_p0n_chart_layout_axis_legend.xlsx";
    workbook.save(output);
    const auto after = xlpp::internal::ZipArchive::open(output);
    const auto xml = after.get("xl/charts/chart1.xml");
    test.checkTrue(xml.find("0.0000") != std::string::npos && xml.find("majorUnit val=\"5") != std::string::npos, "Axis format edits written");
    test.checkTrue(xml.find("XL++ Axis") != std::string::npos && xml.find("123456") != std::string::npos, "Axis rich title written");
    test.checkTrue(xml.find("880000") != std::string::npos && xml.find("00AA00") != std::string::npos, "Axis/gridline formatting written");
    test.checkTrue(xml.find("x val=\"0.15\"") != std::string::npos && xml.find("x val=\"0.25\"") != std::string::npos, "Plot and legend manual layouts written");
    test.checkTrue(xml.find("000088") != std::string::npos && xml.find("EEEEEE") != std::string::npos, "Legend formatting written");
    test.checkEqual(after.get("xl/drawings/_rels/drawing1.xml.rels"), originalDrawingRels, "P0N keeps drawing relationships byte-identical");
    test.checkEqual(after.get("xl/media/image1.png"), originalImage, "P0N keeps sibling image byte-identical");
    test.checkTrue(xlpp::internal::RelationshipGraph::fromArchive(after).validate().ok(), "P0N output package graph validates");

    xlpp::Workbook reloaded; reloaded.load(output);
    const auto* reloadSheet = reloaded.worksheet("P0L");
    test.checkTrue(reloadSheet != nullptr, "P0N output reloads");
    if (reloadSheet) {
        const auto& reloadChart = static_cast<const xlpp::Worksheet&>(*reloadSheet).charts().front();
        const auto* reloadAxis = reloadChart.axisById(10);
        test.checkTrue(reloadAxis != nullptr, "Edited axis remains addressable by axId");
        if (reloadAxis) {
            test.checkEqual(reloadAxis->titleRichText.plainText(), std::string("XL++ Axis"), "Axis rich title survives reload");
            test.checkEqual(reloadAxis->numberFormat, std::string("0.0000"), "Axis number format survives reload");
            test.checkEqual(reloadAxis->majorTickMark, std::string("cross"), "Axis major tick survives reload");
            test.checkEqual(reloadAxis->tickLabelPosition, std::string("high"), "Axis tick-label position survives reload");
            test.checkNear(reloadAxis->majorUnit, 5.0, 1e-12, "Axis major unit survives reload");
            test.checkEqual(reloadAxis->crossBetween, std::string("between"), "Axis crossBetween survives reload");
            test.checkEqual(reloadAxis->lineFormat.color.value, std::string("880000"), "Axis line formatting survives reload");
            test.checkEqual(reloadAxis->majorGridlineFormat.color.value, std::string("00AA00"), "Major gridline formatting survives reload");
            test.checkTrue(reloadAxis->minorGridlineFormat.noFill, "Minor gridline no-fill survives reload");
        }
        test.checkNear(reloadChart.plotAreaLayout().x, 0.15, 1e-12, "Plot layout survives reload");
        test.checkEqual(reloadChart.legendPosition(), std::string("tr"), "Legend position survives reload");
        test.checkTrue(!reloadChart.legendFormat().overlay, "Legend overlay survives reload");
        test.checkNear(reloadChart.legendFormat().layout.x, 0.25, 1e-12, "Legend layout survives reload");
        test.checkEqual(reloadChart.legendFormat().line.color.value, std::string("000088"), "Legend line survives reload");
        test.checkEqual(reloadChart.legendFormat().fill.color.value, std::string("EEEEEE"), "Legend fill survives reload");
    }
    std::filesystem::remove(output);
}

void testImportedChartAxisScalingDisplayUnitsAndAreaFormatting(TestContext& test) {
    const auto source = std::filesystem::path(XLPP_TEST_SOURCE_DIR) / "fixtures/openpyxl/chart_axis_scaling_display_units_area_formatting.xlsx";
    const auto before = xlpp::internal::ZipArchive::open(source);
    const auto originalDrawingRels = before.get("xl/drawings/_rels/drawing1.xml.rels");
    const auto originalImage = before.get("xl/media/image1.png");

    xlpp::Workbook workbook;
    workbook.load(source);
    auto* sheet = workbook.worksheet("P0L");
    test.checkTrue(sheet != nullptr, "P0O scaling/display-units fixture loads");
    if (!sheet) return;
    const auto& charts = static_cast<const xlpp::Worksheet&>(*sheet).charts();
    test.checkEqual(charts.size(), std::size_t{1}, "P0O fixture exposes one chart");
    if (charts.empty()) return;
    const auto& chart = charts.front();
    const auto* xAxis = chart.axisById(10);
    const auto* yAxis = chart.axisById(20);
    test.checkTrue(xAxis != nullptr && yAxis != nullptr, "P0O native axes located");
    if (!xAxis || !yAxis) return;

    test.checkTrue(xAxis->scaling.hasMinimum && xAxis->scaling.hasMaximum && xAxis->scaling.hasLogBase,
                   "Axis scaling min/max/log flags parsed");
    test.checkNear(xAxis->scaling.minimum, 1.0, 1e-12, "Axis scaling minimum parsed");
    test.checkNear(xAxis->scaling.maximum, 100.0, 1e-12, "Axis scaling maximum parsed");
    test.checkNear(xAxis->scaling.logBase, 10.0, 1e-12, "Axis log base parsed");
    test.checkTrue(xAxis->scaling.reverseOrder, "Axis reverse order parsed");
    test.checkTrue(xAxis->hasCrossesAt, "Axis crossesAt presence parsed");
    test.checkNear(xAxis->crossesAt, 2.0, 1e-12, "Axis crossesAt value parsed");
    test.checkTrue(xAxis->hasMajorGridlines && xAxis->hasMinorGridlines, "Axis gridline lifecycle state parsed");

    test.checkTrue(yAxis->displayUnits.present, "Display units parsed");
    test.checkEqual(yAxis->displayUnits.builtInUnit, std::string("thousands"), "Built-in display units parsed");
    test.checkTrue(yAxis->displayUnits.showLabel, "Display-units label presence parsed");
    test.checkEqual(yAxis->displayUnits.labelRichText.plainText(), std::string("Thousands"), "Display-units rich label parsed");

    test.checkEqual(chart.chartAreaFillFormat().color.value, std::string("accent1"), "Chart-area fill parsed");
    test.checkEqual(chart.chartAreaLineFormat().color.value, std::string("336699"), "Chart-area line parsed");
    test.checkEqual(chart.plotAreaFillFormat().color.value, std::string("FFF2CC"), "Plot-area fill parsed");
    test.checkEqual(chart.plotAreaLineFormat().color.value, std::string("CC9900"), "Plot-area line parsed");

    const auto stableId = chart.stableId();
    xlpp::ChartAxisScaling invalidScaling; invalidScaling.hasMinimum=true; invalidScaling.minimum=10.0; invalidScaling.hasMaximum=true; invalidScaling.maximum=1.0;
    test.checkTrue(!sheet->setChartAxisScaling(stableId, 10, invalidScaling), "Invalid axis scaling range rejected");
    xlpp::ChartAxisScaling scaling; scaling.hasMinimum=true; scaling.minimum=0.5; scaling.hasMaximum=true; scaling.maximum=500.0; scaling.hasLogBase=true; scaling.logBase=10.0; scaling.reverseOrder=false;
    test.checkTrue(sheet->setChartAxisScaling(stableId, 10, scaling), "Axis scaling selectively edits");
    test.checkTrue(sheet->setChartAxisCrossesAt(stableId, 10, 5.5), "Axis crossesAt selectively edits");

    xlpp::ChartDisplayUnits invalidUnits; invalidUnits.present=true; invalidUnits.builtInUnit="invalid";
    test.checkTrue(!sheet->setChartAxisDisplayUnits(stableId, 20, invalidUnits), "Invalid display unit rejected");
    xlpp::ChartDisplayUnits units; units.present=true; units.hasCustomUnit=true; units.customUnit=1000000.0; units.showLabel=true;
    units.labelRichText.present=true; xlpp::ChartTextRun unitRun; unitRun.text="Millions"; unitRun.bold=true; unitRun.color={xlpp::ChartColor::Kind::SRgb,"7030A0"}; units.labelRichText.runs.push_back(unitRun);
    test.checkTrue(sheet->setChartAxisDisplayUnits(stableId, 20, units), "Custom display units selectively edit");
    test.checkTrue(sheet->removeChartAxisGridlines(stableId, 10, false), "Minor gridlines selectively remove");

    xlpp::ChartLineFormat chartLine; chartLine.present=true; chartLine.color={xlpp::ChartColor::Kind::SRgb,"112233"}; chartLine.widthPoints=2.25; chartLine.dash="solid";
    test.checkTrue(sheet->setChartAreaLineFormat(stableId, chartLine), "Chart-area line selectively edits");
    xlpp::ChartFillFormat chartFill; chartFill.present=true; chartFill.kind=xlpp::ChartFillFormat::Kind::Solid; chartFill.color={xlpp::ChartColor::Kind::SRgb,"F0F0F0"};
    test.checkTrue(sheet->setChartAreaFillFormat(stableId, chartFill), "Chart-area fill selectively edits");
    xlpp::ChartLineFormat plotLine; plotLine.present=true; plotLine.color={xlpp::ChartColor::Kind::SRgb,"445566"}; plotLine.widthPoints=1.75; plotLine.dash="dash";
    test.checkTrue(sheet->setChartPlotAreaLineFormat(stableId, plotLine), "Plot-area line selectively edits");
    xlpp::ChartFillFormat plotFill; plotFill.present=true; plotFill.kind=xlpp::ChartFillFormat::Kind::Solid; plotFill.color={xlpp::ChartColor::Kind::SRgb,"E2F0D9"};
    test.checkTrue(sheet->setChartPlotAreaFillFormat(stableId, plotFill), "Plot-area fill selectively edits");

    sheet->cell("K23").setValue("p0o-regression");
    const auto output = std::filesystem::temp_directory_path() / "xlpp_p0o_axis_scaling_display_units.xlsx";
    workbook.save(output);
    const auto after = xlpp::internal::ZipArchive::open(output);
    const auto xml = after.get("xl/charts/chart1.xml");
    test.checkTrue(xml.find("logBase val=\"10\"") != std::string::npos && xml.find("max val=\"500\"") != std::string::npos && xml.find("min val=\"0.5\"") != std::string::npos,
                   "Axis scaling edits written");
    test.checkTrue(xml.find("orientation val=\"minMax\"") != std::string::npos && xml.find("crossesAt val=\"5.5\"") != std::string::npos,
                   "Axis orientation and crossesAt edits written");
    test.checkTrue(xml.find("custUnit val=\"1000000\"") != std::string::npos && xml.find("Millions") != std::string::npos,
                   "Custom display units and label written");
    test.checkTrue(xml.find("112233") != std::string::npos && xml.find("F0F0F0") != std::string::npos && xml.find("445566") != std::string::npos && xml.find("E2F0D9") != std::string::npos,
                   "Chart/plot area formatting written");
    test.checkEqual(after.get("xl/drawings/_rels/drawing1.xml.rels"), originalDrawingRels, "P0O keeps drawing relationships byte-identical");
    test.checkEqual(after.get("xl/media/image1.png"), originalImage, "P0O keeps sibling image byte-identical");
    test.checkTrue(xlpp::internal::RelationshipGraph::fromArchive(after).validate().ok(), "P0O output package graph validates");

    xlpp::Workbook reloaded; reloaded.load(output);
    auto* reloadSheet = reloaded.worksheet("P0L");
    test.checkTrue(reloadSheet != nullptr, "P0O output reloads");
    if (reloadSheet) {
        const auto& reloadChart = static_cast<const xlpp::Worksheet&>(*reloadSheet).charts().front();
        const auto* reloadX = reloadChart.axisById(10); const auto* reloadY = reloadChart.axisById(20);
        test.checkTrue(reloadX != nullptr && reloadY != nullptr, "Edited P0O axes remain addressable");
        if (reloadX && reloadY) {
            test.checkNear(reloadX->scaling.minimum, 0.5, 1e-12, "Edited scaling minimum survives reload");
            test.checkNear(reloadX->scaling.maximum, 500.0, 1e-12, "Edited scaling maximum survives reload");
            test.checkTrue(!reloadX->scaling.reverseOrder, "Edited axis orientation survives reload");
            test.checkNear(reloadX->crossesAt, 5.5, 1e-12, "Edited crossesAt survives reload");
            test.checkTrue(!reloadX->hasMinorGridlines, "Removed minor gridlines stay absent after reload");
            test.checkTrue(reloadY->displayUnits.present && reloadY->displayUnits.hasCustomUnit, "Custom display units survive reload");
            test.checkNear(reloadY->displayUnits.customUnit, 1000000.0, 1e-6, "Custom display-unit value survives reload");
            test.checkEqual(reloadY->displayUnits.labelRichText.plainText(), std::string("Millions"), "Display-unit rich label survives reload");
        }
        test.checkEqual(reloadChart.chartAreaLineFormat().color.value, std::string("112233"), "Chart-area line survives reload");
        test.checkEqual(reloadChart.chartAreaFillFormat().color.value, std::string("F0F0F0"), "Chart-area fill survives reload");
        test.checkEqual(reloadChart.plotAreaLineFormat().color.value, std::string("445566"), "Plot-area line survives reload");
        test.checkEqual(reloadChart.plotAreaFillFormat().color.value, std::string("E2F0D9"), "Plot-area fill survives reload");

        const auto reloadStableId = reloadChart.stableId();
        test.checkTrue(reloadSheet->clearChartAxisCrossesAt(reloadStableId, 10), "crossesAt selectively clears");
        test.checkTrue(reloadSheet->clearChartAxisDisplayUnits(reloadStableId, 20), "display units selectively clear");
        test.checkTrue(reloadSheet->removeChartAxisGridlines(reloadStableId, 10, true), "major gridlines selectively remove");
        const auto cleared = std::filesystem::temp_directory_path() / "xlpp_p0o_axis_lifecycle_cleared.xlsx";
        reloaded.save(cleared);
        xlpp::Workbook clearedReload; clearedReload.load(cleared);
        const auto* clearedSheet = clearedReload.worksheet("P0L");
        test.checkTrue(clearedSheet != nullptr, "P0O lifecycle-clear output reloads");
        if (clearedSheet) {
            const auto& clearedChart = static_cast<const xlpp::Worksheet&>(*clearedSheet).charts().front();
            const auto* clearedX = clearedChart.axisById(10); const auto* clearedY = clearedChart.axisById(20);
            test.checkTrue(clearedX && !clearedX->hasCrossesAt, "crossesAt absent after clear/reload");
            test.checkTrue(clearedX && !clearedX->hasMajorGridlines && !clearedX->hasMinorGridlines, "Both gridline collections absent after lifecycle removal");
            test.checkTrue(clearedY && !clearedY->displayUnits.present, "Display units absent after clear/reload");
        }
        std::filesystem::remove(cleared);
    }
    std::filesystem::remove(output);
}

void testImportedChartAuxiliaryObjects(TestContext& test) {
    const auto source = std::filesystem::path(XLPP_TEST_SOURCE_DIR) / "fixtures/openpyxl/chart_auxiliary_objects.xlsx";
    const auto before = xlpp::internal::ZipArchive::open(source);
    const auto originalDrawingRels = before.get("xl/drawings/_rels/drawing1.xml.rels");
    const auto originalImage = before.get("xl/media/image1.png");

    xlpp::Workbook workbook;
    workbook.load(source);
    auto* sheet = workbook.worksheet("Aux");
    test.checkTrue(sheet != nullptr, "P0P auxiliary-object fixture loads");
    if (!sheet) return;
    const auto& charts = static_cast<const xlpp::Worksheet&>(*sheet).charts();
    test.checkEqual(charts.size(), std::size_t{1}, "P0P fixture exposes one chart");
    if (charts.empty()) return;
    const auto& chart = charts.front();
    test.checkEqual(chart.plots().size(), std::size_t{1}, "P0P fixture exposes one plot");
    if (chart.plots().empty()) return;
    const auto& plot = chart.plots().front();

    test.checkTrue(plot.hasDropLines, "Drop lines parsed");
    test.checkEqual(plot.dropLinesFormat.color.value, std::string("FF0000"), "Drop-line color parsed");
    test.checkNear(plot.dropLinesFormat.widthPoints, 1.5, 1e-9, "Drop-line width parsed");
    test.checkTrue(plot.hasHighLowLines, "High-low lines parsed");
    test.checkEqual(plot.highLowLinesFormat.color.value, std::string("00AA00"), "High-low line color parsed");
    test.checkTrue(plot.upDownBars.present, "Up/down bars parsed");
    test.checkEqual(plot.upDownBars.gapWidth, 120, "Up/down gap width parsed");
    test.checkEqual(plot.upDownBars.upFill.color.value, std::string("DDEBF7"), "Up-bar fill parsed");
    test.checkEqual(plot.upDownBars.upLine.color.value, std::string("4472C4"), "Up-bar line parsed");
    test.checkEqual(plot.upDownBars.downFill.color.value, std::string("FCE4D6"), "Down-bar fill parsed");
    test.checkEqual(plot.upDownBars.downLine.color.value, std::string("C00000"), "Down-bar line parsed");

    test.checkTrue(plot.dataLabels.showLeaderLines && plot.dataLabels.hasLeaderLines, "Plot leader lines parsed");
    test.checkEqual(plot.dataLabels.leaderLineFormat.color.value, std::string("7030A0"), "Leader-line color parsed");
    test.checkEqual(plot.dataLabels.leaderLineFormat.dash, std::string("dash"), "Leader-line dash parsed");

    test.checkTrue(chart.dataTable().present, "Chart data table parsed");
    test.checkTrue(chart.dataTable().showHorizontalBorder, "Data-table horizontal border flag parsed");
    test.checkTrue(!chart.dataTable().showVerticalBorder, "Data-table vertical border flag parsed");
    test.checkTrue(chart.dataTable().showOutline && chart.dataTable().showLegendKeys, "Data-table outline/keys flags parsed");
    test.checkEqual(chart.dataTable().fill.color.value, std::string("FFF2CC"), "Data-table fill parsed");
    test.checkEqual(chart.dataTable().line.color.value, std::string("7F6000"), "Data-table line parsed");

    const auto stableId = chart.stableId();
    xlpp::ChartDataTable table;
    table.present = true; table.showHorizontalBorder = false; table.showVerticalBorder = true; table.showOutline = true; table.showLegendKeys = false;
    table.fill.present = true; table.fill.kind = xlpp::ChartFillFormat::Kind::Solid; table.fill.color = {xlpp::ChartColor::Kind::SRgb, "C6E0B4"};
    table.line.present = true; table.line.color = {xlpp::ChartColor::Kind::SRgb, "548235"}; table.line.widthPoints = 1.75; table.line.dash = "dash";
    test.checkTrue(sheet->setChartDataTable(stableId, table), "Data table selectively edits");

    xlpp::ChartLineFormat drop; drop.present = true; drop.color = {xlpp::ChartColor::Kind::SRgb, "112233"}; drop.widthPoints = 2.5; drop.dash = "dash";
    test.checkTrue(sheet->setChartPlotDropLines(stableId, 0, drop), "Drop lines selectively edit");
    test.checkTrue(sheet->removeChartPlotHighLowLines(stableId, 0), "High-low lines selectively remove");

    xlpp::ChartUpDownBars bars; bars.present = true; bars.gapWidth = 80;
    bars.upFill.present = true; bars.upFill.kind = xlpp::ChartFillFormat::Kind::Solid; bars.upFill.color = {xlpp::ChartColor::Kind::SRgb, "E2F0D9"};
    bars.upLine.present = true; bars.upLine.color = {xlpp::ChartColor::Kind::SRgb, "70AD47"}; bars.upLine.widthPoints = 1.25;
    bars.downFill.present = true; bars.downFill.kind = xlpp::ChartFillFormat::Kind::Solid; bars.downFill.color = {xlpp::ChartColor::Kind::SRgb, "F4B183"};
    bars.downLine.present = true; bars.downLine.color = {xlpp::ChartColor::Kind::SRgb, "C65911"}; bars.downLine.widthPoints = 1.25;
    test.checkTrue(sheet->setChartPlotUpDownBars(stableId, 0, bars), "Up/down bars selectively edit");
    xlpp::ChartUpDownBars invalidBars; invalidBars.gapWidth = 501;
    test.checkTrue(!sheet->setChartPlotUpDownBars(stableId, 0, invalidBars), "Invalid up/down gap width rejected");

    xlpp::ChartLineFormat plotLeader; plotLeader.present = true; plotLeader.color = {xlpp::ChartColor::Kind::SRgb, "44546A"}; plotLeader.widthPoints = 1.5; plotLeader.dash = "dot";
    test.checkTrue(sheet->setChartPlotLeaderLineFormat(stableId, 0, plotLeader), "Plot leader-line formatting selectively edits");
    xlpp::ChartLineFormat seriesLeader; seriesLeader.present = true; seriesLeader.color = {xlpp::ChartColor::Kind::SRgb, "A5A5A5"}; seriesLeader.widthPoints = 1.0; seriesLeader.dash = "solid";
    test.checkTrue(sheet->setChartSeriesLeaderLineFormat(stableId, 0, seriesLeader), "Series leader lines selectively add");

    sheet->cell("K24").setValue("p0p-regression");
    const auto output = std::filesystem::temp_directory_path() / "xlpp_p0p_chart_auxiliary_objects.xlsx";
    workbook.save(output);
    const auto after = xlpp::internal::ZipArchive::open(output);
    const auto xml = after.get("xl/charts/chart1.xml");
    test.checkTrue(xml.find("showHorzBorder val=\"0\"") != std::string::npos && xml.find("showVertBorder val=\"1\"") != std::string::npos,
                   "Edited data-table border flags written");
    test.checkTrue(xml.find("C6E0B4") != std::string::npos && xml.find("548235") != std::string::npos, "Edited data-table formatting written");
    test.checkTrue(xml.find("112233") != std::string::npos, "Edited drop-line formatting written");
    test.checkTrue(xml.find("hiLowLines") == std::string::npos, "Removed high-low lines absent from XML");
    test.checkTrue(xml.find("gapWidth val=\"80\"") != std::string::npos && xml.find("E2F0D9") != std::string::npos && xml.find("F4B183") != std::string::npos,
                   "Edited up/down bars written");
    test.checkTrue(xml.find("44546A") != std::string::npos && xml.find("A5A5A5") != std::string::npos, "Plot and series leader-line formatting written");
    test.checkEqual(after.get("xl/drawings/_rels/drawing1.xml.rels"), originalDrawingRels, "P0P keeps drawing relationships byte-identical");
    test.checkEqual(after.get("xl/media/image1.png"), originalImage, "P0P keeps sibling image byte-identical");
    test.checkTrue(xlpp::internal::RelationshipGraph::fromArchive(after).validate().ok(), "P0P output package graph validates");

    xlpp::Workbook reloaded; reloaded.load(output);
    auto* reloadSheet = reloaded.worksheet("Aux");
    test.checkTrue(reloadSheet != nullptr, "P0P output reloads");
    if (reloadSheet) {
        const auto& reloadChart = static_cast<const xlpp::Worksheet&>(*reloadSheet).charts().front();
        const auto& reloadPlot = reloadChart.plots().front();
        test.checkTrue(reloadChart.dataTable().present && !reloadChart.dataTable().showHorizontalBorder && reloadChart.dataTable().showVerticalBorder,
                       "Edited data table survives reload");
        test.checkEqual(reloadChart.dataTable().fill.color.value, std::string("C6E0B4"), "Edited data-table fill survives reload");
        test.checkTrue(reloadPlot.hasDropLines && !reloadPlot.hasHighLowLines, "Drop/high-low lifecycle survives reload");
        test.checkEqual(reloadPlot.dropLinesFormat.color.value, std::string("112233"), "Edited drop-line color survives reload");
        test.checkTrue(reloadPlot.upDownBars.present && reloadPlot.upDownBars.gapWidth == 80, "Edited up/down bars survive reload");
        test.checkEqual(reloadPlot.upDownBars.downLine.color.value, std::string("C65911"), "Down-bar formatting survives reload");
        test.checkTrue(reloadPlot.dataLabels.hasLeaderLines, "Plot leader lines survive reload");
        test.checkEqual(reloadPlot.dataLabels.leaderLineFormat.color.value, std::string("44546A"), "Plot leader-line formatting survives reload");
        test.checkTrue(reloadChart.series()[0].dataLabels().hasLeaderLines, "Series leader lines survive reload");
        test.checkEqual(reloadChart.series()[0].dataLabels().leaderLineFormat.color.value, std::string("A5A5A5"), "Series leader-line formatting survives reload");

        const auto reloadStableId = reloadChart.stableId();
        test.checkTrue(reloadSheet->removeChartDataTable(reloadStableId), "Data table selectively removes");
        test.checkTrue(reloadSheet->removeChartPlotDropLines(reloadStableId, 0), "Drop lines selectively remove");
        test.checkTrue(reloadSheet->removeChartPlotUpDownBars(reloadStableId, 0), "Up/down bars selectively remove");
        test.checkTrue(reloadSheet->removeChartPlotLeaderLines(reloadStableId, 0), "Plot leader lines selectively remove");
        test.checkTrue(reloadSheet->removeChartSeriesLeaderLines(reloadStableId, 0), "Series leader lines selectively remove");
        xlpp::ChartLineFormat high; high.present=true; high.color={xlpp::ChartColor::Kind::SRgb,"00B0F0"}; high.widthPoints=2.0; high.dash="solid";
        test.checkTrue(reloadSheet->setChartPlotHighLowLines(reloadStableId, 0, high), "High-low lines selectively add after removal");
        const auto lifecycle = std::filesystem::temp_directory_path() / "xlpp_p0p_chart_auxiliary_lifecycle.xlsx";
        reloaded.save(lifecycle);
        xlpp::Workbook lifecycleReload; lifecycleReload.load(lifecycle);
        const auto* lifecycleSheet = lifecycleReload.worksheet("Aux");
        test.checkTrue(lifecycleSheet != nullptr, "P0P lifecycle output reloads");
        if (lifecycleSheet) {
            const auto& lifecycleChart = static_cast<const xlpp::Worksheet&>(*lifecycleSheet).charts().front();
            const auto& lifecyclePlot = lifecycleChart.plots().front();
            test.checkTrue(!lifecycleChart.dataTable().present, "Data table absent after lifecycle removal");
            test.checkTrue(!lifecyclePlot.hasDropLines && !lifecyclePlot.upDownBars.present, "Drop/up-down objects absent after lifecycle removal");
            test.checkTrue(lifecyclePlot.hasHighLowLines && lifecyclePlot.highLowLinesFormat.color.value == "00B0F0", "High-low object added with formatting");
            test.checkTrue(!lifecyclePlot.dataLabels.hasLeaderLines, "Plot leader-line container absent after lifecycle removal");
            test.checkTrue(!lifecycleChart.series()[0].dataLabels().hasLeaderLines, "Series leader-line container absent after lifecycle removal");
        }
        std::filesystem::remove(lifecycle);
    }
    std::filesystem::remove(output);
}

void testStockChartStructureGenerationAndDataTableText(TestContext& test) {
    const auto source = std::filesystem::path(XLPP_TEST_SOURCE_DIR) / "fixtures/openpyxl/stock_auxiliary_datatable_text.xlsx";
    const auto before = xlpp::internal::ZipArchive::open(source);
    const auto originalDrawingRels = before.get("xl/drawings/_rels/drawing1.xml.rels");
    const auto originalImage = before.get("xl/media/image1.png");

    xlpp::Workbook workbook; workbook.load(source);
    auto* sheet = workbook.worksheet("Stock");
    test.checkTrue(sheet != nullptr, "P0Q stock fixture loads");
    if (!sheet) return;
    const auto& charts = static_cast<const xlpp::Worksheet&>(*sheet).charts();
    test.checkEqual(charts.size(), std::size_t{1}, "P0Q stock fixture exposes one chart");
    if (charts.empty()) return;
    const auto& chart = charts.front();
    test.checkTrue(chart.type() == xlpp::Chart::Type::Stock, "stockChart maps to Chart::Type::Stock");
    test.checkEqual(chart.series().size(), std::size_t{4}, "Open-high-low-close stock series parsed");
    test.checkEqual(chart.plots().size(), std::size_t{1}, "Stock chart exposes one plot");
    if (chart.plots().empty()) return;
    const auto& plot = chart.plots().front();
    test.checkTrue(plot.type == xlpp::Chart::Type::Stock, "Stock plot type parsed");
    test.checkTrue(plot.hasHighLowLines, "Stock high-low lines parsed");
    test.checkEqual(plot.highLowLinesFormat.color.value, std::string("00AA00"), "Stock high-low line formatting parsed");
    test.checkTrue(plot.upDownBars.present && plot.upDownBars.gapWidth == 120, "Stock up/down bars parsed");
    test.checkEqual(plot.upDownBars.upFill.color.value, std::string("DDEBF7"), "Stock up-bar fill parsed");
    test.checkTrue(chart.dataTable().present, "Stock data table parsed");
    test.checkTrue(chart.dataTable().textStyle.present, "Data-table txPr text style parsed");
    test.checkTrue(chart.dataTable().textStyle.bold && chart.dataTable().textStyle.italic, "Data-table text bold/italic parsed");
    test.checkNear(chart.dataTable().textStyle.fontSizePoints, 11.0, 1e-9, "Data-table font size parsed");
    test.checkEqual(chart.dataTable().textStyle.typeface, std::string("Calibri"), "Data-table typeface parsed");
    test.checkEqual(chart.dataTable().textStyle.color.value, std::string("7030A0"), "Data-table text color parsed");

    const auto stableId = chart.stableId();
    xlpp::ChartDataTable edited = chart.dataTable();
    edited.showHorizontalBorder = false;
    edited.textStyle.present = true; edited.textStyle.bold = false; edited.textStyle.italic = true;
    edited.textStyle.fontSizePoints = 12.5; edited.textStyle.typeface = "Aptos";
    edited.textStyle.color = {xlpp::ChartColor::Kind::SRgb, "C00000"};
    test.checkTrue(sheet->setChartDataTable(stableId, edited), "Imported stock data-table text selectively edits");
    xlpp::ChartLineFormat hi; hi.present=true; hi.color={xlpp::ChartColor::Kind::SRgb,"00B0F0"}; hi.widthPoints=2.25; hi.dash="dash";
    test.checkTrue(sheet->setChartPlotHighLowLines(stableId, 0, hi), "Imported stock high-low lines selectively edit");
    sheet->cell("K24").setValue("p0q-regression");
    const auto output = std::filesystem::temp_directory_path() / "xlpp_p0q_stock_imported.xlsx";
    workbook.save(output);
    const auto after = xlpp::internal::ZipArchive::open(output);
    const auto xml = after.get("xl/charts/chart1.xml");
    test.checkTrue(xml.find("stockChart") != std::string::npos, "Selective stock edit preserves stockChart structure");
    test.checkTrue(xml.find("Aptos") != std::string::npos && xml.find("C00000") != std::string::npos, "Edited data-table txPr written");
    test.checkTrue(xml.find("00B0F0") != std::string::npos, "Edited stock high-low formatting written");
    test.checkEqual(after.get("xl/drawings/_rels/drawing1.xml.rels"), originalDrawingRels, "P0Q keeps drawing relationships byte-identical");
    test.checkEqual(after.get("xl/media/image1.png"), originalImage, "P0Q keeps sibling stock image byte-identical");
    test.checkTrue(xlpp::internal::RelationshipGraph::fromArchive(after).validate().ok(), "P0Q imported stock output package graph validates");

    xlpp::Workbook reloaded; reloaded.load(output);
    const auto* reloadSheet = reloaded.worksheet("Stock");
    test.checkTrue(reloadSheet != nullptr, "P0Q selectively edited stock workbook reloads");
    if (reloadSheet) {
        const auto& reloadChart = static_cast<const xlpp::Worksheet&>(*reloadSheet).charts().front();
        test.checkTrue(reloadChart.type() == xlpp::Chart::Type::Stock, "Stock type survives selective save/reload");
        test.checkEqual(reloadChart.dataTable().textStyle.typeface, std::string("Aptos"), "Edited dTable typeface survives reload");
        test.checkNear(reloadChart.dataTable().textStyle.fontSizePoints, 12.5, 1e-9, "Edited dTable font size survives reload");
        test.checkEqual(reloadChart.plots().front().highLowLinesFormat.color.value, std::string("00B0F0"), "Edited stock high-low lines survive reload");
    }
    std::filesystem::remove(output);

    // Generation path: P0P auxiliary model is now serialized for new charts.
    xlpp::Workbook generated;
    auto& generatedSheet = generated.addWorksheet("GeneratedStock");
    generatedSheet.cell("A1").setValue("Date"); generatedSheet.cell("B1").setValue("Open"); generatedSheet.cell("C1").setValue("High");
    generatedSheet.cell("D1").setValue("Low"); generatedSheet.cell("E1").setValue("Close");
    for (std::size_t r=2; r<=5; ++r) {
        generatedSheet.cell("A"+std::to_string(r)).setValue(static_cast<double>(r-1));
        generatedSheet.cell("B"+std::to_string(r)).setValue(10.0+r);
        generatedSheet.cell("C"+std::to_string(r)).setValue(13.0+r);
        generatedSheet.cell("D"+std::to_string(r)).setValue(8.0+r);
        generatedSheet.cell("E"+std::to_string(r)).setValue(12.0+r);
    }
    xlpp::Chart stock(xlpp::Chart::Type::Stock); stock.setTitle("Generated stock"); stock.setWidth(420); stock.setHeight(260);
    for (std::size_t col=0; col<4; ++col) {
        static const char* names[]{"Open","High","Low","Close"};
        static const char* letters[]{"B","C","D","E"};
        xlpp::ChartSeries series(names[col]);
        series.setCategoriesReference("'GeneratedStock'!$A$2:$A$5");
        series.setValuesReference(std::string("'GeneratedStock'!$") + letters[col] + "$2:$" + letters[col] + "$5");
        stock.addSeries(std::move(series));
    }
    auto& generatedPlot = stock.primaryPlot();
    generatedPlot.hasHighLowLines = true; generatedPlot.highLowLinesFormat = hi;
    generatedPlot.upDownBars.present = true; generatedPlot.upDownBars.gapWidth = 90;
    generatedPlot.upDownBars.upFill.present=true; generatedPlot.upDownBars.upFill.kind=xlpp::ChartFillFormat::Kind::Solid; generatedPlot.upDownBars.upFill.color={xlpp::ChartColor::Kind::SRgb,"D9EAD3"};
    generatedPlot.upDownBars.downFill.present=true; generatedPlot.upDownBars.downFill.kind=xlpp::ChartFillFormat::Kind::Solid; generatedPlot.upDownBars.downFill.color={xlpp::ChartColor::Kind::SRgb,"F4CCCC"};
    xlpp::ChartDataTable table; table.present=true; table.showHorizontalBorder=true; table.showVerticalBorder=true; table.showOutline=true; table.showLegendKeys=true;
    table.textStyle.present=true; table.textStyle.bold=true; table.textStyle.fontSizePoints=10.0; table.textStyle.typeface="Aptos"; table.textStyle.color={xlpp::ChartColor::Kind::SRgb,"1F4E78"};
    stock.setDataTable(table);
    generatedSheet.addChart(std::move(stock));
    const auto generatedPath = std::filesystem::temp_directory_path() / "xlpp_p0q_generated_stock.xlsx";
    generated.save(generatedPath);
    const auto generatedZip = xlpp::internal::ZipArchive::open(generatedPath);
    const auto generatedXml = generatedZip.get("xl/charts/chart1.xml");
    test.checkTrue(generatedXml.find("<c:stockChart>") != std::string::npos, "New Stock chart serializes stockChart");
    test.checkTrue(generatedXml.find("<c:hiLowLines>") != std::string::npos && generatedXml.find("<c:upDownBars>") != std::string::npos, "New stock auxiliary objects serialize");
    test.checkTrue(generatedXml.find("<c:dTable>") != std::string::npos && generatedXml.find("<c:txPr") != std::string::npos && generatedXml.find("Aptos") != std::string::npos, "New chart data-table txPr serializes");
    test.checkTrue(generatedXml.find("<c:cat><c:numRef>") != std::string::npos, "Generated stock categories use numeric references");
    test.checkTrue(xlpp::internal::RelationshipGraph::fromArchive(generatedZip).validate().ok(), "Generated stock package graph validates");
    xlpp::Workbook generatedReload; generatedReload.load(generatedPath);
    const auto* generatedReloadSheet = generatedReload.worksheet("GeneratedStock");
    test.checkTrue(generatedReloadSheet != nullptr, "Generated stock workbook reloads");
    if (generatedReloadSheet) {
        const auto& generatedChart = static_cast<const xlpp::Worksheet&>(*generatedReloadSheet).charts().front();
        test.checkTrue(generatedChart.type() == xlpp::Chart::Type::Stock, "Generated stock chart reloads as Stock");
        test.checkTrue(generatedChart.plots().front().hasHighLowLines && generatedChart.plots().front().upDownBars.present, "Generated stock auxiliary model reloads");
        test.checkTrue(generatedChart.dataTable().textStyle.present && generatedChart.dataTable().textStyle.bold, "Generated dTable text style reloads");
    }
    std::filesystem::remove(generatedPath);

    xlpp::Workbook invalid;
    auto& invalidSheet = invalid.addWorksheet("InvalidStock");
    xlpp::Chart invalidStock(xlpp::Chart::Type::Stock);
    invalidStock.addSeries(xlpp::ChartSeries("Only one")); invalidSheet.addChart(std::move(invalidStock));
    bool rejected=false; try { invalid.save(std::filesystem::temp_directory_path()/"xlpp_invalid_stock.xlsx"); } catch (const std::invalid_argument&) { rejected=true; }
    test.checkTrue(rejected, "Invalid stock series count rejected predictably");
}

void testThreeDSurfaceChartPreservationFoundation(TestContext& test) {
    const auto source = std::filesystem::path(XLPP_TEST_SOURCE_DIR) / "fixtures/openpyxl/chart_3d_surface.xlsx";
    const auto before = xlpp::internal::ZipArchive::open(source);
    const auto originalDrawingRels = before.get("xl/drawings/_rels/drawing1.xml.rels");
    const auto originalImage = before.get("xl/media/image1.png");
    std::vector<std::string> untouchedCharts;
    for (int index = 2; index <= 6; ++index)
        untouchedCharts.push_back(before.get("xl/charts/chart" + std::to_string(index) + ".xml"));

    xlpp::Workbook workbook; workbook.load(source);
    auto* sheet = workbook.worksheet("ThreeD");
    test.checkTrue(sheet != nullptr, "P0R 3D/surface fixture loads");
    if (!sheet) return;
    const auto& charts = static_cast<const xlpp::Worksheet&>(*sheet).charts();
    test.checkEqual(charts.size(), std::size_t{6}, "P0R fixture exposes six charts");
    if (charts.size() != 6) return;
    const std::array<xlpp::Chart::Type, 6> expected{{
        xlpp::Chart::Type::Bar3D, xlpp::Chart::Type::Line3D, xlpp::Chart::Type::Area3D,
        xlpp::Chart::Type::Pie3D, xlpp::Chart::Type::Surface, xlpp::Chart::Type::Surface3D}};
    for (std::size_t i=0; i<expected.size(); ++i)
        test.checkTrue(charts[i].type() == expected[i], "P0R chart type parsed at index " + std::to_string(i));

    const auto& bar3d = charts[0];
    test.checkTrue(bar3d.view3D().present, "Bar3D view3D parsed");
    test.checkTrue(bar3d.view3D().hasRotationX && bar3d.view3D().rotationX == 20, "Bar3D rotX parsed");
    test.checkTrue(bar3d.view3D().hasRotationY && bar3d.view3D().rotationY == 35, "Bar3D rotY parsed");
    test.checkTrue(bar3d.view3D().hasHeightPercent && bar3d.view3D().heightPercent == 120, "Bar3D hPercent parsed");
    test.checkTrue(bar3d.view3D().hasDepthPercent && bar3d.view3D().depthPercent == 180, "Bar3D depthPercent parsed");
    test.checkTrue(bar3d.view3D().hasRightAngleAxes && !bar3d.view3D().rightAngleAxes, "Bar3D rAngAx parsed");
    test.checkTrue(bar3d.view3D().hasPerspective && bar3d.view3D().perspective == 35, "Bar3D perspective parsed");
    test.checkTrue(bar3d.floorFormat().present && bar3d.floorFormat().hasThickness && bar3d.floorFormat().thickness == 12, "Bar3D floor thickness parsed");
    test.checkEqual(bar3d.floorFormat().fill.color.value, std::string("D9EAF7"), "Bar3D floor fill parsed");
    test.checkTrue(bar3d.sideWallFormat().present && bar3d.sideWallFormat().thickness == 18, "Bar3D side wall parsed");
    test.checkEqual(bar3d.backWallFormat().fill.color.value, std::string("FCE4D6"), "Bar3D back wall fill parsed");

    test.checkTrue(charts[4].type() == xlpp::Chart::Type::Surface && charts[4].plots().front().axisIds.size() == 3,
                   "Surface chart exposes three native axes");
    test.checkTrue(charts[5].type() == xlpp::Chart::Type::Surface3D && charts[5].view3D().present,
                   "Surface3D view model parsed");
    test.checkTrue(charts[5].floorFormat().thickness == 14 && charts[5].sideWallFormat().thickness == 16 && charts[5].backWallFormat().thickness == 22,
                   "Surface3D wall thickness metadata parsed");

    auto view = bar3d.view3D();
    view.present=true; view.hasRotationX=true; view.rotationX=40; view.hasRotationY=true; view.rotationY=75;
    view.hasHeightPercent=true; view.heightPercent=140; view.hasDepthPercent=true; view.depthPercent=220;
    view.hasRightAngleAxes=true; view.rightAngleAxes=false; view.hasPerspective=true; view.perspective=50;
    test.checkTrue(sheet->setChartView3D(bar3d.stableId(), view), "Selective view3D edit accepted");

    xlpp::ChartWallFormat floor = bar3d.floorFormat(); floor.present=true; floor.hasThickness=true; floor.thickness=30;
    floor.fill.present=true; floor.fill.kind=xlpp::ChartFillFormat::Kind::Solid; floor.fill.color={xlpp::ChartColor::Kind::SRgb,"4472C4"};
    floor.line.present=true; floor.line.color={xlpp::ChartColor::Kind::SRgb,"203864"}; floor.line.widthPoints=1.75;
    test.checkTrue(sheet->setChartFloorFormat(bar3d.stableId(), floor), "Selective floor formatting edit accepted");
    xlpp::ChartWallFormat side = bar3d.sideWallFormat(); side.present=true; side.hasThickness=true; side.thickness=28;
    side.fill.present=true; side.fill.kind=xlpp::ChartFillFormat::Kind::Solid; side.fill.color={xlpp::ChartColor::Kind::SRgb,"A9D18E"};
    test.checkTrue(sheet->setChartSideWallFormat(bar3d.stableId(), side), "Selective side-wall formatting edit accepted");
    xlpp::ChartWallFormat back = bar3d.backWallFormat(); back.present=true; back.hasThickness=true; back.thickness=26;
    back.fill.present=true; back.fill.kind=xlpp::ChartFillFormat::Kind::Solid; back.fill.color={xlpp::ChartColor::Kind::SRgb,"F4B183"};
    test.checkTrue(sheet->setChartBackWallFormat(bar3d.stableId(), back), "Selective back-wall formatting edit accepted");
    sheet->cell("K25").setValue("p0r-regression");

    const auto output = std::filesystem::temp_directory_path() / "xlpp_p0r_3d_surface.xlsx";
    workbook.save(output);
    const auto after = xlpp::internal::ZipArchive::open(output);
    const auto chart1 = after.get("xl/charts/chart1.xml");
    test.checkTrue(chart1.find("bar3DChart") != std::string::npos, "Selective edit preserves bar3DChart structure");
    test.checkTrue(chart1.find("rotX val=\"40\"") != std::string::npos && chart1.find("rotY val=\"75\"") != std::string::npos, "Edited view3D values written");
    test.checkTrue(chart1.find("4472C4") != std::string::npos && chart1.find("203864") != std::string::npos, "Edited floor formatting written");
    test.checkTrue(chart1.find("A9D18E") != std::string::npos && chart1.find("F4B183") != std::string::npos, "Edited wall fills written");
    for (int index = 2; index <= 6; ++index)
        test.checkEqual(after.get("xl/charts/chart" + std::to_string(index) + ".xml"), untouchedCharts[static_cast<std::size_t>(index-2)],
                        "Unedited 3D/surface chart part remains byte-identical");
    test.checkEqual(after.get("xl/drawings/_rels/drawing1.xml.rels"), originalDrawingRels, "P0R keeps drawing relationships byte-identical");
    test.checkEqual(after.get("xl/media/image1.png"), originalImage, "P0R keeps sibling image byte-identical");
    test.checkTrue(xlpp::internal::RelationshipGraph::fromArchive(after).validate().ok(), "P0R output package graph validates");

    xlpp::Workbook reloaded; reloaded.load(output);
    const auto* reloadSheet = reloaded.worksheet("ThreeD");
    test.checkTrue(reloadSheet != nullptr, "P0R output reloads");
    if (reloadSheet) {
        const auto& reCharts = static_cast<const xlpp::Worksheet&>(*reloadSheet).charts();
        test.checkEqual(reCharts.size(), std::size_t{6}, "P0R reload retains all chart objects");
        if (!reCharts.empty()) {
            test.checkTrue(reCharts[0].view3D().rotationX == 40 && reCharts[0].view3D().rotationY == 75, "Edited view3D survives reload");
            test.checkTrue(reCharts[0].floorFormat().thickness == 30, "Edited floor thickness survives reload");
            test.checkEqual(reCharts[0].floorFormat().fill.color.value, std::string("4472C4"), "Edited floor fill survives reload");
        }
    }
    std::filesystem::remove(output);

    xlpp::Workbook generated;
    auto& gs = generated.addWorksheet("Generated3D");
    gs.cell("A1").setValue("Category"); gs.cell("B1").setValue("S1"); gs.cell("C1").setValue("S2");
    for (std::size_t row=2; row<=5; ++row) {
        gs.cell("A"+std::to_string(row)).setValue("C"+std::to_string(row-1));
        gs.cell("B"+std::to_string(row)).setValue(static_cast<double>(row*2));
        gs.cell("C"+std::to_string(row)).setValue(static_cast<double>(row*3));
    }
    auto addSeries = [](xlpp::Chart& chart, const char* title, const char* column) {
        xlpp::ChartSeries series(title);
        series.setCategoriesReference("'Generated3D'!$A$2:$A$5");
        series.setValuesReference(std::string("'Generated3D'!$") + column + "$2:$" + column + "$5");
        chart.addSeries(std::move(series));
    };
    xlpp::Chart generatedBar(xlpp::Chart::Type::Bar3D); generatedBar.setTitle("Generated Bar3D");
    addSeries(generatedBar,"S1","B"); addSeries(generatedBar,"S2","C");
    xlpp::ChartView3D generatedView; generatedView.present=true; generatedView.hasRotationX=true; generatedView.rotationX=25;
    generatedView.hasRotationY=true; generatedView.rotationY=45; generatedView.hasDepthPercent=true; generatedView.depthPercent=180;
    generatedView.hasRightAngleAxes=true; generatedView.rightAngleAxes=false; generatedView.hasPerspective=true; generatedView.perspective=35;
    generatedBar.setView3D(generatedView);
    xlpp::ChartWallFormat generatedFloor; generatedFloor.present=true; generatedFloor.hasThickness=true; generatedFloor.thickness=12;
    generatedFloor.fill.present=true; generatedFloor.fill.kind=xlpp::ChartFillFormat::Kind::Solid; generatedFloor.fill.color={xlpp::ChartColor::Kind::SRgb,"DDEBF7"};
    generatedBar.setFloorFormat(generatedFloor);
    auto& barPlot=generatedBar.primaryPlot(); barPlot.hasGapDepth=true; barPlot.gapDepth=175; barPlot.shape="box";
    gs.addChart(std::move(generatedBar));

    xlpp::Chart generatedSurface(xlpp::Chart::Type::Surface3D); generatedSurface.setTitle("Generated Surface3D");
    addSeries(generatedSurface,"S1","B"); addSeries(generatedSurface,"S2","C");
    generatedSurface.setView3D(generatedView); auto& surfacePlot=generatedSurface.primaryPlot(); surfacePlot.hasWireframe=true; surfacePlot.wireframe=true;
    gs.addChart(std::move(generatedSurface));
    const auto generatedPath=std::filesystem::temp_directory_path()/"xlpp_p0r_generated_3d.xlsx";
    generated.save(generatedPath);
    const auto generatedZip=xlpp::internal::ZipArchive::open(generatedPath);
    const auto generatedBarXml=generatedZip.get("xl/charts/chart1.xml");
    const auto generatedSurfaceXml=generatedZip.get("xl/charts/chart2.xml");
    test.checkTrue(generatedBarXml.find("bar3DChart")!=std::string::npos && generatedBarXml.find("<c:serAx>")!=std::string::npos, "Generated Bar3D writes three-axis structure");
    test.checkTrue(generatedBarXml.find("gapDepth val=\"175\"")!=std::string::npos && generatedBarXml.find("shape val=\"box\"")!=std::string::npos, "Generated Bar3D gap-depth and shape written");
    test.checkTrue(generatedBarXml.find("view3D")!=std::string::npos && generatedBarXml.find("DDEBF7")!=std::string::npos, "Generated Bar3D view and floor formatting written");
    test.checkTrue(generatedSurfaceXml.find("surface3DChart")!=std::string::npos && generatedSurfaceXml.find("wireframe val=\"1\"")!=std::string::npos, "Generated Surface3D writes wireframe structure");
    test.checkTrue(generatedSurfaceXml.find("<c:serAx>")!=std::string::npos, "Generated Surface3D writes series axis");
    test.checkTrue(xlpp::internal::RelationshipGraph::fromArchive(generatedZip).validate().ok(), "Generated P0R package graph validates");
    xlpp::Workbook generatedReload; generatedReload.load(generatedPath);
    const auto* grs=generatedReload.worksheet("Generated3D");
    test.checkTrue(grs!=nullptr && static_cast<const xlpp::Worksheet&>(*grs).charts().size()==2, "Generated 3D workbook reloads with two charts");
    if (grs && static_cast<const xlpp::Worksheet&>(*grs).charts().size()==2) {
        const auto& gcharts=static_cast<const xlpp::Worksheet&>(*grs).charts();
        test.checkTrue(gcharts[0].type()==xlpp::Chart::Type::Bar3D && gcharts[1].type()==xlpp::Chart::Type::Surface3D, "Generated chart types survive reload");
        test.checkEqual(gcharts[0].axes().size(), std::size_t{3}, "Generated Bar3D exposes three axes after reload");
        test.checkEqual(gcharts[1].axes().size(), std::size_t{3}, "Generated Surface3D exposes three axes after reload");
    }
    std::filesystem::remove(generatedPath);
}

void testProjectedPieDoughnutRadarExpansion(TestContext& test) {
    const auto source = std::filesystem::path(XLPP_TEST_SOURCE_DIR) / "fixtures/openpyxl/projected_pie_doughnut_radar.xlsx";
    const auto before = xlpp::internal::ZipArchive::open(source);
    const auto originalChart2 = before.get("xl/charts/chart2.xml");
    const auto originalDrawingRels = before.get("xl/drawings/_rels/drawing1.xml.rels");
    const auto originalImage = before.get("xl/media/image1.png");

    xlpp::Workbook workbook; workbook.load(source);
    auto* sheet = workbook.worksheet("Projected");
    test.checkTrue(sheet != nullptr, "P0S projected-pie/doughnut/radar fixture loads");
    if (!sheet) return;
    const auto& charts = static_cast<const xlpp::Worksheet&>(*sheet).charts();
    test.checkEqual(charts.size(), std::size_t{4}, "P0S fixture exposes four charts");
    if (charts.size() != 4) return;
    test.checkTrue(charts[0].type()==xlpp::Chart::Type::PieOfPie, "Pie-of-Pie type parsed");
    test.checkTrue(charts[1].type()==xlpp::Chart::Type::BarOfPie, "Bar-of-Pie type parsed");
    test.checkTrue(charts[2].type()==xlpp::Chart::Type::Doughnut, "Doughnut type parsed");
    test.checkTrue(charts[3].type()==xlpp::Chart::Type::Radar, "Radar type parsed");

    const auto& piePlot=charts[0].plots().front();
    test.checkTrue(piePlot.projectedPie.present && piePlot.projectedPie.ofPieType=="pie", "Pie-of-Pie options parsed");
    test.checkEqual(piePlot.projectedPie.gapWidth, 180, "Pie-of-Pie gap width parsed");
    test.checkEqual(piePlot.projectedPie.splitType, std::string("cust"), "Pie-of-Pie custom split type parsed");
    test.checkEqual(piePlot.projectedPie.customSplitPoints.size(), std::size_t{3}, "Pie-of-Pie custom split points parsed");
    test.checkEqual(piePlot.projectedPie.secondPlotSize, 120, "Pie-of-Pie second plot size parsed");
    const auto& barPlot=charts[1].plots().front();
    test.checkTrue(barPlot.projectedPie.present && barPlot.projectedPie.ofPieType=="bar", "Bar-of-Pie options parsed");
    test.checkEqual(barPlot.projectedPie.splitType, std::string("val"), "Bar-of-Pie value split parsed");
    test.checkTrue(barPlot.projectedPie.hasSplitPosition, "Bar-of-Pie split position present");
    test.checkNear(barPlot.projectedPie.splitPosition, 12.0, 1e-12, "Bar-of-Pie split position parsed");

    const auto& doughnutPlot=charts[2].plots().front();
    test.checkTrue(doughnutPlot.hasFirstSliceAngle && doughnutPlot.firstSliceAngle==45, "Doughnut first-slice angle parsed");
    test.checkTrue(doughnutPlot.hasHoleSize && doughnutPlot.holeSize==55, "Doughnut hole size parsed");
    const auto& radarPlot=charts[3].plots().front();
    test.checkEqual(radarPlot.radarStyle, std::string("filled"), "Radar style parsed");
    test.checkEqual(charts[3].series().size(), std::size_t{2}, "Radar series parsed");
    test.checkEqual(charts[3].series()[0].markerFormat().symbol, std::string("circle"), "Radar marker symbol parsed");
    test.checkEqual(charts[3].series()[0].markerFormat().size, 7, "Radar marker size parsed");

    auto projected=piePlot.projectedPie; projected.present=true; projected.splitType="percent"; projected.hasSplitPosition=true;
    projected.splitPosition=25; projected.gapWidth=200; projected.secondPlotSize=110; projected.customSplitPoints.clear();
    projected.hasSeriesLines=true; projected.seriesLinesFormat.present=true; projected.seriesLinesFormat.widthPoints=1.5;
    projected.seriesLinesFormat.color={xlpp::ChartColor::Kind::SRgb,"7030A0"};
    test.checkTrue(sheet->setChartPlotProjectedPieOptions(charts[0].stableId(),0,projected), "Selective projected-pie options edit accepted");
    test.checkTrue(sheet->setChartPlotFirstSliceAngle(charts[2].stableId(),0,120), "Selective doughnut first-slice edit accepted");
    test.checkTrue(sheet->setChartPlotDoughnutHoleSize(charts[2].stableId(),0,70), "Selective doughnut hole-size edit accepted");
    test.checkTrue(sheet->setChartPlotRadarStyle(charts[3].stableId(),0,"marker"), "Selective radar-style edit accepted");
    auto marker=charts[3].series()[0].markerFormat(); marker.present=true; marker.symbol="star"; marker.size=9;
    marker.fill.present=true; marker.fill.kind=xlpp::ChartFillFormat::Kind::Solid; marker.fill.color={xlpp::ChartColor::Kind::SRgb,"FF0000"};
    test.checkTrue(sheet->setChartSeriesMarkerFormat(charts[3].stableId(),0,marker), "Selective radar marker edit accepted");
    test.checkTrue(!sheet->setChartPlotDoughnutHoleSize(charts[2].stableId(),0,5), "Invalid doughnut hole size rejected");
    test.checkTrue(!sheet->setChartPlotRadarStyle(charts[3].stableId(),0,"unsupported"), "Invalid radar style rejected");
    sheet->cell("K26").setValue("p0s-regression");

    const auto output=std::filesystem::temp_directory_path()/"xlpp_p0s_projected_pie_doughnut_radar.xlsx";
    workbook.save(output);
    const auto after=xlpp::internal::ZipArchive::open(output);
    const auto chart1=after.get("xl/charts/chart1.xml");
    const auto chart3=after.get("xl/charts/chart3.xml");
    const auto chart4=after.get("xl/charts/chart4.xml");
    test.checkTrue(chart1.find("ofPieType val=\"pie\"")!=std::string::npos && chart1.find("splitType val=\"percent\"")!=std::string::npos, "Projected-pie selective options serialized");
    test.checkTrue(chart1.find("splitPos val=\"25")!=std::string::npos && chart1.find("secondPieSize val=\"110\"")!=std::string::npos, "Projected-pie split position and size serialized");
    test.checkTrue(chart1.find("7030A0")!=std::string::npos, "Projected-pie series-line formatting serialized");
    test.checkTrue(chart3.find("firstSliceAng val=\"120\"")!=std::string::npos && chart3.find("holeSize val=\"70\"")!=std::string::npos, "Doughnut options serialized");
    test.checkTrue(chart4.find("radarStyle val=\"marker\"")!=std::string::npos && chart4.find("symbol val=\"star\"")!=std::string::npos, "Radar style and marker serialized");
    test.checkEqual(after.get("xl/charts/chart2.xml"), originalChart2, "Untouched Bar-of-Pie chart remains byte-identical");
    test.checkEqual(after.get("xl/drawings/_rels/drawing1.xml.rels"), originalDrawingRels, "P0S keeps drawing relationships byte-identical");
    test.checkEqual(after.get("xl/media/image1.png"), originalImage, "P0S keeps sibling image byte-identical");
    test.checkTrue(xlpp::internal::RelationshipGraph::fromArchive(after).validate().ok(), "P0S selective output package graph validates");

    xlpp::Workbook reloaded; reloaded.load(output); const auto* rs=reloaded.worksheet("Projected");
    test.checkTrue(rs!=nullptr, "P0S selective output reloads");
    if(rs){ const auto& rc=static_cast<const xlpp::Worksheet&>(*rs).charts(); test.checkEqual(rc.size(),std::size_t{4},"P0S reload keeps four charts");
        if(rc.size()==4){
            test.checkEqual(rc[0].plots()[0].projectedPie.splitType,std::string("percent"),"Projected-pie edited split survives reload");
            test.checkEqual(rc[2].plots()[0].holeSize,70,"Doughnut edited hole survives reload");
            test.checkEqual(rc[3].plots()[0].radarStyle,std::string("marker"),"Radar edited style survives reload");
            test.checkEqual(rc[3].series()[0].markerFormat().symbol,std::string("star"),"Radar edited marker survives reload");
        }
    }
    std::filesystem::remove(output);

    xlpp::Workbook generated; auto& gs=generated.addWorksheet("Generated");
    gs.append({std::string("Category"),std::string("Primary"),std::string("Secondary")});
    const std::array<std::array<double,2>,6> vals{{{{10,14}},{{20,18}},{{30,12}},{{5,9}},{{8,7}},{{13,11}}}};
    for(std::size_t i=0;i<vals.size();++i){ gs.cell(i+2,1).setValue(std::string(1,static_cast<char>('A'+i))); gs.cell(i+2,2).setValue(vals[i][0]); gs.cell(i+2,3).setValue(vals[i][1]); }
    auto addOneSeries=[](xlpp::Chart& c,const char* title,const char* col){ xlpp::ChartSeries s(title); s.setCategoriesReference("'Generated'!$A$2:$A$7"); s.setValuesReference(std::string("'Generated'!$")+col+"$2:$"+col+"$7"); c.addSeries(std::move(s)); };
    xlpp::Chart gp(xlpp::Chart::Type::PieOfPie); gp.setTitle("Generated Pie-of-Pie"); addOneSeries(gp,"Primary","B");
    auto& gpp=gp.primaryPlot(); gpp.projectedPie.present=true; gpp.projectedPie.splitType="cust"; gpp.projectedPie.customSplitPoints={1,4}; gpp.projectedPie.gapWidth=175; gpp.projectedPie.secondPlotSize=105; gs.addChart(std::move(gp));
    xlpp::Chart gb(xlpp::Chart::Type::BarOfPie); gb.setTitle("Generated Bar-of-Pie"); addOneSeries(gb,"Primary","B");
    auto& gbp=gb.primaryPlot(); gbp.projectedPie.present=true; gbp.projectedPie.splitType="val"; gbp.projectedPie.hasSplitPosition=true; gbp.projectedPie.splitPosition=12; gbp.projectedPie.secondPlotSize=95; gs.addChart(std::move(gb));
    xlpp::Chart gd(xlpp::Chart::Type::Doughnut); gd.setTitle("Generated Doughnut"); addOneSeries(gd,"Primary","B"); auto& gdp=gd.primaryPlot(); gdp.hasFirstSliceAngle=true; gdp.firstSliceAngle=90; gdp.hasHoleSize=true; gdp.holeSize=60; gs.addChart(std::move(gd));
    xlpp::Chart gr(xlpp::Chart::Type::Radar); gr.setTitle("Generated Radar"); addOneSeries(gr,"Primary","B"); addOneSeries(gr,"Secondary","C"); auto& grp=gr.primaryPlot(); grp.radarStyle="marker";
    xlpp::ChartMarkerFormat gm; gm.present=true; gm.symbol="diamond"; gm.size=8; gr.series()[0].setMarkerFormat(gm); gs.addChart(std::move(gr));
    const auto generatedPath=std::filesystem::temp_directory_path()/"xlpp_p0s_generated_projected_radar.xlsx"; generated.save(generatedPath);
    const auto gz=xlpp::internal::ZipArchive::open(generatedPath);
    test.checkTrue(gz.get("xl/charts/chart1.xml").find("<c:ofPieChart>")!=std::string::npos && gz.get("xl/charts/chart1.xml").find("splitType val=\"cust\"")!=std::string::npos, "Generated Pie-of-Pie XML written");
    test.checkTrue(gz.get("xl/charts/chart2.xml").find("ofPieType val=\"bar\"")!=std::string::npos, "Generated Bar-of-Pie XML written");
    test.checkTrue(gz.get("xl/charts/chart3.xml").find("firstSliceAng val=\"90\"")!=std::string::npos && gz.get("xl/charts/chart3.xml").find("holeSize val=\"60\"")!=std::string::npos, "Generated Doughnut options written");
    test.checkTrue(gz.get("xl/charts/chart4.xml").find("radarStyle val=\"marker\"")!=std::string::npos && gz.get("xl/charts/chart4.xml").find("symbol val=\"diamond\"")!=std::string::npos, "Generated Radar style and marker written");
    test.checkTrue(xlpp::internal::RelationshipGraph::fromArchive(gz).validate().ok(), "Generated P0S package graph validates");
    xlpp::Workbook generatedReload; generatedReload.load(generatedPath); const auto* grs=generatedReload.worksheet("Generated");
    test.checkTrue(grs!=nullptr,"Generated P0S workbook reloads"); if(grs){ const auto& gc=static_cast<const xlpp::Worksheet&>(*grs).charts(); test.checkEqual(gc.size(),std::size_t{4},"Generated P0S charts reload"); if(gc.size()==4){ test.checkTrue(gc[0].type()==xlpp::Chart::Type::PieOfPie&&gc[1].type()==xlpp::Chart::Type::BarOfPie,"Generated projected-pie types survive reload"); test.checkEqual(gc[2].plots()[0].holeSize,60,"Generated doughnut hole survives reload"); test.checkEqual(gc[3].plots()[0].radarStyle,std::string("marker"),"Generated radar style survives reload"); }}
    std::filesystem::remove(generatedPath);
}

void testChartStyleThemeAndSeriesCaches(TestContext& test) {
    const auto source = std::filesystem::path(XLPP_TEST_SOURCE_DIR) / "fixtures/libreoffice/chart_style_theme_caches.xlsx";
    const auto before = xlpp::internal::ZipArchive::open(source);
    const auto originalStylePart = before.get("xl/charts/style1.xml");
    const auto originalColorStylePart = before.get("xl/charts/colors1.xml");
    const auto originalChartRels = before.get("xl/charts/_rels/chart1.xml.rels");
    const auto originalDrawingRels = before.get("xl/drawings/_rels/drawing1.xml.rels");

    xlpp::Workbook workbook; workbook.load(source);
    auto* sheet = workbook.worksheet("Objects");
    test.checkTrue(sheet != nullptr, "P0T style/theme/cache fixture loads");
    if (!sheet) return;
    const auto& charts = static_cast<const xlpp::Worksheet&>(*sheet).charts();
    test.checkEqual(charts.size(), std::size_t{1}, "P0T fixture exposes one chart");
    if (charts.empty()) return;
    const auto& chart = charts.front();
    test.checkEqual(chart.style(), std::string("10"), "Imported chart style ID parsed");
    test.checkTrue(chart.themePalette().present, "Workbook theme palette parsed for chart");
    test.checkEqual(chart.themePalette().baseColor("accent1"), std::string("4f81bd"), "Theme accent1 base color parsed");
    test.checkTrue(chart.styleResources().chartStylePresent && chart.styleResources().colorStylePresent, "Chart style/color-style relationships discovered");
    test.checkEqual(chart.styleResources().chartStylePart, std::string("xl/charts/style1.xml"), "Chart style part resolved");
    test.checkEqual(chart.styleResources().colorStylePart, std::string("xl/charts/colors1.xml"), "Chart color-style part resolved");
    test.checkEqual(chart.series().size(), std::size_t{1}, "P0T fixture series parsed");
    if (chart.series().empty()) return;
    const auto& series = chart.series().front();
    test.checkEqual(series.titleReference(), std::string("Objects!B1"), "Series title reference parsed");
    test.checkTrue(series.titleCache().present && !series.titleCache().numeric, "Series title strCache parsed");
    test.checkEqual(series.titleCache().points.size(), std::size_t{1}, "Title cache point count parsed");
    test.checkEqual(series.titleCache().points.front().value, std::string("Amount"), "Title cache value parsed");
    test.checkTrue(series.categoriesCache().present && !series.categoriesCache().numeric, "Category strCache parsed");
    test.checkEqual(series.categoriesCache().effectivePointCount(), std::size_t{3}, "Category cache count parsed");
    test.checkTrue(series.valuesCache().present && series.valuesCache().numeric, "Value numCache parsed");
    test.checkEqual(series.valuesCache().formatCode, std::string("General"), "Value cache format code parsed");
    test.checkEqual(series.valuesCache().points.size(), std::size_t{3}, "Value cache points parsed");
    test.checkTrue(series.fillFormat().color.kind == xlpp::ChartColor::Kind::Scheme, "Scheme series color parsed");
    test.checkEqual(series.fillFormat().color.value, std::string("accent1"), "Scheme series color name parsed");
    test.checkEqual(chart.resolveThemeBaseColor(series.fillFormat().color), std::string("4f81bd"), "Scheme color resolves through workbook theme");
    test.checkTrue(!series.fillFormat().color.transforms.empty(), "Scheme color transforms preserved in model");

    // Unrelated edit must keep chart/style resources byte-identical.
    sheet->cell("K27").setValue("p0t-unrelated");
    const auto unrelated = std::filesystem::temp_directory_path()/"xlpp_p0t_unrelated.xlsx";
    workbook.save(unrelated);
    const auto unrelatedZip = xlpp::internal::ZipArchive::open(unrelated);
    test.checkEqual(unrelatedZip.get("xl/charts/chart1.xml"), before.get("xl/charts/chart1.xml"), "Unrelated edit keeps chart XML byte-identical");
    test.checkEqual(unrelatedZip.get("xl/charts/style1.xml"), originalStylePart, "Unrelated edit preserves chart-style part byte-identical");
    test.checkEqual(unrelatedZip.get("xl/charts/colors1.xml"), originalColorStylePart, "Unrelated edit preserves color-style part byte-identical");
    test.checkEqual(unrelatedZip.get("xl/charts/_rels/chart1.xml.rels"), originalChartRels, "Unrelated edit preserves chart relationships byte-identical");
    std::filesystem::remove(unrelated);

    xlpp::ChartSeriesCache cat; cat.present=true; cat.numeric=false; cat.pointCount=3; cat.points={{0,"Alpha"},{1,"Beta"},{2,"Gamma"}};
    xlpp::ChartSeriesCache val; val.present=true; val.numeric=true; val.formatCode="0.00"; val.pointCount=3; val.points={{0,"11.25"},{1,"22.5"},{2,"33.75"}};
    xlpp::ChartSeriesCache title; title.present=true; title.numeric=false; title.pointCount=1; title.points={{0,"Updated Amount"}};
    test.checkTrue(sheet->setChartStyle(chart.stableId(), "15"), "Selective chart style edit accepted");
    test.checkTrue(sheet->setChartSeriesCategoryCache(chart.stableId(),0,cat), "Selective category cache edit accepted");
    test.checkTrue(sheet->setChartSeriesValueCache(chart.stableId(),0,val), "Selective value cache edit accepted");
    test.checkTrue(sheet->setChartSeriesTitleCache(chart.stableId(),0,title), "Selective title cache edit accepted");
    xlpp::ChartSeriesCache invalid=val; invalid.pointCount=1;
    test.checkTrue(!sheet->setChartSeriesValueCache(chart.stableId(),0,invalid), "Inconsistent cache pointCount rejected");
    const auto output=std::filesystem::temp_directory_path()/"xlpp_p0t_style_cache_edit.xlsx";
    workbook.save(output);
    const auto after=xlpp::internal::ZipArchive::open(output);
    const auto chartXml=after.get("xl/charts/chart1.xml");
    test.checkTrue(chartXml.find("style val=\"15\"")!=std::string::npos, "Selective chart style serialized");
    test.checkTrue(chartXml.find("Updated Amount")!=std::string::npos && chartXml.find("Alpha")!=std::string::npos && chartXml.find("33.75")!=std::string::npos, "Selective cache values serialized");
    test.checkTrue(chartXml.find("<c:formatCode>0.00</c:formatCode>")!=std::string::npos, "Numeric cache format code serialized");
    test.checkEqual(after.get("xl/charts/style1.xml"), originalStylePart, "Style resource survives chart metadata edit byte-identical");
    test.checkEqual(after.get("xl/charts/colors1.xml"), originalColorStylePart, "Color-style resource survives chart metadata edit byte-identical");
    test.checkEqual(after.get("xl/charts/_rels/chart1.xml.rels"), originalChartRels, "Chart style relationships survive selective edit byte-identical");
    test.checkEqual(after.get("xl/drawings/_rels/drawing1.xml.rels"), originalDrawingRels, "Drawing relationships remain byte-identical");
    test.checkTrue(xlpp::internal::RelationshipGraph::fromArchive(after).validate().ok(), "P0T selective package graph validates");

    xlpp::Workbook reload; reload.load(output); const auto* rs=reload.worksheet("Objects");
    test.checkTrue(rs!=nullptr, "P0T selective output reloads");
    if(rs){ const auto& rc=static_cast<const xlpp::Worksheet&>(*rs).charts(); test.checkEqual(rc.size(),std::size_t{1},"P0T reload keeps chart"); if(!rc.empty()){
        const auto& r=rc.front(); test.checkEqual(r.style(),std::string("15"),"Edited style reloads");
        test.checkEqual(r.series()[0].categoriesCache().points[1].value,std::string("Beta"),"Edited category cache reloads");
        test.checkEqual(r.series()[0].valuesCache().formatCode,std::string("0.00"),"Edited numeric cache format reloads");
        test.checkEqual(r.series()[0].titleCache().points[0].value,std::string("Updated Amount"),"Edited title cache reloads");
    }}
    std::filesystem::remove(output);

    // Generated chart caches should be first-class rather than preservation-only metadata.
    xlpp::Workbook generated; auto& gs=generated.addWorksheet("Caches");
    gs.append({std::string("Category"),std::string("Amount")});
    gs.append({std::string("A"),10.0}); gs.append({std::string("B"),20.0}); gs.append({std::string("C"),30.0});
    xlpp::Chart gc(xlpp::Chart::Type::Bar); gc.setStyle("12"); xlpp::ChartSeries gseries("Amount");
    gseries.setTitleReference("'Caches'!$B$1"); gseries.setCategoriesReference("'Caches'!$A$2:$A$4"); gseries.setValuesReference("'Caches'!$B$2:$B$4");
    xlpp::ChartSeriesCache gtc; gtc.present=true; gtc.points={{0,"Amount"}}; gseries.setTitleCache(gtc);
    xlpp::ChartSeriesCache gcc; gcc.present=true; gcc.points={{0,"A"},{1,"B"},{2,"C"}}; gseries.setCategoriesCache(gcc);
    xlpp::ChartSeriesCache gvc; gvc.present=true; gvc.numeric=true; gvc.formatCode="0"; gvc.points={{0,"10"},{1,"20"},{2,"30"}}; gseries.setValuesCache(gvc);
    gc.addSeries(std::move(gseries)); gs.addChart(std::move(gc));
    const auto generatedPath=std::filesystem::temp_directory_path()/"xlpp_p0t_generated_caches.xlsx"; generated.save(generatedPath);
    const auto generatedZip=xlpp::internal::ZipArchive::open(generatedPath); const auto generatedXml=generatedZip.get("xl/charts/chart1.xml");
    test.checkTrue(generatedXml.find("style val=\"12\"")!=std::string::npos && generatedXml.find("strCache")!=std::string::npos && generatedXml.find("numCache")!=std::string::npos, "Generated chart writes style and caches");
    xlpp::Workbook generatedReload; generatedReload.load(generatedPath); const auto* grs=generatedReload.worksheet("Caches");
    test.checkTrue(grs!=nullptr,"Generated cache workbook reloads"); if(grs){ const auto& c=static_cast<const xlpp::Worksheet&>(*grs).charts(); if(!c.empty()){ test.checkEqual(c[0].series()[0].titleCache().points[0].value,std::string("Amount"),"Generated title cache reloads"); test.checkEqual(c[0].series()[0].valuesCache().points.size(),std::size_t{3},"Generated value cache reloads"); }}
    std::filesystem::remove(generatedPath);
}

void testChartCacheSynchronizationAndThemeTransforms(TestContext& test) {
    const auto source = std::filesystem::path(XLPP_TEST_SOURCE_DIR) / "fixtures/libreoffice/chart_style_theme_caches.xlsx";
    xlpp::Workbook workbook; workbook.load(source);
    auto* sheet = workbook.worksheet("Objects");
    test.checkTrue(sheet != nullptr, "P0U cache-sync fixture loads");
    if (!sheet) return;
    const auto& initialCharts = static_cast<const xlpp::Worksheet&>(*sheet).charts();
    test.checkEqual(initialCharts.size(), std::size_t{1}, "P0U fixture exposes one chart");
    if (initialCharts.empty()) return;
    const auto& initial = initialCharts.front();
    test.checkTrue(initial.themePalette().fontScheme.present, "Theme font scheme parsed");
    test.checkEqual(initial.themePalette().fontScheme.name, std::string("Office"), "Theme font scheme name parsed");
    test.checkEqual(initial.themePalette().fontScheme.majorLatinTypeface, std::string("Cambria"), "Theme major Latin font parsed");
    test.checkEqual(initial.themePalette().fontScheme.minorLatinTypeface, std::string("Calibri"), "Theme minor Latin font parsed");
    test.checkTrue(initial.themePalette().effectScheme.present, "Theme effect scheme parsed");
    test.checkEqual(initial.themePalette().effectScheme.fillStyleCount, std::size_t{3}, "Theme fill style count parsed");
    test.checkEqual(initial.themePalette().effectScheme.lineStyleCount, std::size_t{3}, "Theme line style count parsed");
    test.checkEqual(initial.themePalette().effectScheme.effectStyleCount, std::size_t{3}, "Theme effect style count parsed");
    test.checkEqual(initial.themePalette().effectScheme.backgroundFillStyleCount, std::size_t{3}, "Theme background fill count parsed");
    const auto transformed = initial.resolveThemeColor(initial.series()[0].fillFormat().color);
    test.checkTrue(transformed.present, "Theme transformed color resolves");
    test.checkEqual(transformed.srgb(), std::string("729ACA"), "Theme tint transform produces final RGB");
    test.checkNear(transformed.alpha, 1.0, 1e-12, "Theme transformed alpha defaults opaque");

    xlpp::ChartColor transformedColor{xlpp::ChartColor::Kind::Scheme, "accent1"};
    transformedColor.transforms.push_back({xlpp::ChartColorTransform::Kind::Shade, 50000});
    transformedColor.transforms.push_back({xlpp::ChartColorTransform::Kind::Alpha, 60000});
    const auto shaded = initial.resolveThemeColor(transformedColor);
    test.checkEqual(shaded.srgb(), std::string("28415F"), "Theme shade transform applies sequentially");
    test.checkNear(shaded.alpha, 0.6, 1e-12, "Theme alpha transform resolves");

    // Edit source cells and rebuild the imported chart caches from A1 references.
    sheet->cell("B1").setValue("Synced Amount");
    sheet->cell("A2").setValue("North");
    sheet->cell("A3").clear();
    sheet->cell("A4").setValue("South");
    sheet->cell("B2").setValue(101.5); sheet->cell("B2").setNumberFormat("0.0");
    sheet->cell("B3").clear();
    sheet->cell("B4").setValue(303.25); sheet->cell("B4").setNumberFormat("0.0");
    const auto report = workbook.synchronizeChartCaches();
    test.checkEqual(report.chartsVisited, std::size_t{1}, "Cache sync visits imported chart");
    test.checkEqual(report.seriesVisited, std::size_t{1}, "Cache sync visits imported series");
    test.checkEqual(report.cachesUpdated, std::size_t{3}, "Cache sync updates title/category/value caches");
    test.checkEqual(report.referencesSkipped, std::size_t{0}, "Cache sync accepts local A1 references");
    const auto& synced = static_cast<const xlpp::Worksheet&>(*sheet).charts().front().series().front();
    test.checkEqual(synced.titleCache().points[0].value, std::string("Synced Amount"), "Title cache rebuilt from worksheet cell");
    test.checkTrue(synced.categoriesCache().sparse(), "String category cache records sparse blank cell");
    test.checkEqual(synced.categoriesCache().pointCount, std::size_t{3}, "Sparse category cache keeps source range length");
    test.checkEqual(synced.categoriesCache().points.size(), std::size_t{2}, "Sparse category cache omits blank point");
    test.checkEqual(synced.categoriesCache().points[1].index, std::size_t{2}, "Sparse category cache retains source index");
    test.checkTrue(synced.valuesCache().numeric && synced.valuesCache().sparse(), "Numeric value cache rebuilt as sparse numCache");
    test.checkEqual(synced.valuesCache().formatCode, std::string("0.0"), "Numeric cache adopts worksheet number format");
    test.checkEqual(synced.valuesCache().points[1].value, std::string("303.25"), "Numeric cache reads edited worksheet value");
    test.checkTrue(synced.valuesCache().ordered(), "Synchronized cache points are index ordered");

    const auto output = std::filesystem::temp_directory_path() / "xlpp_p0u_cache_sync.xlsx";
    workbook.save(output);
    xlpp::Workbook reload; reload.load(output);
    const auto* reloadedSheet = reload.worksheet("Objects");
    test.checkTrue(reloadedSheet != nullptr, "P0U synchronized workbook reloads");
    if (reloadedSheet) {
        const auto& rs = static_cast<const xlpp::Worksheet&>(*reloadedSheet).charts().front().series().front();
        test.checkEqual(rs.titleCache().points[0].value, std::string("Synced Amount"), "Synchronized title cache survives save/reload");
        test.checkTrue(rs.categoriesCache().sparse(), "Sparse category cache survives save/reload");
        test.checkTrue(rs.valuesCache().sparse(), "Sparse numeric cache survives save/reload");
        test.checkEqual(rs.valuesCache().formatCode, std::string("0.0"), "Synchronized numeric format survives reload");
    }
    std::filesystem::remove(output);

    // Generated charts can synchronize cross-sheet quoted references without pre-built caches.
    xlpp::Workbook generated;
    auto& data = generated.addWorksheet("Data O'Brien");
    data.cell("A1").setValue("Category"); data.cell("B1").setValue("Value");
    data.cell("A2").setValue("A"); data.cell("B2").setValue(10.0); data.cell("B2").setNumberFormat("0.00");
    data.cell("A3").setValue("B"); // B3 intentionally blank -> sparse value cache.
    data.cell("A4").setValue("C"); data.cell("B4").setValue(30.0); data.cell("B4").setNumberFormat("0.00");
    auto& chartSheet = generated.addWorksheet("Chart Sheet");
    xlpp::Chart generatedChart(xlpp::Chart::Type::Line);
    xlpp::ChartSeries generatedSeries("Value");
    generatedSeries.setTitleReference("'Data O''Brien'!$B$1");
    generatedSeries.setCategoriesReference("'Data O''Brien'!$A$2:$A$4");
    generatedSeries.setValuesReference("'Data O''Brien'!$B$2:$B$4");
    generatedChart.addSeries(std::move(generatedSeries)); chartSheet.addChart(std::move(generatedChart));
    const auto generatedReport = generated.synchronizeChartCaches();
    test.checkEqual(generatedReport.cachesUpdated, std::size_t{3}, "Generated cache sync handles quoted cross-sheet references");
    test.checkEqual(generatedReport.referencesSkipped, std::size_t{0}, "Generated cache sync accepts apostrophe-escaped sheet name");
    const auto& generatedCaches = static_cast<const xlpp::Worksheet&>(chartSheet).charts().front().series().front();
    test.checkEqual(generatedCaches.categoriesCache().points.size(), std::size_t{3}, "Generated string cache contains all categories");
    test.checkTrue(generatedCaches.valuesCache().sparse(), "Generated numeric cache preserves blank source point");
    test.checkEqual(generatedCaches.valuesCache().formatCode, std::string("0.00"), "Generated cache uses source number format");

    xlpp::ChartSeriesCache duplicate; duplicate.present=true; duplicate.pointCount=2; duplicate.points={{0,"A"},{0,"B"}};
    test.checkTrue(duplicate.hasDuplicateIndexes() && !duplicate.valid(), "Duplicate cache indexes are rejected by public validation");
    xlpp::ChartSeriesCache unordered; unordered.present=true; unordered.pointCount=3; unordered.points={{2,"C"},{0,"A"}};
    test.checkTrue(unordered.valid() && !unordered.ordered() && unordered.sparse(), "Sparse unordered cache is valid but detectable before serialization sorting");

    const auto generatedPath = std::filesystem::temp_directory_path() / "xlpp_p0u_generated_cache_sync.xlsx";
    generated.save(generatedPath);
    xlpp::Workbook generatedReload; generatedReload.load(generatedPath);
    const auto* gcSheet = generatedReload.worksheet("Chart Sheet");
    test.checkTrue(gcSheet != nullptr, "Generated synchronized cache workbook reloads");
    if (gcSheet) {
        const auto& cache = static_cast<const xlpp::Worksheet&>(*gcSheet).charts().front().series().front().valuesCache();
        test.checkTrue(cache.present && cache.numeric && cache.sparse(), "Generated synchronized sparse cache is serialized");
        test.checkEqual(cache.pointCount, std::size_t{3}, "Generated synchronized cache pointCount survives reload");
    }
    std::filesystem::remove(generatedPath);
}

void testChartCacheDependencyTrackingAndAutoSave(TestContext& test) {
    xlpp::Workbook workbook;
    auto& data = workbook.addWorksheet("Data");
    data.cell("A1").setValue("Category"); data.cell("B1").setValue("Value");
    data.cell("A2").setValue("A"); data.cell("B2").setValue(10.0); data.cell("B2").setNumberFormat("0.0");
    data.cell("A3").setValue("B"); data.cell("B3").setValue(20.0); data.cell("B3").setNumberFormat("0.0");
    data.cell("A4").setValue("C"); data.cell("B4").setValue(30.0); data.cell("B4").setNumberFormat("0.0");

    auto& charts = workbook.addWorksheet("Charts");
    xlpp::Chart chart(xlpp::Chart::Type::Line);
    xlpp::ChartSeries series("Value");
    series.setTitleReference("'Data'!$B$1");
    series.setCategoriesReference("'Data'!$A$2:$A$4");
    series.setValuesReference("'Data'!$B$2:$B$4");
    chart.addSeries(std::move(series));
    charts.addChart(std::move(chart));

    xlpp::ChartCacheSyncOptions incremental;
    incremental.changedReferencesOnly = true;
    const auto first = workbook.synchronizeChartCaches(incremental);
    test.checkEqual(first.referencesChecked, std::size_t{3}, "P0V initial dependency sync checks three references");
    test.checkEqual(first.dependenciesRegistered, std::size_t{3}, "P0V initial dependency sync registers three references");
    test.checkEqual(first.cachesUpdated, std::size_t{3}, "P0V initial dependency sync materializes all caches");
    test.checkEqual(workbook.trackedChartCacheDependencyCount(), std::size_t{3}, "P0V workbook tracks three chart dependencies");

    const auto unchanged = workbook.synchronizeChartCaches(incremental);
    test.checkEqual(unchanged.referencesChecked, std::size_t{3}, "P0V repeated dependency sync still validates reference identities");
    test.checkEqual(unchanged.referencesUnchanged, std::size_t{3}, "P0V repeated dependency sync identifies all references as unchanged");
    test.checkEqual(unchanged.cachesUpdated, std::size_t{0}, "P0V repeated dependency sync skips cache rebuilds");

    // An unrelated edit must not invalidate any tracked chart source.
    data.cell("D10").setValue("unrelated");
    const auto unrelated = workbook.synchronizeChartCaches(incremental);
    test.checkEqual(unrelated.referencesUnchanged, std::size_t{3}, "P0V unrelated worksheet edit does not invalidate chart dependencies");
    test.checkEqual(unrelated.dependenciesChanged, std::size_t{0}, "P0V unrelated worksheet edit reports no dependency change");
    test.checkEqual(unrelated.cachesUpdated, std::size_t{0}, "P0V unrelated worksheet edit does not rewrite caches");

    // A value edit invalidates only the value-range dependency.
    data.cell("B3").setValue(25.0);
    const auto valueChanged = workbook.synchronizeChartCaches(incremental);
    test.checkEqual(valueChanged.referencesUnchanged, std::size_t{2}, "P0V value edit leaves title/category dependencies unchanged");
    test.checkEqual(valueChanged.dependenciesChanged, std::size_t{1}, "P0V value edit invalidates one dependency");
    test.checkEqual(valueChanged.cachesUpdated, std::size_t{1}, "P0V value edit rebuilds only one cache");
    const auto& syncedSeries = static_cast<const xlpp::Worksheet&>(charts).charts().front().series().front();
    test.checkEqual(syncedSeries.valuesCache().points[1].value, std::string("25"), "P0V selective rebuild updates changed value cache");

    // Number-format changes also matter to numCache output and must invalidate.
    data.cell("B2").setNumberFormat("0.00");
    const auto formatChanged = workbook.synchronizeChartCaches(incremental);
    test.checkEqual(formatChanged.dependenciesChanged, std::size_t{1}, "P0V number-format edit invalidates numeric cache dependency");
    test.checkEqual(formatChanged.cachesUpdated, std::size_t{0}, "P0V preserves existing explicit cache format when content is otherwise equal");

    workbook.resetChartCacheDependencyTracking();
    test.checkEqual(workbook.trackedChartCacheDependencyCount(), std::size_t{0}, "P0V dependency tracking can be reset explicitly");
    const auto afterReset = workbook.synchronizeChartCaches(incremental);
    test.checkEqual(afterReset.dependenciesRegistered, std::size_t{3}, "P0V reset forces all supported references to register again");

    // Save-time synchronization is opt-in and works on a private copy: the
    // caller's generated series remains cache-free while the serialized file
    // contains materialized caches.
    xlpp::Workbook autosave;
    auto& autoData = autosave.addWorksheet("Source");
    autoData.cell("A1").setValue("Name"); autoData.cell("B1").setValue("Amount");
    autoData.cell("A2").setValue("One"); autoData.cell("B2").setValue(1.0);
    autoData.cell("A3").setValue("Two"); autoData.cell("B3").setValue(2.0);
    auto& autoChartSheet = autosave.addWorksheet("Chart");
    xlpp::Chart autoChart(xlpp::Chart::Type::Bar);
    xlpp::ChartSeries autoSeries("Amount");
    autoSeries.setTitleReference("'Source'!$B$1");
    autoSeries.setCategoriesReference("'Source'!$A$2:$A$3");
    autoSeries.setValuesReference("'Source'!$B$2:$B$3");
    autoChart.addSeries(std::move(autoSeries)); autoChartSheet.addChart(std::move(autoChart));
    const auto& beforeSaveSeries = static_cast<const xlpp::Worksheet&>(autoChartSheet).charts().front().series().front();
    test.checkTrue(!beforeSaveSeries.titleCache().present && !beforeSaveSeries.categoriesCache().present && !beforeSaveSeries.valuesCache().present,
                   "P0V autosave fixture begins without materialized caches");

    xlpp::SaveOptions saveOptions;
    saveOptions.synchronizeChartCaches = true;
    saveOptions.synchronizeChangedChartCachesOnly = true;
    const auto path = std::filesystem::temp_directory_path() / "xlpp_p0v_auto_chart_cache_sync.xlsx";
    autosave.save(path, saveOptions);
    const auto& afterSaveSeries = static_cast<const xlpp::Worksheet&>(autoChartSheet).charts().front().series().front();
    test.checkTrue(!afterSaveSeries.titleCache().present && !afterSaveSeries.categoriesCache().present && !afterSaveSeries.valuesCache().present,
                   "P0V save-time cache sync does not mutate caller workbook");
    xlpp::Workbook reloaded; reloaded.load(path);
    const auto* reloadedChartSheet = reloaded.worksheet("Chart");
    test.checkTrue(reloadedChartSheet != nullptr, "P0V auto-synchronized workbook reloads");
    if (reloadedChartSheet) {
        const auto& savedSeries = static_cast<const xlpp::Worksheet&>(*reloadedChartSheet).charts().front().series().front();
        test.checkTrue(savedSeries.titleCache().present && savedSeries.categoriesCache().present && savedSeries.valuesCache().present,
                       "P0V save-time option materializes all chart caches in output");
        test.checkEqual(savedSeries.valuesCache().points[1].value, std::string("2"), "P0V save-time synchronized numeric cache reloads");
    }
    std::filesystem::remove(path);
}

void testImportedChartRemoveAndAppend(TestContext& test) {
    const auto source = std::filesystem::path(XLPP_TEST_SOURCE_DIR) / "fixtures/openpyxl/image_chart.xlsx";
    const auto before = xlpp::internal::ZipArchive::open(source);
    const auto originalImage = before.get("xl/media/image1.png");

    xlpp::Workbook workbook;
    workbook.load(source);
    auto* sheet = workbook.worksheet("Objects");
    const auto& importedCharts = static_cast<const xlpp::Worksheet&>(*sheet).charts();
    test.checkEqual(importedCharts.size(), std::size_t{1}, "Remove/append fixture starts with one imported chart");
    const auto oldStableId = importedCharts.front().stableId();
    test.checkTrue(sheet->removeChart(oldStableId), "Imported chart removed by stable ID");
    test.checkTrue(!sheet->removeChart(oldStableId), "Removed stable chart ID cannot be removed twice");

    xlpp::Chart replacement(xlpp::Chart::Type::Line);
    replacement.setTitle("Replacement chart");
    replacement.setXAxisTitle("Category");
    replacement.setYAxisTitle("Amount");
    replacement.setWidth(300);
    replacement.setHeight(180);
    auto& replacementSeries = replacement.addSeries(xlpp::ChartSeries("Replacement series"));
    replacementSeries.setCategoriesReference("'Objects'!$A$2:$A$4");
    replacementSeries.setValuesReference("'Objects'!$B$2:$B$4");
    sheet->addChart(std::move(replacement));

    const auto output = std::filesystem::temp_directory_path() / "xlpp_p0i_chart_remove_append.xlsx";
    workbook.save(output);
    const auto after = xlpp::internal::ZipArchive::open(output);
    test.checkTrue(!after.contains("xl/charts/chart1.xml"), "Removed imported chart part is cleaned up");
    test.checkTrue(after.contains("xl/charts/chart2.xml"), "Appended replacement chart receives collision-free part ID");
    test.checkEqual(after.get("xl/media/image1.png"), originalImage, "Sibling imported image remains byte-identical after chart remove/append");
    const auto drawingRels = after.get("xl/drawings/_rels/drawing1.xml.rels");
    test.checkTrue(drawingRels.find("../charts/chart2.xml") != std::string::npos, "Preserved drawing points to appended chart part");
    test.checkTrue(drawingRels.find("../charts/chart1.xml") == std::string::npos, "Removed chart relationship is deleted");
    const auto contentTypes = after.get("[Content_Types].xml");
    test.checkTrue(contentTypes.find("/xl/charts/chart2.xml") != std::string::npos, "Appended chart has content-type override");
    test.checkTrue(contentTypes.find("/xl/charts/chart1.xml") == std::string::npos, "Removed chart content-type override cleaned up");

    const auto graph = xlpp::internal::RelationshipGraph::fromArchive(after);
    test.checkTrue(graph.validate().ok(), "Remove/append chart package graph validates");
    test.checkEqual(graph.objectInventory().charts, std::size_t{1}, "Remove/append keeps one visible chart");
    test.checkEqual(graph.objectInventory().images, std::size_t{1}, "Remove/append retains one sibling image");

    xlpp::Workbook reloaded;
    reloaded.load(output);
    const auto* reloadedSheet = reloaded.worksheet("Objects");
    test.checkEqual(static_cast<const xlpp::Worksheet&>(*reloadedSheet).charts().size(), std::size_t{1}, "Replacement chart reloads as one imported chart");
    test.checkEqual(static_cast<const xlpp::Worksheet&>(*reloadedSheet).charts().front().title(), std::string("Replacement chart"), "Replacement chart title survives reload");
    test.checkEqual(static_cast<int>(static_cast<const xlpp::Worksheet&>(*reloadedSheet).charts().front().type()),
                    static_cast<int>(xlpp::Chart::Type::Line), "Replacement chart type survives reload");

    const auto output2 = std::filesystem::temp_directory_path() / "xlpp_p0i_chart_remove_append_resave.xlsx";
    workbook.save(output2);
    const auto graph2 = xlpp::internal::RelationshipGraph::fromArchive(xlpp::internal::ZipArchive::open(output2));
    test.checkTrue(graph2.validate().ok(), "Repeated save of remove/append chart stays graph-clean");
    test.checkEqual(graph2.objectInventory().charts, std::size_t{1}, "Repeated save does not duplicate appended chart");
    std::filesystem::remove(output);
    std::filesystem::remove(output2);
}

void testChartTypeNameMap(TestContext& test) {
    test.checkEqual(xlpp::Chart::typeName(xlpp::Chart::Type::Bar), std::string("barChart"), "Bar standard");
    test.checkEqual(xlpp::Chart::typeName(xlpp::Chart::Type::Bar, xlpp::Chart::Grouping::Stacked), std::string("barStacked"), "Bar stacked");
    test.checkEqual(xlpp::Chart::typeName(xlpp::Chart::Type::Bar, xlpp::Chart::Grouping::PercentStacked), std::string("barPercentStacked"), "Bar percent stacked");
    test.checkEqual(xlpp::Chart::typeName(xlpp::Chart::Type::Line), std::string("lineChart"), "Line standard");
    test.checkEqual(xlpp::Chart::typeName(xlpp::Chart::Type::Line, xlpp::Chart::Grouping::Stacked), std::string("lineStacked"), "Line stacked");
    test.checkEqual(xlpp::Chart::typeName(xlpp::Chart::Type::Pie), std::string("pieChart"), "Pie");
    test.checkEqual(xlpp::Chart::typeName(xlpp::Chart::Type::Scatter), std::string("scatterChart"), "Scatter");
    test.checkEqual(xlpp::Chart::typeName(xlpp::Chart::Type::Doughnut), std::string("doughnutChart"), "Doughnut");
    test.checkEqual(xlpp::Chart::typeName(xlpp::Chart::Type::Radar), std::string("radarChart"), "Radar");
    test.checkEqual(xlpp::Chart::typeName(xlpp::Chart::Type::Area), std::string("areaChart"), "Area standard");
    test.checkEqual(xlpp::Chart::typeName(xlpp::Chart::Type::Area, xlpp::Chart::Grouping::Stacked), std::string("areaStacked"), "Area stacked");
    test.checkEqual(xlpp::Chart::typeName(xlpp::Chart::Type::Bubble), std::string("bubbleChart"), "Bubble");
    test.checkEqual(xlpp::Chart::typeName(xlpp::Chart::Type::Stock), std::string("stockChart"), "Stock");
    test.checkEqual(xlpp::Chart::typeName(xlpp::Chart::Type::PieOfPie), std::string("ofPieChart"), "Pie-of-Pie");
    test.checkEqual(xlpp::Chart::typeName(xlpp::Chart::Type::BarOfPie), std::string("ofPieChart"), "Bar-of-Pie");
    test.checkEqual(xlpp::Chart::typeName(xlpp::Chart::Type::Bar3D), std::string("bar3DChart"), "Bar3D");
    test.checkEqual(xlpp::Chart::typeName(xlpp::Chart::Type::Line3D), std::string("line3DChart"), "Line3D");
    test.checkEqual(xlpp::Chart::typeName(xlpp::Chart::Type::Area3D), std::string("area3DChart"), "Area3D");
    test.checkEqual(xlpp::Chart::typeName(xlpp::Chart::Type::Pie3D), std::string("pie3DChart"), "Pie3D");
    test.checkEqual(xlpp::Chart::typeName(xlpp::Chart::Type::Surface), std::string("surfaceChart"), "Surface");
    test.checkEqual(xlpp::Chart::typeName(xlpp::Chart::Type::Surface3D), std::string("surface3DChart"), "Surface3D");
}

void testChartAndPivotAdvancedModel(TestContext& test) {
    xlpp::Workbook workbook;
    auto& sheet = workbook.addWorksheet("Models");
    xlpp::Chart chart(xlpp::Chart::Type::Area);
    chart.setGrouping(xlpp::Chart::Grouping::Stacked);
    chart.setTitle("Area"); chart.setXAxisTitle("X"); chart.setYAxisTitle("Y");
    chart.setStyle("10"); chart.setWidth(700); chart.setHeight(300);
    chart.setShowLegend(false); chart.setLegendPosition("b");
    xlpp::ChartSeries series("Series");
    series.reference("Models", "$B$2:$B$4");
    series.categories("Models", "$A$2:$A$4");
    chart.addSeries(std::move(series));
    sheet.addChart(std::move(chart));
    sheet.chart(0).series()[0].setTitle("Renamed");
    const auto& constSheet = static_cast<const xlpp::Worksheet&>(sheet);
    test.checkEqual(static_cast<int>(constSheet.chart(0).grouping()), static_cast<int>(xlpp::Chart::Grouping::Stacked), "Chart grouping setter");
    test.checkEqual(constSheet.chart(0).series()[0].title(), std::string("Renamed"), "Mutable chart series accessor");
    test.checkEqual(constSheet.chart(0).series()[0].valuesReference(), std::string("='Models'!$B$2:$B$4"), "Chart value reference helper");
    test.checkEqual(constSheet.chart(0).series()[0].categoriesReference(), std::string("='Models'!$A$2:$A$4"), "Chart category reference helper");
    test.checkTrue(!constSheet.chart(0).showLegend(), "Chart legend visibility setter");
    test.checkEqual(constSheet.charts().size(), std::size_t{1}, "Const charts accessor");

    xlpp::PivotTable pivot("Pivot");
    pivot.setName("PivotRenamed"); pivot.setLocation("H2");
    pivot.cache().setCacheId(3); pivot.cache().setSourceData("Models!A1:C4");
    pivot.cache().setFields({"Region", "Quarter", "Amount"});
    pivot.cache().setRecords({{"East", "Q1", "10"}, {"West", "Q2", "20"}});
    test.checkEqual(pivot.cache().cacheId(), 3, "Pivot cache ID getter");
    test.checkEqual(pivot.cache().records().size(), std::size_t{2}, "Pivot setRecords stores records");
    pivot.cache().clearRecords();
    test.checkTrue(pivot.cache().records().empty(), "Pivot clearRecords clears records");
    pivot.cache().addRecord({"East", "Q1", "10"});
    auto& row = pivot.addRowField("Region"); row.setName("Region"); row.setShowAll(true); row.setSortType(1);
    pivot.addColumnField("Quarter");
    pivot.addPageField("Quarter");
    pivot.addDataField("Amount", "average");
    test.checkEqual(pivot.name(), std::string("PivotRenamed"), "Pivot name setter");
    test.checkEqual(row.axis(), std::string("axisRow"), "Pivot row axis getter");
    test.checkTrue(row.showAll(), "Pivot showAll setter");
    test.checkEqual(row.sortType(), 1, "Pivot sort setter");
    test.checkEqual(pivot.pageFields().size(), std::size_t{1}, "Mutable page fields accessor");
    test.checkEqual(pivot.dataFields().size(), std::size_t{1}, "Mutable data fields accessor");
    sheet.addPivotTable(std::move(pivot));
    test.checkEqual(constSheet.pivotTables().size(), std::size_t{1}, "Const pivot table accessor");
}

void testExcelChartFamilyGeneration(TestContext& test) {
    const auto path = std::filesystem::temp_directory_path() / "xlpp_excel_chart_families.xlsx";
    xlpp::Workbook workbook;
    auto& data = workbook.addWorksheet("Data");
    data.cell("A1").setValue("Category"); data.cell("B1").setValue("Primary"); data.cell("C1").setValue("Secondary"); data.cell("D1").setValue("Bubble");
    for (int i = 0; i < 5; ++i) {
        data.cell(static_cast<std::size_t>(i + 2), 1).setValue(std::string("Item ") + std::to_string(i + 1));
        data.cell(static_cast<std::size_t>(i + 2), 2).setValue(static_cast<double>((i + 1) * 10));
        data.cell(static_cast<std::size_t>(i + 2), 3).setValue(static_cast<double>(55 - i * 5));
        data.cell(static_cast<std::size_t>(i + 2), 4).setValue(static_cast<double>((i + 1) * 3));
    }
    auto& sheet = workbook.addWorksheet("Charts");

    // Generated combo charts are modeled as multiple plots, not as a fake
    // standalone OOXML chart type. This also covers Excel's custom-combo UI.
    xlpp::Chart combo(xlpp::Chart::Type::Bar);
    combo.setTitle("Column + line combo");
    auto& comboA = combo.addSeries(xlpp::ChartSeries("Primary"));
    comboA.setCategoriesReference("'Data'!$A$2:$A$6"); comboA.setValuesReference("'Data'!$B$2:$B$6");
    auto& comboB = combo.addSeries(xlpp::ChartSeries("Secondary"));
    comboB.setCategoriesReference("'Data'!$A$2:$A$6"); comboB.setValuesReference("'Data'!$C$2:$C$6"); comboB.setSmooth(true);
    combo.addPlot(xlpp::Chart::Type::Bar, 0, 1, false).grouping = xlpp::Chart::Grouping::Clustered;
    combo.addPlot(xlpp::Chart::Type::Line, 1, 1, true).grouping = xlpp::Chart::Grouping::Standard;
    sheet.addChart(std::move(combo));

    xlpp::Chart scatter(xlpp::Chart::Type::Scatter);
    auto& scatterSeries = scatter.addSeries(xlpp::ChartSeries("XY"));
    scatterSeries.setCategoriesReference("'Data'!$B$2:$B$6"); scatterSeries.setValuesReference("'Data'!$C$2:$C$6"); scatterSeries.setSmooth(true);
    scatter.primaryPlot().scatterStyle = xlpp::Chart::ScatterStyle::SmoothMarker;
    sheet.addChart(std::move(scatter));

    xlpp::Chart bubble(xlpp::Chart::Type::Bubble);
    auto& bubbleSeries = bubble.addSeries(xlpp::ChartSeries("Bubble"));
    bubbleSeries.setCategoriesReference("'Data'!$B$2:$B$6"); bubbleSeries.setValuesReference("'Data'!$C$2:$C$6"); bubbleSeries.setBubbleSizeReference("'Data'!$D$2:$D$6");
    bubble.primaryPlot().hasBubbleScale = true; bubble.primaryPlot().bubbleScale = 140; bubble.primaryPlot().showNegativeBubbles = true;
    sheet.addChart(std::move(bubble));

    const std::array<xlpp::Chart::Type, 8> modernTypes{
        xlpp::Chart::Type::Histogram, xlpp::Chart::Type::Pareto, xlpp::Chart::Type::BoxWhisker,
        xlpp::Chart::Type::Waterfall, xlpp::Chart::Type::Funnel, xlpp::Chart::Type::Treemap,
        xlpp::Chart::Type::Sunburst, xlpp::Chart::Type::FilledMap
    };
    for (const auto type : modernTypes) {
        xlpp::Chart chart(type);
        chart.setTitle(xlpp::Chart::typeName(type));
        auto& series = chart.addSeries(xlpp::ChartSeries("Series"));
        series.setCategoriesReference("'Data'!$A$2:$A$6");
        series.setValuesReference("'Data'!$B$2:$B$6");
        if (type == xlpp::Chart::Type::Histogram || type == xlpp::Chart::Type::Pareto) {
            auto& plot = chart.primaryPlot(); plot.histogramAutomaticBins = false; plot.histogramBinCount = 4;
            plot.histogramHasUnderflow = true; plot.histogramUnderflow = 10; plot.histogramHasOverflow = true; plot.histogramOverflow = 50;
        } else if (type == xlpp::Chart::Type::BoxWhisker) {
            auto& plot = chart.primaryPlot(); plot.boxWhiskerShowMeanLine = true; plot.boxWhiskerQuartileInclusive = true;
        } else if (type == xlpp::Chart::Type::Waterfall) {
            chart.primaryPlot().waterfallShowConnectorLines = true;
        }
        sheet.addChart(std::move(chart));
    }

    workbook.save(path);
    const auto archive = xlpp::internal::ZipArchive::open(path);
    const auto contentTypes = archive.get("[Content_Types].xml");
    test.checkTrue(contentTypes.find("application/vnd.ms-office.chartex+xml") != std::string::npos,
                   "Modern Excel charts use the ChartEx content type");
    const auto drawingXml = archive.get("xl/drawings/drawing1.xml");
    const auto drawingRels = archive.get("xl/drawings/_rels/drawing1.xml.rels");
    test.checkTrue(drawingXml.find("http://schemas.microsoft.com/office/drawing/2014/chartex") != std::string::npos,
                   "Drawing contains ChartEx graphic-data references");
    test.checkTrue(drawingRels.find("/2014/relationships/chartEx") != std::string::npos,
                   "Drawing relationships use the ChartEx relationship type");

    const auto comboXml = archive.get("xl/charts/chart1.xml");
    test.checkTrue(comboXml.find("<c:barChart>") != std::string::npos && comboXml.find("<c:lineChart>") != std::string::npos,
                   "Generated combo chart serializes multiple legacy plots");
    test.checkTrue(comboXml.find("<c:axId val=\"3\"/>") != std::string::npos && comboXml.find("<c:axId val=\"4\"/>") != std::string::npos,
                   "Generated combo chart serializes a secondary axis pair");
    const auto scatterXml = archive.get("xl/charts/chart2.xml");
    test.checkTrue(scatterXml.find("<c:xVal>") != std::string::npos && scatterXml.find("<c:yVal>") != std::string::npos,
                   "Scatter chart uses xVal/yVal instead of category/value payloads");
    const auto firstValAxis = scatterXml.find("<c:valAx>");
    const auto secondValAxis = firstValAxis == std::string::npos ? std::string::npos : scatterXml.find("<c:valAx>", firstValAxis + 1);
    test.checkTrue(firstValAxis != std::string::npos && secondValAxis != std::string::npos, "Scatter chart emits two value axes");
    const auto bubbleXml = archive.get("xl/charts/chart3.xml");
    test.checkTrue(bubbleXml.find("<c:bubbleSize>") != std::string::npos && bubbleXml.find("<c:bubbleScale val=\"140\"/>") != std::string::npos,
                   "Bubble chart serializes bubble-size data and plot options");

    const std::array<std::string, 8> layouts{"clusteredColumn", "paretoLine", "boxWhisker", "waterfall", "funnel", "treemap", "sunburst", "regionMap"};
    for (std::size_t i = 0; i < modernTypes.size(); ++i) {
        const auto xml = archive.get("xl/charts/chart" + std::to_string(i + 4) + ".xml");
        test.checkTrue(xml.find("http://schemas.microsoft.com/office/drawing/2014/chartex") != std::string::npos,
                       "Modern chart part uses ChartEx namespace");
        test.checkTrue(xml.find("layoutId=\"" + layouts[i] + "\"") != std::string::npos,
                       "Modern chart has the expected Excel series layout");
    }
    test.checkTrue(archive.get("xl/charts/chart4.xml").find("<cx:binCount val=\"4\"/>") != std::string::npos,
                   "Histogram ChartEx binning options serialize");
    test.checkTrue(archive.get("xl/charts/chart6.xml").find("quartileMethod=\"inclusive\"") != std::string::npos,
                   "Box-and-whisker quartile mode serializes");

    xlpp::Workbook reloaded; reloaded.load(path);
    const auto* loadedSheet = reloaded.worksheet("Charts");
    test.checkTrue(loadedSheet != nullptr, "Workbook containing all chart families reloads");
    if (loadedSheet) {
        const auto& charts = static_cast<const xlpp::Worksheet&>(*loadedSheet).charts();
        test.checkEqual(charts.size(), std::size_t{11}, "All generated chart objects reload");
        if (charts.size() == 11) {
            test.checkTrue(charts[0].combined() && charts[0].plots().size() == 2, "Generated combo plot structure reloads");
            for (std::size_t i = 0; i < modernTypes.size(); ++i)
                test.checkEqual(static_cast<int>(charts[i + 3].type()), static_cast<int>(modernTypes[i]), "Modern Excel chart type reloads semantically");
        }
    }
    std::filesystem::remove(path);
}
