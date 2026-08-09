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

namespace {
void testWorksheetRenameTransaction(TestContext& test) {
    xlpp::Workbook wb;
    auto& source = wb.addWorksheet("Data O'Brien");
    auto& report = wb.addWorksheet("Report");
    wb.addWorksheet("End");
    source.cell("A1").setValue(10.0);
    source.cell("B2").setValue(20.0);

    report.cell("A1").setFormula("='Data O''Brien'!B2+\"Data O'Brien!B2\"");
    report.cell("A2").setFormula("=SUM('Data O''Brien:End'!A1)");
    report.cell("A3").setFormula("='[External.xlsx]Data O''Brien'!A1");
    xlpp::Hyperlink internal("#'Data O''Brien'!B2");
    internal.setExternal(false);
    report.cell("A4").setHyperlink(std::move(internal));
    report.conditionalFormatting().addRule("B1:B2", xlpp::ConditionalRule::formula("'Data O''Brien'!A1>0"));
    report.dataValidations().add(xlpp::DataValidation::list("C1:C2", "'Data O''Brien'!$A$1:$A$2"));

    xlpp::Chart chart(xlpp::Chart::Type::Line);
    xlpp::ChartSeries series("Values");
    series.setCategoriesReference("='Data O''Brien'!$A$1:$A$2");
    series.setValuesReference("='Data O''Brien'!$B$1:$B$2");
    chart.addSeries(std::move(series));
    report.addChart(std::move(chart));

    xlpp::PivotTable pivot("RenamePivot");
    pivot.setLocation("D2:G10");
    pivot.cache().setSourceData("'Data O''Brien'!A1:B2");
    report.addPivotTable(std::move(pivot));
    wb.addDefinedName(xlpp::DefinedName("RenameRange", "'Data O''Brien'!$A$1:$B$2"));

    auto* sourceHandle = wb.worksheet("Data O'Brien");
    xlpp::WorksheetRenameOptions options;
    options.synchronizeChartCaches = false;
    const auto renamed = wb.renameWorksheet("Data O'Brien", "Data Archive", options);
    test.checkTrue(wb.worksheet("Data Archive") == sourceHandle,
                   "Dependency-aware rename preserves live worksheet object identity");
    test.checkTrue(wb.worksheet("Data O'Brien") == nullptr,
                   "Old worksheet name is removed after dependency-aware rename");
    test.checkTrue(renamed.formulasUpdated >= 2, "Worksheet rename reports rewritten formulas");
    test.checkTrue(renamed.definedNamesUpdated >= 1, "Worksheet rename reports rewritten defined names");
    test.checkTrue(renamed.chartReferencesUpdated >= 1, "Worksheet rename reports rewritten chart references");
    test.checkTrue(renamed.pivotReferencesUpdated >= 1, "Worksheet rename reports rewritten pivot references");
    test.checkTrue(renamed.hyperlinksUpdated >= 1, "Worksheet rename reports rewritten internal hyperlinks");

    test.checkEqual(report.cell("A1").formula(),
                    std::string("='Data Archive'!B2+\"Data O'Brien!B2\""),
                    "Rename rewrites qualified formula but preserves formula string literal");
    test.checkEqual(report.cell("A2").formula(), std::string("=SUM('Data Archive:End'!A1)"),
                    "Rename rewrites a 3-D sheet qualifier endpoint");
    test.checkEqual(report.cell("A3").formula(), std::string("='[External.xlsx]Data O''Brien'!A1"),
                    "Rename preserves external-workbook qualifiers");
    test.checkEqual(report.cell("A4").hyperlinkValue()->target(), std::string("#'Data Archive'!B2"),
                    "Rename rewrites internal hyperlink target");
    test.checkEqual(wb.definedName("RenameRange")->value(), std::string("'Data Archive'!$A$1:$B$2"),
                    "Rename rewrites defined-name formula");
    test.checkEqual(report.charts().front().series().front().valuesReference(),
                    std::string("='Data Archive'!$B$1:$B$2"),
                    "Rename rewrites chart source reference");
    test.checkEqual(report.pivotTables().front().cache().sourceData(), std::string("'Data Archive'!A1:B2"),
                    "Rename rewrites pivot source and invalidates its cache");
    test.checkEqual(report.dataValidations().items().front().formula1(),
                    std::string("'Data Archive'!$A$1:$A$2"),
                    "Rename rewrites validation formula");
    test.checkEqual(report.conditionalFormatting().entries().front().rules().front().formulas().front(),
                    std::string("'Data Archive'!A1>0"),
                    "Rename rewrites conditional-format formula");

    bool duplicateRejected = false;
    try { wb.renameWorksheet("Data Archive", "Report", options); }
    catch (const std::invalid_argument&) { duplicateRejected = true; }
    test.checkTrue(duplicateRejected, "Rename rejects case-insensitive worksheet-name collisions");
    test.checkTrue(wb.worksheet("Data Archive") == sourceHandle,
                   "Rejected rename leaves live workbook topology unchanged");

    const auto out = std::filesystem::temp_directory_path() / "xlpp_core_worksheet_rename.xlsx";
    wb.save(out);
    xlpp::Workbook loaded;
    loaded.load(out);
    test.checkTrue(loaded.worksheet("Data Archive") != nullptr, "Renamed worksheet survives save/load round-trip");
    test.checkEqual(loaded.worksheet("Report")->cell("A1").formula(),
                    std::string("='Data Archive'!B2+\"Data O'Brien!B2\""),
                    "Renamed cross-sheet formula survives OOXML round-trip");
    std::filesystem::remove(out);
}

void testFormulaCalculationEngine(TestContext& test) {
    xlpp::Workbook workbook;
    auto& data = workbook.addWorksheet("Data Sheet");
    auto& other = workbook.addWorksheet("Other");
    data.cell("A1").setValue(10.0);
    data.cell("A2").setValue(5.0);
    other.cell("A1").setValue(3.0);
    workbook.addDefinedName(xlpp::DefinedName("Rate", "2.5"));

    data.cell("B1").setFormula("A1+A2*2");
    data.cell("B2").setFormula("SUM(A1:A2)");
    data.cell("B3").setFormula("IF(B1>15,\"high\",\"low\")");
    data.cell("B4").setFormula("\"Total: \"&B2");
    data.cell("C1").setFormula("B1+Other!A1");
    data.cell("C2").setFormula("A1*Rate");
    data.cell("C3").setFormula("SUM(A1:A2)+MAX(A1:A2)-MIN(A1:A2)");
    data.cell("C4").setFormula("ROUND(10/3,2)");
    data.cell("C5").setFormula("COUNTIF(A1:A2,\">=6\")");
    data.cell("C6").setFormula("LEFT(\"abcdef\",3)&UPPER(\"xy\")");
    data.cell("C7").setFormula("'Data Sheet'!A1+'Other'!A1");
    data.cell("C8").setFormula("DATE(2026,8,8)");
    data.cell("C9").setFormula("YEAR(C8)+MONTH(C8)+DAY(C8)");

    const auto report = workbook.calculateFormulas();
    test.checkEqual(report.formulaCellsVisited, std::size_t{13}, "Formula engine visits every formula cell");
    test.checkEqual(report.unsupportedFormulas, std::size_t{0}, "Core formula set has no unsupported functions");
    test.checkEqual(report.circularReferences, std::size_t{0}, "Acyclic formula graph stays acyclic");
    test.checkNear(std::get<double>(data.cell("B1").value()), 20.0, 1e-12, "Arithmetic precedence is calculated");
    test.checkNear(std::get<double>(data.cell("B2").value()), 15.0, 1e-12, "Range SUM is calculated");
    test.checkEqual(std::get<std::string>(data.cell("B3").value()), std::string("high"), "IF and comparison are calculated");
    test.checkEqual(std::get<std::string>(data.cell("B4").value()), std::string("Total: 15"), "String concatenation uses calculated dependency");
    test.checkNear(std::get<double>(data.cell("C1").value()), 23.0, 1e-12, "Cross-sheet dependency is calculated");
    test.checkNear(std::get<double>(data.cell("C2").value()), 25.0, 1e-12, "Defined name is resolved");
    test.checkNear(std::get<double>(data.cell("C3").value()), 20.0, 1e-12, "Multiple range aggregations are calculated");
    test.checkNear(std::get<double>(data.cell("C4").value()), 3.33, 1e-12, "ROUND is calculated");
    test.checkNear(std::get<double>(data.cell("C5").value()), 1.0, 1e-12, "COUNTIF criteria are calculated");
    test.checkEqual(std::get<std::string>(data.cell("C6").value()), std::string("abcXY"), "Text functions are calculated");
    test.checkNear(std::get<double>(data.cell("C7").value()), 13.0, 1e-12, "Quoted sheet references are calculated");
    test.checkTrue(std::holds_alternative<xlpp::DateTime>(data.cell("C8").value()), "DATE returns a DateTime cached value");
    test.checkNear(std::get<double>(data.cell("C9").value()), 2042.0, 1e-12, "YEAR MONTH DAY consume calculated date dependencies");

    const auto savePath=std::filesystem::temp_directory_path()/"xlpp_formula_calculate_on_save.xlsx";
    xlpp::Workbook saveWorkbook;auto& saveSheet=saveWorkbook.addWorksheet("Calc");saveSheet.cell("A1").setValue(7.0);saveSheet.cell("A2").setFormula("A1*6");
    xlpp::SaveOptions calculateOnSave;calculateOnSave.calculateFormulasBeforeSave=true;saveWorkbook.save(savePath,calculateOnSave);
    test.checkTrue(std::holds_alternative<std::monostate>(saveSheet.cell("A2").value()),"Calculate-on-save leaves caller workbook cache untouched");
    xlpp::Workbook saveLoaded;saveLoaded.load(savePath);test.checkNear(saveLoaded.worksheet("Calc")->cell("A2").numericValueOr(-1),42.0,1e-12,"Calculate-on-save persists evaluated formula cache");
    std::filesystem::remove(savePath);

    data.cell("D1").setFormula("D2+1");
    data.cell("D2").setFormula("D1+1");
    const auto cycleReport = workbook.calculateFormulas();
    test.checkTrue(cycleReport.circularReferences >= 1, "Circular references are detected");
    test.checkTrue(data.cell("D1").isError(), "Circular formula caches an error value");

    xlpp::CalculationOptions diagnostic;
    diagnostic.updateCachedValues = false;
    data.cell("A1").setValue(100.0);
    const auto before = std::get<double>(data.cell("B1").value());
    const auto diagnosticReport = workbook.calculateFormulas(diagnostic);
    test.checkTrue(diagnosticReport.formulaCellsEvaluated >= 13, "Diagnostic calculation still evaluates formulas");
    test.checkNear(std::get<double>(data.cell("B1").value()), before, 1e-12, "Diagnostic calculation does not mutate cached values");
}

void testAdvancedFormulaCalculation(TestContext& test) {
    xlpp::Workbook workbook;
    auto& sheet=workbook.addWorksheet("Lookup");
    for(int row=1;row<=4;++row){
        sheet.cell(row,1).setValue(static_cast<double>(row));
        sheet.cell(row,3).setValue(static_cast<double>(row*10));
    }
    sheet.cell("B1").setValue("one");sheet.cell("B2").setValue("two");sheet.cell("B3").setValue("three");sheet.cell("B4").setValue("four");
    sheet.cell("D1").setValue("A");sheet.cell("D2").setValue("A");sheet.cell("D3").setValue("B");sheet.cell("D4").setValue("B");
    sheet.cell("A6").setValue(1.0);sheet.cell("B6").setValue(2.0);sheet.cell("C6").setValue(3.0);
    sheet.cell("A7").setValue(11.0);sheet.cell("B7").setValue(22.0);sheet.cell("C7").setValue(33.0);

    sheet.cell("F1").setFormula("INDEX(A1:C4,3,2)");
    sheet.cell("F2").setFormula("MATCH(3,A1:A4,0)");
    sheet.cell("F3").setFormula("VLOOKUP(2,A1:C4,3,FALSE)");
    sheet.cell("F4").setFormula("XLOOKUP(4,A1:A4,B1:B4,\"missing\")");
    sheet.cell("F5").setFormula("HLOOKUP(2,A6:C7,2,FALSE)");
    sheet.cell("F6").setFormula("COUNTIFS(D1:D4,\"A\",C1:C4,\">=15\")");
    sheet.cell("F7").setFormula("SUMIFS(C1:C4,D1:D4,\"B\",A1:A4,\">=3\")");
    sheet.cell("F8").setFormula("AVERAGEIFS(C1:C4,D1:D4,\"B\")");
    sheet.cell("F9").setFormula("COUNTIF(B1:B4,\"t*\")");
    sheet.cell("F10").setFormula("SUMPRODUCT(A1:A4,C1:C4)");
    sheet.cell("F11").setFormula("IFNA(#N/A,\"fallback\")");
    sheet.cell("F12").setFormula("ISERR(#DIV/0!)");
    sheet.cell("F13").setFormula("SEARCH(\"RE\",\"three\")");
    sheet.cell("F14").setFormula("SUBSTITUTE(\"a-b-a\",\"a\",\"x\",2)");
    sheet.cell("F15").setFormula("VALUE(\" 12.5 \")+SIGN(-5)");
    sheet.cell("F16").setFormula("TIME(1,30,0)*24");
    sheet.cell("F17").setFormula("HOUR(TIME(13,45,30))");
    sheet.cell("F18").setFormula("XLOOKUP(\"t*\",B1:B4,C1:C4,\"none\",2)");
    sheet.cell("F19").setFormula("CHOOSE(2,\"x\",\"y\",\"z\")");
    sheet.cell("F20").setFormula("MATCH(3.5,A1:A4,1)");
    sheet.cell("E2").setValue("");sheet.cell("E3").setValue("x");
    sheet.cell("F21").setFormula("MEDIAN(C1:C4)");
    sheet.cell("F22").setFormula("STDEV.P(C1:C4)");
    sheet.cell("F23").setFormula("LARGE(C1:C4,2)");
    sheet.cell("F24").setFormula("COUNTBLANK(E1:E4)");
    sheet.cell("F25").setFormula("IFS(A1=2,\"no\",A1=1,\"yes\")");
    sheet.cell("F26").setFormula("SWITCH(B2,\"one\",1,\"two\",2,0)");
    sheet.cell("F27").setFormula("XOR(TRUE,FALSE,TRUE)");
    sheet.cell("F28").setFormula("TEXTJOIN(\"-\",TRUE,B1:B3)");
    sheet.cell("F29").setFormula("MINIFS(C1:C4,D1:D4,\"B\")");
    sheet.cell("F30").setFormula("MAXIFS(C1:C4,D1:D4,\"B\")");
    sheet.cell("F31").setFormula("CEILING(2.1,1)+FLOOR(2.9,1)");
    sheet.cell("F32").setFormula("DAYS(DATE(2026,8,8),DATE(2026,8,1))");
    sheet.cell("F33").setFormula("ISNA(NA())");
    sheet.cell("F34").setFormula("PI()");
    sheet.cell("F35").setFormula("SIN(PI()/2)");
    sheet.cell("F36").setFormula("DEGREES(PI())");
    sheet.cell("F37").setFormula("SUMSQ(3,4)");
    sheet.cell("F38").setFormula("PMT(0.05/12,12,1000)");
    sheet.cell("F39").setFormula("FV(0.05/12,12,-100,0)");
    sheet.cell("F40").setFormula("PV(0.05/12,12,-100)");
    sheet.cell("F41").setFormula("NPV(0.1,500,600)");
    sheet.cell("G1").setValue(-1000.0);sheet.cell("G2").setValue(500.0);sheet.cell("G3").setValue(600.0);
    sheet.cell("F42").setFormula("IRR(G1:G3)");
    sheet.cell("F43").setFormula("ATAN2(1,1)");
    sheet.cell("F44").setFormula("AND(TRUE(),NOT(FALSE()))");
    sheet.cell("F45").setFormula("SUM({1,2,3;4,5,6})");
    sheet.cell("F46").setFormula("INDEX({1,2;3,4},2,1)");
    sheet.cell("F47").setFormula("LEN(CLEAN(CHAR(9)))");
    sheet.cell("F48").setFormula("PROPER(\"hELLO-world\")");
    sheet.cell("F49").setFormula("REPLACE(\"abcdef\",2,3,\"XY\")");
    sheet.cell("F50").setFormula("T(42)");
    sheet.cell("F51").setFormula("N(TRUE)");
    sheet.cell("F52").setFormula("CODE(CHAR(65))");
    sheet.cell("F53").setFormula("MONTH(EDATE(DATE(2026,1,31),1))");
    sheet.cell("F54").setFormula("DAY(EOMONTH(DATE(2026,2,4),0))");
    sheet.cell("F55").setFormula("WEEKDAY(DATE(2026,8,9),2)");
    sheet.cell("F56").setFormula("TRUNC(-12.345,2)");
    sheet.cell("F57").setFormula("QUOTIENT(-7,2)");
    sheet.cell("F58").setFormula("EVEN(3.1)+ODD(2.1)");
    sheet.cell("F59").setFormula("GCD(24,18)+LCM(4,6)");
    sheet.cell("F60").setFormula("MROUND(10,3)+RANK.EQ(30,C1:C4,0)");
    sheet.cell("F61").setFormula("MODE.SNGL({1,2,2,3})");
    sheet.cell("F62").setFormula("PERCENTILE.INC(C1:C4,0.25)");
    sheet.cell("F63").setFormula("PERCENTILE.EXC(C1:C4,0.25)");
    sheet.cell("F64").setFormula("QUARTILE.INC(C1:C4,3)");
    sheet.cell("F65").setFormula("QUARTILE.EXC(C1:C4,3)");
    sheet.cell("F66").setFormula("GEOMEAN(1,4)+HARMEAN(1,2,4)");
    sheet.cell("F67").setFormula("RANK.AVG(30,{40,30,30,20},0)");
    sheet.cell("F68").setFormula("WEEKNUM(DATE(2026,8,9),2)");
    sheet.cell("F69").setFormula("ISOWEEKNUM(DATE(2026,8,9))");
    sheet.cell("H1").setValue(xlpp::DateTime{2026,8,5});
    sheet.cell("F70").setFormula("NETWORKDAYS(DATE(2026,8,3),DATE(2026,8,9),H1)");
    sheet.cell("F71").setFormula("DAY(WORKDAY(DATE(2026,8,7),1))");
    sheet.cell("F72").setFormula("NETWORKDAYS.INTL(DATE(2026,8,3),DATE(2026,8,9),11)");
    sheet.cell("F73").setFormula("DAY(WORKDAY.INTL(DATE(2026,8,7),1,\"0000011\"))");
    sheet.cell("F74").setFormula("CORREL(A1:A4,C1:C4)");
    sheet.cell("F75").setFormula("COVARIANCE.P(A1:A4,C1:C4)");
    sheet.cell("F76").setFormula("SLOPE(C1:C4,A1:A4)+INTERCEPT(C1:C4,A1:A4)");
    sheet.cell("F77").setFormula("FORECAST.LINEAR(5,C1:C4,A1:A4)");
    sheet.cell("F78").setFormula("RSQ(C1:C4,A1:A4)");

    const auto report=workbook.calculateFormulas();
    test.checkEqual(report.unsupportedFormulas,std::size_t{0},"Advanced core formula matrix has no unsupported functions");
    test.checkEqual(std::get<std::string>(sheet.cell("F1").value()),std::string("three"),"INDEX preserves two-dimensional range shape");
    test.checkNear(std::get<double>(sheet.cell("F2").value()),3.0,1e-12,"MATCH exact lookup works");
    test.checkNear(std::get<double>(sheet.cell("F3").value()),20.0,1e-12,"VLOOKUP exact lookup works");
    test.checkEqual(std::get<std::string>(sheet.cell("F4").value()),std::string("four"),"XLOOKUP exact lookup works");
    test.checkNear(std::get<double>(sheet.cell("F5").value()),22.0,1e-12,"HLOOKUP exact lookup works");
    test.checkNear(std::get<double>(sheet.cell("F6").value()),1.0,1e-12,"COUNTIFS evaluates multiple criteria");
    test.checkNear(std::get<double>(sheet.cell("F7").value()),70.0,1e-12,"SUMIFS evaluates multiple criteria");
    test.checkNear(std::get<double>(sheet.cell("F8").value()),35.0,1e-12,"AVERAGEIFS evaluates multiple criteria");
    test.checkNear(std::get<double>(sheet.cell("F9").value()),2.0,1e-12,"COUNTIF supports Excel wildcards");
    test.checkNear(std::get<double>(sheet.cell("F10").value()),300.0,1e-12,"SUMPRODUCT evaluates aligned arrays");
    test.checkEqual(std::get<std::string>(sheet.cell("F11").value()),std::string("fallback"),"IFNA consumes #N/A literal");
    test.checkTrue(std::get<bool>(sheet.cell("F12").value()),"ISERR recognizes #DIV/0! literal");
    test.checkNear(std::get<double>(sheet.cell("F13").value()),3.0,1e-12,"SEARCH is case-insensitive and one-based");
    test.checkEqual(std::get<std::string>(sheet.cell("F14").value()),std::string("a-b-x"),"SUBSTITUTE instance selection works");
    test.checkNear(std::get<double>(sheet.cell("F15").value()),11.5,1e-12,"VALUE and SIGN calculate");
    test.checkNear(std::get<double>(sheet.cell("F16").value()),1.5,1e-12,"TIME produces Excel day fraction");
    test.checkNear(std::get<double>(sheet.cell("F17").value()),13.0,1e-12,"HOUR extracts time component");
    test.checkNear(std::get<double>(sheet.cell("F18").value()),20.0,1e-12,"XLOOKUP wildcard mode works");
    test.checkEqual(std::get<std::string>(sheet.cell("F19").value()),std::string("y"),"CHOOSE selects one-based argument");
    test.checkNear(std::get<double>(sheet.cell("F20").value()),3.0,1e-12,"MATCH approximate mode works on ascending numeric range");
    test.checkNear(std::get<double>(sheet.cell("F21").value()),25.0,1e-12,"MEDIAN calculates range midpoint");
    test.checkNear(std::get<double>(sheet.cell("F22").value()),std::sqrt(125.0),1e-12,"STDEV.P calculates population standard deviation");
    test.checkNear(std::get<double>(sheet.cell("F23").value()),30.0,1e-12,"LARGE selects kth value");
    test.checkNear(std::get<double>(sheet.cell("F24").value()),3.0,1e-12,"COUNTBLANK counts absent and empty-string cells");
    test.checkEqual(std::get<std::string>(sheet.cell("F25").value()),std::string("yes"),"IFS selects first true branch");
    test.checkNear(std::get<double>(sheet.cell("F26").value()),2.0,1e-12,"SWITCH selects matching pair");
    test.checkTrue(!std::get<bool>(sheet.cell("F27").value()),"XOR combines logical arguments");
    test.checkEqual(std::get<std::string>(sheet.cell("F28").value()),std::string("one-two-three"),"TEXTJOIN consumes range values");
    test.checkNear(std::get<double>(sheet.cell("F29").value()),30.0,1e-12,"MINIFS returns matching minimum");
    test.checkNear(std::get<double>(sheet.cell("F30").value()),40.0,1e-12,"MAXIFS returns matching maximum");
    test.checkNear(std::get<double>(sheet.cell("F31").value()),5.0,1e-12,"CEILING and FLOOR calculate significance rounding");
    test.checkNear(std::get<double>(sheet.cell("F32").value()),7.0,1e-12,"DAYS subtracts Excel date serials");
    test.checkTrue(std::get<bool>(sheet.cell("F33").value()),"NA and ISNA preserve Excel error semantics");
    test.checkNear(std::get<double>(sheet.cell("F34").value()),3.14159265358979323846,1e-14,"PI returns the mathematical constant");
    test.checkNear(std::get<double>(sheet.cell("F35").value()),1.0,1e-12,"Trigonometric functions calculate radians");
    test.checkNear(std::get<double>(sheet.cell("F36").value()),180.0,1e-12,"DEGREES converts radians");
    test.checkNear(std::get<double>(sheet.cell("F37").value()),25.0,1e-12,"SUMSQ aggregates squared numeric arguments");
    test.checkNear(std::get<double>(sheet.cell("F38").value()),-85.60748178846747,1e-9,"PMT calculates periodic payment");
    test.checkNear(std::get<double>(sheet.cell("F39").value()),1227.8855491615914,1e-9,"FV calculates future value");
    test.checkNear(std::get<double>(sheet.cell("F40").value()),1168.1222004298158,1e-9,"PV calculates present value");
    test.checkNear(std::get<double>(sheet.cell("F41").value()),950.4132231404958,1e-9,"NPV discounts periodic cash flows");
    test.checkNear(std::get<double>(sheet.cell("F42").value()),0.0639410298049854,1e-9,"IRR solves a mixed-sign cash-flow series");
    test.checkNear(std::get<double>(sheet.cell("F43").value()),3.14159265358979323846/4.0,1e-12,"ATAN2 follows Excel x/y argument order");
    test.checkTrue(std::get<bool>(sheet.cell("F44").value()),"TRUE and FALSE function forms calculate");
    test.checkNear(std::get<double>(sheet.cell("F45").value()),21.0,1e-12,"Array constants preserve row/column separators in aggregate functions");
    test.checkNear(std::get<double>(sheet.cell("F46").value()),3.0,1e-12,"Two-dimensional array constants work with INDEX");
    test.checkNear(std::get<double>(sheet.cell("F47").value()),0.0,1e-12,"CLEAN removes ASCII control characters");
    test.checkEqual(std::get<std::string>(sheet.cell("F48").value()),std::string("Hello-World"),"PROPER normalizes word casing");
    test.checkEqual(std::get<std::string>(sheet.cell("F49").value()),std::string("aXYef"),"REPLACE uses one-based character offsets");
    test.checkEqual(std::get<std::string>(sheet.cell("F50").value()),std::string{},"T returns an empty string for non-text values");
    test.checkNear(std::get<double>(sheet.cell("F51").value()),1.0,1e-12,"N coerces logical values");
    test.checkNear(std::get<double>(sheet.cell("F52").value()),65.0,1e-12,"CHAR and CODE round-trip an ANSI byte");
    test.checkNear(std::get<double>(sheet.cell("F53").value()),2.0,1e-12,"EDATE clamps end-of-month dates");
    test.checkNear(std::get<double>(sheet.cell("F54").value()),28.0,1e-12,"EOMONTH returns the target month end");
    test.checkNear(std::get<double>(sheet.cell("F55").value()),7.0,1e-12,"WEEKDAY supports Monday-based numbering");
    test.checkNear(std::get<double>(sheet.cell("F56").value()),-12.34,1e-12,"TRUNC rounds toward zero");
    test.checkNear(std::get<double>(sheet.cell("F57").value()),-3.0,1e-12,"QUOTIENT truncates signed division");
    test.checkNear(std::get<double>(sheet.cell("F58").value()),7.0,1e-12,"EVEN and ODD round away from zero");
    test.checkNear(std::get<double>(sheet.cell("F59").value()),18.0,1e-12,"GCD and LCM evaluate integer arguments");
    test.checkNear(std::get<double>(sheet.cell("F60").value()),11.0,1e-12,"MROUND and RANK.EQ follow Excel ordering semantics");
    test.checkNear(std::get<double>(sheet.cell("F61").value()),2.0,1e-12,"MODE.SNGL returns the lowest most-frequent value");
    test.checkNear(std::get<double>(sheet.cell("F62").value()),17.5,1e-12,"PERCENTILE.INC interpolates inclusive ranks");
    test.checkNear(std::get<double>(sheet.cell("F63").value()),12.5,1e-12,"PERCENTILE.EXC interpolates exclusive ranks");
    test.checkNear(std::get<double>(sheet.cell("F64").value()),32.5,1e-12,"QUARTILE.INC maps quartiles to inclusive percentiles");
    test.checkNear(std::get<double>(sheet.cell("F65").value()),37.5,1e-12,"QUARTILE.EXC maps quartiles to exclusive percentiles");
    test.checkNear(std::get<double>(sheet.cell("F66").value()),2.0+12.0/7.0,1e-12,"GEOMEAN and HARMEAN evaluate positive samples");
    test.checkNear(std::get<double>(sheet.cell("F67").value()),2.5,1e-12,"RANK.AVG averages tied rank positions");
    test.checkNear(std::get<double>(sheet.cell("F68").value()),32.0,1e-12,"WEEKNUM supports Monday-based week numbering");
    test.checkNear(std::get<double>(sheet.cell("F69").value()),32.0,1e-12,"ISOWEEKNUM follows ISO-8601 week numbering");
    test.checkNear(std::get<double>(sheet.cell("F70").value()),4.0,1e-12,"NETWORKDAYS excludes weekends and holiday ranges");
    test.checkNear(std::get<double>(sheet.cell("F71").value()),10.0,1e-12,"WORKDAY advances across a weekend");
    test.checkNear(std::get<double>(sheet.cell("F72").value()),6.0,1e-12,"NETWORKDAYS.INTL supports numeric weekend codes");
    test.checkNear(std::get<double>(sheet.cell("F73").value()),10.0,1e-12,"WORKDAY.INTL supports seven-character weekend masks");
    test.checkNear(std::get<double>(sheet.cell("F74").value()),1.0,1e-12,"CORREL calculates Pearson correlation");
    test.checkNear(std::get<double>(sheet.cell("F75").value()),12.5,1e-12,"COVARIANCE.P calculates population covariance");
    test.checkNear(std::get<double>(sheet.cell("F76").value()),10.0,1e-12,"SLOPE and INTERCEPT calculate linear regression coefficients");
    test.checkNear(std::get<double>(sheet.cell("F77").value()),50.0,1e-12,"FORECAST.LINEAR projects a value from linear regression");
    test.checkNear(std::get<double>(sheet.cell("F78").value()),1.0,1e-12,"RSQ returns the coefficient of determination");
}

void testFormulaDependencyGraph(TestContext& test) {
    xlpp::Workbook workbook;
    auto& data = workbook.addWorksheet("Data");
    data.cell("A1").setValue(2.0); data.cell("B1").setValue(3.0);
    data.cell("A2").setValue("East"); data.cell("B2").setValue(10.0);
    data.cell("A3").setValue("West"); data.cell("B3").setValue(20.0);
    auto& table = data.addTable("Sales", "A1:B3"); table.addColumn("Region"); table.addColumn("Amount");
    xlpp::DefinedName rate("Rate", "'Data'!$A$1"); workbook.addDefinedName(std::move(rate));
    auto& calc = workbook.addWorksheet("Calc");
    calc.cell("A1").setFormula("Data!A1+Data!B1");
    calc.cell("A2").setFormula("SUM(Data!A1:B1)");
    calc.cell("A3").setFormula("Rate*2");
    calc.cell("A4").setFormula("SUM(Sales[Amount])");
    calc.cell("A5").setFormula("INDIRECT(\"Data!A1\")");

    const auto graph = workbook.dependencyGraph();
    test.checkEqual(graph.report().formulaCells, std::size_t{5}, "Dependency graph visits every formula cell");
    test.checkTrue(graph.report().cellOrRangeEdges >= 5, "Dependency graph records direct, named and table-backed ranges");
    test.checkEqual(graph.report().definedNameEdges, std::size_t{1}, "Dependency graph records defined-name dependency");
    test.checkEqual(graph.report().tableEdges, std::size_t{1}, "Dependency graph records structured-table dependency");
    test.checkEqual(graph.report().volatileReferences, std::size_t{1}, "Dependency graph marks INDIRECT as volatile");
    test.checkTrue(graph.dependsOn("Calc", "A1", "Data", "A1"), "Direct cross-sheet dependency is queryable");
    test.checkTrue(graph.dependsOn("Calc", "A2", "Data", "B1"), "Range dependency contains interior cell");
    test.checkTrue(graph.dependsOn("Calc", "A3", "Data", "A1"), "Defined name expands to its concrete cell dependency");
    test.checkTrue(graph.dependsOn("Calc", "A4", "Data", "B2"), "Structured table expands to a concrete table range dependency");
    const auto precedents = graph.precedentsOf("Calc", "A3");
    test.checkTrue(std::any_of(precedents.begin(), precedents.end(), [](const auto& edge) { return edge.kind == xlpp::FormulaDependencyKind::DefinedName && edge.symbol == "Rate"; }), "Precedent query retains the defined-name symbol");
}

void testDynamicArraysAndStructuredReferences(TestContext& test) {
    xlpp::Workbook workbook;
    auto& sheet = workbook.addWorksheet("Modern");
    sheet.cell("A1").setValue("Region"); sheet.cell("B1").setValue("Amount");
    sheet.cell("A2").setValue("East"); sheet.cell("B2").setValue(10.0); sheet.cell("C2").setValue(false);
    sheet.cell("A3").setValue("West"); sheet.cell("B3").setValue(30.0); sheet.cell("C3").setValue(true);
    sheet.cell("A4").setValue("East"); sheet.cell("B4").setValue(20.0); sheet.cell("C4").setValue(true);
    auto& table = sheet.addTable("Sales", "A1:B4");
    table.addColumn("Region"); table.addColumn("Amount");

    sheet.cell("D1").setFormula("SUM(Sales[Amount])");
    sheet.cell("D2").setFormula("Sales[@Amount]*2");
    sheet.cell("D3").setFormula("SUM(Sales[#Data])");
    sheet.cell("D4").setFormula("SUM(Sales[[Region]:[Amount]])");
    sheet.cell("F1").setFormula("SEQUENCE(2,3,1,1)");
    sheet.cell("F4").setFormula("SORT(A2:B4,2,-1)");
    sheet.cell("I1").setFormula("UNIQUE(A2:A4)");
    sheet.cell("K1").setFormula("FILTER(A2:B4,C2:C4)");
    sheet.cell("N1").setFormula("TRANSPOSE({1,2,3})");
    sheet.cell("P1").setFormula("TAKE({1,2,3;4,5,6;7,8,9},2,2)");
    sheet.cell("S1").setFormula("DROP({1,2,3;4,5,6;7,8,9},1,1)");
    sheet.cell("V1").setFormula("CHOOSECOLS({1,2,3;4,5,6},3,1)");
    sheet.cell("Y1").setFormula("CHOOSEROWS({1,2;3,4;5,6},3,1)");
    sheet.cell("AB1").setFormula("HSTACK({1;2},{3;4})");
    sheet.cell("AE1").setFormula("VSTACK({1,2},{3,4})");
    sheet.cell("AH1").setFormula("TOROW({1,2;3,4})");
    sheet.cell("AM1").setFormula("TOCOL({1,2;3,4})");
    sheet.cell("AO1").setFormula("INDIRECT(\"B3\")");
    sheet.cell("AP1").setFormula("SUM(OFFSET(B2,1,0,2,1))");
    sheet.cell("AQ1").setFormula("ROWS(A2:B4)+COLUMNS(A2:B4)");
    sheet.cell("AR1").setFormula("LET(x,2,y,x+3,x*y)");
    sheet.cell("AS1").setFormula("ROW(B3)+COLUMN(B3)");
    sheet.cell("AT1").setFormula("ADDRESS(3,2,4)");
    sheet.cell("AU1").setFormula("ADDRESS(3,2,1,TRUE,\"Data Sheet\")");

    const auto report = workbook.calculateFormulas();
    test.checkEqual(report.unsupportedFormulas, std::size_t{0}, "Modern formula matrix has no unsupported functions");
    test.checkTrue(report.dynamicArraysSpilled >= 7, "Dynamic array functions materialize spill ranges");
    test.checkTrue(report.structuredReferencesResolved >= 4, "Structured table references resolve through table metadata");
    test.checkNear(sheet.cell("D1").numericValueOr(-1), 60.0, 1e-12, "Structured single-column reference excludes table header");
    test.checkNear(sheet.cell("D2").numericValueOr(-1), 20.0, 1e-12, "Structured current-row selector resolves the formula row");
    test.checkNear(sheet.cell("D3").numericValueOr(-1), 60.0, 1e-12, "Structured #Data reference resolves the table body");
    test.checkNear(sheet.cell("D4").numericValueOr(-1), 60.0, 1e-12, "Structured multi-column selector resolves a rectangular table slice");
    test.checkNear(sheet.cell("F1").numericValueOr(-1), 1.0, 1e-12, "SEQUENCE top-left cache is calculated");
    test.checkNear(sheet.cell("H2").numericValueOr(-1), 6.0, 1e-12, "SEQUENCE spills a two-dimensional result");
    test.checkTrue(sheet.cell("F1").formulaMetadata().type() == xlpp::FormulaType::DynamicArray, "Calculated spill formula is marked DynamicArray");
    test.checkEqual(sheet.cell("F1").formulaMetadata().reference(), std::string("F1:H2"), "Dynamic array metadata records spill extent");
    test.checkEqual(sheet.cell("F4").stringValueOr(""), std::string("West"), "SORT spills rows by selected column");
    test.checkNear(sheet.cell("G4").numericValueOr(-1), 30.0, 1e-12, "SORT preserves row values");
    test.checkEqual(sheet.cell("I1").stringValueOr(""), std::string("East"), "UNIQUE keeps first occurrence");
    test.checkEqual(sheet.cell("I2").stringValueOr(""), std::string("West"), "UNIQUE spills distinct value");
    test.checkEqual(sheet.cell("K1").stringValueOr(""), std::string("West"), "FILTER applies vertical include mask");
    test.checkNear(sheet.cell("L2").numericValueOr(-1), 20.0, 1e-12, "FILTER preserves selected row shape");
    test.checkNear(sheet.cell("N3").numericValueOr(-1), 3.0, 1e-12, "TRANSPOSE swaps array dimensions");
    test.checkNear(sheet.cell("Q2").numericValueOr(-1), 5.0, 1e-12, "TAKE selects top-left array subset");
    test.checkNear(sheet.cell("T2").numericValueOr(-1), 9.0, 1e-12, "DROP removes leading rows and columns");
    test.checkNear(sheet.cell("V1").numericValueOr(-1), 3.0, 1e-12, "CHOOSECOLS selects requested columns in requested order");
    test.checkNear(sheet.cell("W2").numericValueOr(-1), 4.0, 1e-12, "CHOOSECOLS preserves selected matrix shape");
    test.checkNear(sheet.cell("Y1").numericValueOr(-1), 5.0, 1e-12, "CHOOSEROWS selects requested rows in requested order");
    test.checkNear(sheet.cell("Z2").numericValueOr(-1), 2.0, 1e-12, "CHOOSEROWS preserves row values");
    test.checkNear(sheet.cell("AC2").numericValueOr(-1), 4.0, 1e-12, "HSTACK combines arrays horizontally");
    test.checkNear(sheet.cell("AF2").numericValueOr(-1), 4.0, 1e-12, "VSTACK combines arrays vertically");
    test.checkNear(sheet.cell("AK1").numericValueOr(-1), 4.0, 1e-12, "TOROW flattens into a horizontal spill");
    test.checkNear(sheet.cell("AM4").numericValueOr(-1), 4.0, 1e-12, "TOCOL flattens into a vertical spill");
    test.checkNear(sheet.cell("AO1").numericValueOr(-1), 30.0, 1e-12, "INDIRECT resolves A1 text references");
    test.checkNear(sheet.cell("AP1").numericValueOr(-1), 50.0, 1e-12, "OFFSET returns a range with source-reference provenance");
    test.checkNear(sheet.cell("AQ1").numericValueOr(-1), 5.0, 1e-12, "ROWS and COLUMNS inspect range dimensions");
    test.checkNear(sheet.cell("AR1").numericValueOr(-1), 10.0, 1e-12, "LET binds multiple local variables and evaluates dependent bindings");
    test.checkNear(sheet.cell("AS1").numericValueOr(-1), 5.0, 1e-12, "ROW and COLUMN inspect reference provenance");
    test.checkEqual(sheet.cell("AT1").stringValueOr(""), std::string("B3"), "ADDRESS supports relative A1 references");
    test.checkEqual(sheet.cell("AU1").stringValueOr(""), std::string("'Data Sheet'!$B$3"), "ADDRESS can qualify an absolute A1 reference with a sheet name");

    xlpp::Workbook blocked;
    auto& bs = blocked.addWorksheet("Blocked");
    bs.cell("A1").setFormula("SEQUENCE(1,3)"); bs.cell("B1").setValue("occupied");
    const auto blockedReport = blocked.calculateFormulas();
    test.checkEqual(blockedReport.spillConflicts, std::size_t{1}, "Blocked dynamic array reports a spill conflict");
    test.checkTrue(bs.cell("A1").error() == xlpp::CellError::Spill, "Blocked spill caches #SPILL!");
    test.checkEqual(bs.cell("B1").stringValueOr(""), std::string("occupied"), "Blocked spill never overwrites existing data");
}

void testStandardEncryptionLibreOfficeFixture(TestContext& test) {
    const auto path=std::filesystem::path(XLPP_TEST_SOURCE_DIR)/"fixtures"/"libreoffice"/"standard_aes128_sha1_encrypted.xlsx";
    test.checkTrue(std::filesystem::exists(path),"LibreOffice Standard-encryption fixture exists");
    const auto info=xlpp::inspectOfficeEncryption(path);
    test.checkTrue(info.encrypted&&info.supported,"Standard AES/SHA-1 encryption is recognized");
    test.checkTrue(info.mode==xlpp::OfficeEncryptionMode::StandardAesSha1,"Inspector distinguishes Standard from Agile encryption");
    test.checkEqual(info.keyBits,std::uint32_t{128},"LibreOffice fixture uses AES-128");
    test.checkEqual(info.hashAlgorithm,std::string("SHA1"),"LibreOffice Standard fixture uses SHA-1");
    test.checkEqual(info.spinCount,std::uint32_t{50000},"Standard password derivation reports 50,000 iterations");
    bool wrong=false;try{xlpp::Workbook w;xlpp::LoadOptions o;o.password="wrong";w.load(path,o);}catch(const std::exception&){wrong=true;}
    test.checkTrue(wrong,"Standard verifier rejects wrong password");
    xlpp::Workbook loaded;xlpp::LoadOptions options;options.password="LibreOfficePass!";loaded.load(path,options);auto* sheet=loaded.worksheet("Sheet1");
    test.checkTrue(sheet!=nullptr,"XL++ decrypts independently generated LibreOffice Standard workbook");
    test.checkEqual(sheet->cell("A1").stringValueOr(""),std::string("from libreoffice"),"Standard decrypted text matches fixture");
    test.checkNear(sheet->cell("A2").numericValueOr(-1),77.0,1e-12,"Standard decrypted numeric value matches fixture");
}

void testStandardEncryptionWriter(TestContext& test) {
    for (const std::uint32_t bits : {128u, 192u, 256u}) {
        const auto path = std::filesystem::temp_directory_path() / ("xlpp_standard_writer_" + std::to_string(bits) + ".xlsx");
        xlpp::Workbook workbook; auto& sheet = workbook.addWorksheet("Secure");
        sheet.cell("A1").setValue(std::string("standard outbound")); sheet.cell("A2").setValue(static_cast<double>(bits));
        xlpp::SaveOptions save; save.encryptionPassword = "StandardPass!"; save.encryptionMode = xlpp::OfficeEncryptionMode::StandardAesSha1; save.encryptionKeyBits = bits;
        workbook.save(path, save);
        const auto info = xlpp::inspectOfficeEncryption(path);
        test.checkTrue(info.encrypted && info.supported && info.mode == xlpp::OfficeEncryptionMode::StandardAesSha1,
                       "Standard writer produces a recognized supported encrypted package");
        test.checkEqual(info.keyBits, bits, "Standard writer reports requested AES key size");
        test.checkEqual(info.spinCount, std::uint32_t{50000}, "Standard writer uses the required 50,000 SHA-1 iterations");
        xlpp::LoadOptions load; load.password = "StandardPass!"; xlpp::Workbook reopened; reopened.load(path, load);
        test.checkEqual(reopened.worksheet("Secure")->cell("A1").stringValueOr(""), std::string("standard outbound"), "Standard writer round-trips plaintext through independent load path");
        test.checkNear(reopened.worksheet("Secure")->cell("A2").numericValueOr(-1), static_cast<double>(bits), 1e-12, "Standard writer preserves numeric content");
        bool wrong = false; try { xlpp::Workbook invalid; xlpp::LoadOptions bad; bad.password = "wrong"; invalid.load(path, bad); } catch (...) { wrong = true; }
        test.checkTrue(wrong, "Standard writer verifier rejects a wrong password");
        std::filesystem::remove(path);
    }
}

void testLargeAgileDIFATRoundTrip(TestContext& test) {
    // Exercise the CFB DIFAT path: once the encrypted package is large enough, the
    // FAT itself no longer fits entirely in the 109 DIFAT entries stored in the CFB header.
    constexpr std::size_t kPayloadSize = 9u * 1024u * 1024u + 123u;
    std::vector<unsigned char> payload(kPayloadSize);
    for (std::size_t i = 0; i < payload.size(); ++i) {
        payload[i] = static_cast<unsigned char>((i * 29u + (i >> 7u) + 17u) & 0xffu);
    }

    const auto compound = xlpp::internal::encryptAgileOfficePackage(payload, "large-difat", 4);
    test.checkTrue(compound.size() > payload.size(), "Large Agile package emits a CFB envelope");
    const auto readU32 = [&](std::size_t offset) -> std::uint32_t {
        return static_cast<std::uint32_t>(compound.at(offset)) |
               (static_cast<std::uint32_t>(compound.at(offset + 1)) << 8u) |
               (static_cast<std::uint32_t>(compound.at(offset + 2)) << 16u) |
               (static_cast<std::uint32_t>(compound.at(offset + 3)) << 24u);
    };
    test.checkTrue(compound.size() >= 76 && readU32(72) > 0,
                   "Large CFB package uses one or more DIFAT sectors");
    const auto decrypted = xlpp::internal::decryptAgileOfficePackage(compound, "large-difat", true);
    test.checkEqual(decrypted.size(), payload.size(), "Large Agile DIFAT round-trip preserves package size");
    test.checkTrue(decrypted == payload, "Large Agile DIFAT round-trip preserves every package byte");
}

void testExternalFormulaReferenceResolver(TestContext& test) {
    xlpp::Workbook workbook;
    auto& sheet = workbook.addWorksheet("External");
    sheet.cell("A1").setFormula("SUM('[Budget.xlsx]Rates'!A1:A2)+[Budget.xlsx]Rates!B1");
    sheet.cell("A2").setFormula("'[Budget.xlsx]Rates'!Z9");

    xlpp::CalculationOptions options;
    options.externalReferenceResolver = [](const std::string& workbookToken, const std::string& sheetName, const std::string& address) -> std::optional<xlpp::CellValue> {
        if (workbookToken != "Budget.xlsx" || sheetName != "Rates") return std::nullopt;
        if (address == "A1") return xlpp::CellValue{10.0};
        if (address == "A2") return xlpp::CellValue{20.0};
        if (address == "B1") return xlpp::CellValue{5.0};
        return std::nullopt;
    };
    const auto report = workbook.calculateFormulas(options);
    test.checkNear(sheet.cell("A1").numericValueOr(-1), 35.0, 1e-12, "External workbook resolver feeds scalar and range references into calculation");
    test.checkEqual(report.externalReferencesResolved, std::size_t{3}, "External reference report counts resolved source cells");
    test.checkEqual(report.unresolvedExternalReferences, std::size_t{1}, "External reference report counts missing source cells");
    test.checkTrue(sheet.cell("A2").isError() && *sheet.cell("A2").error() == xlpp::CellError::Reference, "Missing external resolver value caches #REF!");
    test.checkTrue(!report.success(), "Unresolved external references participate in calculation success status");

    xlpp::Workbook noResolver;
    auto& nr = noResolver.addWorksheet("NoResolver");
    nr.cell("A1").setFormula("'[Book.xlsx]Sheet1'!A1");
    const auto missing = noResolver.calculateFormulas();
    test.checkEqual(missing.unresolvedExternalReferences, std::size_t{1}, "External workbook reference without a resolver is explicit rather than silently treated as a local name");
}

void testIterativeFormulaCalculation(TestContext& test) {
    xlpp::Workbook workbook;
    auto& sheet = workbook.addWorksheet("Iterative");
    sheet.cell("A1").setFormula("(B1+1)/2");
    sheet.cell("A1").setValue(0.0);
    sheet.cell("B1").setFormula("(A1+1)/2");
    sheet.cell("B1").setValue(0.0);

    xlpp::CalculationOptions options;
    options.iterativeCalculation = true;
    options.maxIterations = 100;
    options.maxChange = 1e-10;
    const auto report = workbook.calculateFormulas(options);
    test.checkEqual(report.circularReferences, std::size_t{0}, "Iterative calculation consumes circular back-edges instead of returning #REF!");
    test.checkEqual(report.iterativeConvergenceFailures, std::size_t{0}, "Convergent circular model reports success");
    test.checkTrue(report.iterativeIterations > 1 && report.iterativeIterations < options.maxIterations, "Circular model converges before the iteration cap");
    test.checkNear(sheet.cell("A1").numericValueOr(-1), 1.0, 1e-8, "Iterative calculation converges first circular formula");
    test.checkNear(sheet.cell("B1").numericValueOr(-1), 1.0, 1e-8, "Iterative calculation converges second circular formula");

    xlpp::Workbook divergent;
    auto& d = divergent.addWorksheet("Divergent");
    d.cell("A1").setFormula("A1+1");
    d.cell("A1").setValue(0.0);
    options.maxIterations = 4;
    options.maxChange = 1e-12;
    const auto failed = divergent.calculateFormulas(options);
    test.checkEqual(failed.iterativeIterations, std::size_t{4}, "Non-convergent circular model respects the iteration cap");
    test.checkEqual(failed.iterativeConvergenceFailures, std::size_t{1}, "Non-convergent circular model is reported explicitly");
    test.checkTrue(!failed.success(), "Iteration convergence failure participates in calculation success status");
}

void testAgileEncryptionRoundTrip(TestContext& test) {
    const auto path = std::filesystem::temp_directory_path() / "xlpp_agile_encryption_test.xlsx";
    const auto streamPath = std::filesystem::temp_directory_path() / "xlpp_agile_encryption_stream_test.xlsx";
    const auto rotatedPath = std::filesystem::temp_directory_path() / "xlpp_agile_encryption_rotated_test.xlsx";
    const auto decryptedPath = std::filesystem::temp_directory_path() / "xlpp_agile_encryption_removed_test.xlsx";
    const auto calcEncryptedPath = std::filesystem::temp_directory_path() / "xlpp_agile_calculate_before_save_test.xlsx";
    std::filesystem::remove(path);
    std::filesystem::remove(streamPath);
    std::filesystem::remove(rotatedPath);
    std::filesystem::remove(decryptedPath);
    std::filesystem::remove(calcEncryptedPath);

    xlpp::Workbook workbook;
    auto& sheet = workbook.addWorksheet("Secure Data");
    sheet.cell("A1").setValue("encrypted payload");
    sheet.cell("A2").setValue(41.0);
    sheet.cell("A3").setFormula("A2+1");
    workbook.calculateFormulas();

    xlpp::SaveOptions saveOptions;
    saveOptions.encryptionPassword = "S3cure-Päss-测试";
    saveOptions.encryptionSpinCount = 2500; // Keep the unit suite fast; production default is 100,000.
    workbook.save(path, saveOptions);

    test.checkTrue(xlpp::looksLikeEncryptedOfficeFile(path), "Encrypted save emits an OLE/CFB container");
    const auto info = xlpp::inspectOfficeEncryption(path);
    test.checkTrue(info.encrypted, "Encryption inspector identifies encrypted OOXML");
    test.checkTrue(info.supported, "Encryption inspector recognizes the generated Agile profile");
    test.checkEqual(info.keyBits, std::uint32_t{256}, "Agile encryption uses a 256-bit package key");
    test.checkEqual(info.hashAlgorithm, std::string("SHA512"), "Agile encryption uses SHA-512");
    test.checkEqual(info.spinCount, std::uint32_t{2500}, "Configured password spin count is serialized");

    bool missingPasswordRejected = false;
    try {
        xlpp::Workbook missing;
        missing.load(path);
    } catch (const std::exception&) {
        missingPasswordRejected = true;
    }
    test.checkTrue(missingPasswordRejected, "Encrypted workbook requires a password");

    bool wrongPasswordRejected = false;
    try {
        xlpp::Workbook wrong;
        xlpp::LoadOptions options;
        options.password = "wrong password";
        wrong.load(path, options);
    } catch (const std::exception&) {
        wrongPasswordRejected = true;
    }
    test.checkTrue(wrongPasswordRejected, "Wrong Agile password is rejected by verifier hash");

    xlpp::Workbook loaded;
    xlpp::LoadOptions loadOptions;
    loadOptions.password = "S3cure-Päss-测试";
    loaded.load(path, loadOptions);
    auto* loadedSheet = loaded.worksheet("Secure Data");
    test.checkTrue(loadedSheet != nullptr, "Correct password decrypts workbook package");
    test.checkEqual(std::get<std::string>(loadedSheet->cell("A1").value()), std::string("encrypted payload"), "Encrypted string round-trips");
    test.checkNear(std::get<double>(loadedSheet->cell("A3").value()), 42.0, 1e-12, "Encrypted formula cached value round-trips");

    std::stringstream encryptedStream(std::ios::in | std::ios::out | std::ios::binary);
    workbook.save(encryptedStream, saveOptions);
    const auto streamBytes = encryptedStream.str();
    test.checkTrue(streamBytes.size() > 8 && static_cast<unsigned char>(streamBytes[0]) == 0xD0,
                   "Encrypted ostream save returns CFB bytes");
    encryptedStream.seekg(0);
    xlpp::Workbook streamLoaded;
    streamLoaded.load(encryptedStream, loadOptions);
    test.checkEqual(std::get<std::string>(streamLoaded.worksheet("Secure Data")->cell("A1").value()),
                    std::string("encrypted payload"), "Encrypted stream load decrypts through normal workbook pipeline");

    // An ordinary unencrypted workbook remains a ZIP and ignores LoadOptions::password.
    xlpp::SaveOptions plainOptions;
    workbook.save(streamPath, plainOptions);
    test.checkTrue(!xlpp::looksLikeEncryptedOfficeFile(streamPath), "Plain save remains an OOXML ZIP package");
    xlpp::Workbook plainLoaded;
    plainLoaded.load(streamPath, loadOptions);
    test.checkTrue(plainLoaded.worksheet("Secure Data") != nullptr, "Password option does not interfere with plain workbooks");

    // Password lifecycle is a load + save operation: the decrypted in-memory workbook
    // can be re-encrypted with a different password or saved as an ordinary ZIP.
    xlpp::SaveOptions rotatedOptions;
    rotatedOptions.encryptionPassword = "Rotated-Password-2026!";
    rotatedOptions.encryptionSpinCount = 64;
    loaded.save(rotatedPath, rotatedOptions);
    bool oldPasswordRejectedAfterRotation = false;
    try { xlpp::Workbook oldPassword; oldPassword.load(rotatedPath, loadOptions); }
    catch (const std::exception&) { oldPasswordRejectedAfterRotation = true; }
    test.checkTrue(oldPasswordRejectedAfterRotation, "Password rotation invalidates the previous password");
    xlpp::LoadOptions rotatedLoad; rotatedLoad.password = "Rotated-Password-2026!";
    xlpp::Workbook rotated; rotated.load(rotatedPath, rotatedLoad);
    test.checkEqual(rotated.worksheet("Secure Data")->cell("A1").stringValueOr(""), std::string("encrypted payload"),
                    "Workbook reloads after password rotation");
    rotated.save(decryptedPath);
    test.checkTrue(!xlpp::looksLikeEncryptedOfficeFile(decryptedPath), "Saving without a password removes file-open encryption");
    xlpp::Workbook decrypted; decrypted.load(decryptedPath);
    test.checkEqual(decrypted.worksheet("Secure Data")->cell("A1").stringValueOr(""), std::string("encrypted payload"),
                    "Workbook content survives encryption removal");

    xlpp::Workbook pendingCalculation;
    auto& pendingSheet = pendingCalculation.addWorksheet("CalcEncrypted");
    pendingSheet.cell("A1").setValue(9.0);
    pendingSheet.cell("A2").setFormula("A1*7");
    xlpp::SaveOptions calcEncryptedOptions;
    calcEncryptedOptions.calculateFormulasBeforeSave = true;
    calcEncryptedOptions.encryptionPassword = "Calc+Encryption";
    calcEncryptedOptions.encryptionSpinCount = 32;
    pendingCalculation.save(calcEncryptedPath, calcEncryptedOptions);
    test.checkTrue(std::holds_alternative<std::monostate>(pendingSheet.cell("A2").value()),
                   "Encrypted calculate-before-save remains const for the caller workbook");
    xlpp::LoadOptions calcEncryptedLoad; calcEncryptedLoad.password = "Calc+Encryption";
    xlpp::Workbook calcEncrypted; calcEncrypted.load(calcEncryptedPath, calcEncryptedLoad);
    test.checkNear(calcEncrypted.worksheet("CalcEncrypted")->cell("A2").numericValueOr(-1), 63.0, 1e-12,
                   "Formula calculation runs before encrypted package serialization");

    // Integrity is authenticated over the complete EncryptedPackage stream. Flip a ciphertext byte
    // while leaving the CFB structure and password verifier intact.
    std::vector<unsigned char> payload(1237);
    for(std::size_t i=0;i<payload.size();++i)payload[i]=static_cast<unsigned char>((i*37u+11u)&0xffu);
    auto compound=xlpp::internal::encryptAgileOfficePackage(payload,"integrity-password",32);
    std::array<unsigned char,8> lengthBytes{};
    const std::uint64_t payloadLength=payload.size();
    for(unsigned shift=0;shift<64;shift+=8)lengthBytes[shift/8]=static_cast<unsigned char>(payloadLength>>shift);
    auto marker=std::search(compound.begin(),compound.end(),lengthBytes.begin(),lengthBytes.end());
    test.checkTrue(marker!=compound.end(),"EncryptedPackage stream-size marker is locatable in integrity fixture");
    if(marker!=compound.end()&&std::distance(marker,compound.end())>32){
        *(marker+19)^=0x40;
        bool integrityRejected=false;
        try{(void)xlpp::internal::decryptAgileOfficePackage(compound,"integrity-password",true);}catch(const std::exception&){integrityRejected=true;}
        test.checkTrue(integrityRejected,"Agile HMAC rejects ciphertext tampering");
        const auto unchecked=xlpp::internal::decryptAgileOfficePackage(compound,"integrity-password",false);
        test.checkTrue(unchecked!=payload,"Disabling integrity verification exposes the modified plaintext rather than hiding corruption");
    }

    std::filesystem::remove(path);
    std::filesystem::remove(streamPath);
    std::filesystem::remove(rotatedPath);
    std::filesystem::remove(decryptedPath);
    std::filesystem::remove(calcEncryptedPath);
}

void testReferenceTranslatorStructuralMatrix(TestContext& test) {
    xlpp::StructuralEdit insertRows{"Data", xlpp::StructuralEditKind::InsertRows, 2, 1};
    const auto formula = xlpp::translateFormulaReferences(
        "SUM(A1:A4)+Data!$B$2+'Other Sheet'!C3+LOG10(100)+\"A1:A4\"", "Data", insertRows);
    test.checkEqual(formula.value,
                    std::string("SUM(A1:A5)+Data!$B$3+'Other Sheet'!C3+LOG10(100)+\"A1:A4\""),
                    "Reference translator expands local ranges, shifts qualified cells, and preserves functions/strings");
    test.checkEqual(formula.referencesChanged, std::size_t{2}, "Reference translator counts changed formula references");

    xlpp::StructuralEdit deleteRows{"Data", xlpp::StructuralEditKind::DeleteRows, 2, 2};
    auto shrink = xlpp::translateFormulaReferences("SUM(A1:A5)+A2+Data!A4", "Data", deleteRows);
    test.checkEqual(shrink.value, std::string("SUM(A1:A3)+#REF!+Data!A2"),
                    "Row deletion shrinks ranges and emits #REF for deleted singular references");
    test.checkEqual(shrink.referencesInvalidated, std::size_t{1}, "Deleted singular reference is counted as invalid");

    xlpp::StructuralEdit insertColumns{"Data", xlpp::StructuralEditKind::InsertColumns, 2, 2};
    auto whole = xlpp::translateRangeReferences("$A:$D $1:$4 A1:D10", "Data", insertColumns);
    test.checkEqual(whole.value, std::string("$A:$F $1:$4 A1:F10"),
                    "Range translator expands whole-column and cell ranges");

    xlpp::StructuralEdit deleteColumns{"Data", xlpp::StructuralEditKind::DeleteColumns, 1, 4};
    auto invalid = xlpp::translateRangeReferences("A1:D10", "Data", deleteColumns);
    test.checkEqual(invalid.value, std::string("#REF!"), "Deleting an entire referenced column interval invalidates the range");
}

void testWorkbookStructuralTransaction(TestContext& test) {
    xlpp::Workbook wb;
    auto& data = wb.addWorksheet("Data");
    auto& reportSheet = wb.addWorksheet("Report");

    data.cell("A1").setValue("Category");
    data.cell("B1").setValue("Value");
    for (std::size_t row = 2; row <= 4; ++row) {
        data.cell(row, 1).setValue(std::string("R") + std::to_string(row));
        data.cell(row, 2).setValue(static_cast<double>(row * 10));
    }
    data.cell("C5").setFormula("SUM(B2:B4)");
    data.cell("D5").setArrayFormula("SUM(B2:B4)", "D5:D6");
    data.mergeCells("E1:F2");
    data.freezePanes("A2");
    data.rowDimension(3).height = 25.0;
    data.columnDimension(2).width = 18.0;
    data.autoFilter().setReference("A1:B4");
    data.autoFilter().sortState().setReference("A1:B4");
    data.autoFilter().sortState().addCondition("B2:B4", true);
    data.conditionalFormatting().addRule("B2:B4", xlpp::ConditionalRule::formula("B2>Data!$B$2"));
    auto validation = xlpp::DataValidation::list("A2:A4", "Data!$A$2:$A$4");
    data.dataValidations().add(std::move(validation));
    auto& table = data.addTable("SalesTable", "A1:B4");
    table.addColumn("Category");
    table.addColumn("Value");
    data.setPrintArea("A1:F5");
    data.setPrintTitlesRows("$1:$1");
    data.setPrintTitlesCols("$A:$B");
    xlpp::Image generatedImage("H3", {1,2,3,4}, "png");
    data.addImage(std::move(generatedImage));

    reportSheet.cell("A1").setFormula("Data!B4");
    xlpp::Hyperlink internal("#Data!B4");
    internal.setExternal(false);
    reportSheet.cell("A2").setHyperlink(std::move(internal));

    xlpp::Chart chart(xlpp::Chart::Type::Line);
    xlpp::ChartSeries series("Values");
    series.setCategoriesReference("=Data!$A$2:$A$4");
    series.setValuesReference("=Data!$B$2:$B$4");
    chart.addSeries(std::move(series));
    reportSheet.addChart(std::move(chart));

    xlpp::PivotTable pivot("SummaryPivot");
    pivot.setLocation("D2:G10");
    pivot.cache().setSourceData("Data!A1:B4");
    reportSheet.addPivotTable(std::move(pivot));

    xlpp::DefinedName sales("SalesRange", "Data!$A$1:$B$4");
    wb.addDefinedName(std::move(sales));

    xlpp::StructuralEditOptions options;
    options.synchronizeChartCaches = false;
    auto edit = wb.insertRows("Data", 3, 1, options);
    test.checkEqual(edit.worksheetsVisited, std::size_t{2}, "Workbook structural transaction visits all worksheets");
    test.checkTrue(edit.cellsMoved >= 4, "Workbook structural transaction reports moved cells");
    test.checkTrue(edit.formulasUpdated >= 2, "Local and cross-sheet formulas are rewritten");
    test.checkTrue(edit.definedNamesUpdated >= 1, "Defined names are rewritten");
    test.checkTrue(edit.chartReferencesUpdated >= 1, "Chart source references are rewritten");
    test.checkTrue(edit.pivotReferencesUpdated >= 1, "Pivot source references are rewritten");

    auto* updatedData = wb.worksheet("Data");
    auto* updatedReport = wb.worksheet("Report");
    test.checkTrue(updatedData != nullptr && updatedReport != nullptr, "Structural transaction preserves worksheets");
    test.checkEqual(updatedData->cell("C6").formula(), std::string("SUM(B2:B5)"),
                    "Formula cell moves and its source range expands");
    test.checkEqual(updatedData->cell("D6").formulaMetadata().reference(), std::string("D6:D7"),
                    "Array formula spill/reference metadata moves with rows");
    test.checkEqual(updatedReport->cell("A1").formula(), std::string("Data!B5"),
                    "Cross-sheet formula follows inserted row");
    test.checkEqual(updatedReport->cell("A2").hyperlinkValue()->target(), std::string("#Data!B5"),
                    "Internal hyperlink follows inserted row");
    test.checkEqual(wb.definedName("SalesRange")->value(), std::string("Data!$A$1:$B$5"),
                    "Workbook defined name expands across inserted row");
    test.checkEqual(updatedData->mergedRanges().front(), std::string("E1:F2"),
                    "Merge outside insertion boundary remains stable");
    test.checkEqual(updatedData->autoFilter().reference(), std::string("A1:B5"), "AutoFilter expands with inserted row");
    test.checkEqual(updatedData->tables().front().reference(), std::string("A1:B5"), "Table expands with inserted row");
    test.checkEqual(updatedData->dataValidations().items().front().reference(), std::string("A2:A5"),
                    "Data validation sqref expands with inserted row");
    test.checkEqual(updatedData->conditionalFormatting().entries().front().reference(), std::string("B2:B5"),
                    "Conditional formatting sqref expands with inserted row");
    test.checkEqual(updatedData->printArea(), std::string("A1:F6"), "Print area expands with inserted row");
    test.checkEqual(updatedData->images().front().anchor(), std::string("H4"), "Generated image anchor moves with inserted row");
    test.checkEqual(updatedReport->charts().front().series().front().valuesReference(), std::string("=Data!$B$2:$B$5"),
                    "Chart values reference expands with inserted row");
    test.checkEqual(updatedReport->pivotTables().front().cache().sourceData(), std::string("Data!A1:B5"),
                    "Pivot cache source expands with inserted row");
    test.checkTrue(updatedReport->pivotTables().front().cache().records().empty(),
                   "Pivot records invalidate after source structural edit");
    test.checkTrue(wb.calcProperties().fullCalcOnLoad(), "Structural transaction requests full host recalculation");

    // Column insertion inside a table expands both ref width and semantic
    // TableColumn count, preventing a package that Excel would repair.
    auto colEdit = wb.insertColumns("Data", 2, 1, options);
    (void)colEdit;
    updatedData = wb.worksheet("Data");
    test.checkEqual(updatedData->tables().front().reference(), std::string("A1:C5"),
                    "Table reference expands for an inserted interior column");
    test.checkEqual(updatedData->tables().front().columns().size(), std::size_t{3},
                    "Table column model expands with table reference width");
    test.checkEqual(updatedData->tables().front().columns()[2].id(), std::size_t{3},
                    "Table columns are renumbered after structural edit");

    // Strict transactional mode must not partially mutate the caller when a
    // deletion would invalidate a dependency.
    const auto before = wb.worksheet("Report")->cell("A1").formula();
    xlpp::StructuralEditOptions strict = options;
    strict.failOnInvalidReference = true;
    bool rejected = false;
    try {
        wb.deleteRows("Data", 5, 1, strict); // Report!A1 currently references Data!C5/B5 depending column shift; delete target row.
    } catch (const std::exception&) { rejected = true; }
    test.checkTrue(rejected, "Strict structural transaction rejects edits that create #REF dependencies");
    test.checkEqual(wb.worksheet("Report")->cell("A1").formula(), before,
                    "Rejected structural transaction rolls back workbook state");
}


} // namespace

void registerFormulaDependencyTests(std::vector<TestCase>& tests) {
    tests.push_back({"Dependency-aware worksheet rename", testWorksheetRenameTransaction});
    tests.push_back({"Formula calculation engine", testFormulaCalculationEngine});
    tests.push_back({"Advanced formula calculation matrix", testAdvancedFormulaCalculation});
    tests.push_back({"Formula dependency graph", testFormulaDependencyGraph});
    tests.push_back({"Dynamic arrays and structured references", testDynamicArraysAndStructuredReferences});
    tests.push_back({"Standard AES LibreOffice encrypted fixture", testStandardEncryptionLibreOfficeFixture});
    tests.push_back({"Standard encryption writer", testStandardEncryptionWriter});
    tests.push_back({"Large Agile CFB DIFAT round-trip", testLargeAgileDIFATRoundTrip});
    tests.push_back({"External formula reference resolver", testExternalFormulaReferenceResolver});
    tests.push_back({"Iterative formula calculation", testIterativeFormulaCalculation});
    tests.push_back({"Agile AES-256 workbook encryption", testAgileEncryptionRoundTrip});
    tests.push_back({"Reference translator structural matrix", testReferenceTranslatorStructuralMatrix});
    tests.push_back({"Workbook structural transaction", testWorkbookStructuralTransaction});
}
