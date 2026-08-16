#include "WorkbookChartEditing.h"
#include "WorkbookChartSerializer.h"
#include "WorkbookChartReader.h"
#include "WorkbookDrawingEdit.h"
#include "WorkbookDrawingAccess.h"
#include "WorkbookPartXml.h"
#include "WorkbookNamespaces.h"
#include "../Packaging/RelationshipGraph.h"
#include <XLPP/Chart/Chart.h>
#include <XLPP/Worksheet/Drawings/Image.h>
#include "../XML/XmlUtilities.h"
#include "../XML/NumericParsing.h"
#include "../Packaging/ZipArchive.h"
#include <algorithm>
#include <array>
#include <cctype>
#include <cmath>
#include <cstdlib>
#include <deque>
#include <functional>
#include <iomanip>
#include <limits>
#include <memory>
#include <optional>
#include <set>
#include <sstream>
#include <stdexcept>
#include <string>
#include <tuple>
#include <unordered_map>
#include <unordered_set>
#include <utility>
#include <vector>
#include <cstdint>

namespace xlpp {
namespace internal {
std::string generatedChartSeriesXml(const xlpp::ChartSeries& series, xlpp::Chart::Type type, std::size_t index) {
    std::ostringstream seriesXml;
    seriesXml << "<c:ser><c:idx val=\"" << index << "\"/><c:order val=\"" << index << "\"/>";
    if (!series.titleReference().empty()) {
        seriesXml << "<c:tx><c:strRef><c:f>" << xmlEscape(series.titleReference()) << "</c:f>"
                  << chartSeriesCacheXml(series.titleCache()) << "</c:strRef></c:tx>";
    } else if (!series.title().empty()) {
        seriesXml << "<c:tx><c:v>" << xmlEscape(series.title()) << "</c:v></c:tx>";
    }

    if (type == xlpp::Chart::Type::Radar && series.markerFormat().present) {
        std::string markerOwner = "<c:owner></c:owner>";
        if (!patchMarkerFormatInOwner(markerOwner, series.markerFormat()))
            throw std::runtime_error("Failed to serialize radar marker format");
        const auto markerNodes = drawingTags(markerOwner, "c:marker", "marker");
        if (!markerNodes.empty()) seriesXml << markerNodes.front();
    }

    if (generatedChartTypeUsesXYAxes(type)) {
        if (series.categoriesReference().empty() || series.valuesReference().empty())
            throw std::invalid_argument("Scatter/Bubble chart series require both X and Y references");
        seriesXml << "<c:xVal><c:numRef><c:f>" << xmlEscape(series.categoriesReference()) << "</c:f>"
                  << chartSeriesCacheXml(series.categoriesCache()) << "</c:numRef></c:xVal>";
        seriesXml << "<c:yVal><c:numRef><c:f>" << xmlEscape(series.valuesReference()) << "</c:f>"
                  << chartSeriesCacheXml(series.valuesCache()) << "</c:numRef></c:yVal>";
        if (type == xlpp::Chart::Type::Bubble) {
            if (series.bubbleSizeReference().empty())
                throw std::invalid_argument("Bubble chart series require a bubble-size reference");
            seriesXml << "<c:bubbleSize><c:numRef><c:f>" << xmlEscape(series.bubbleSizeReference()) << "</c:f>"
                      << chartSeriesCacheXml(series.bubbleSizeCache()) << "</c:numRef></c:bubbleSize>";
        }
    } else {
        if (!series.categoriesReference().empty()) {
            const bool numericCategories = type == xlpp::Chart::Type::Stock || (series.categoriesCache().present && series.categoriesCache().numeric);
            const auto refTag = numericCategories ? "numRef" : "strRef";
            seriesXml << "<c:cat><c:" << refTag << "><c:f>" << xmlEscape(series.categoriesReference()) << "</c:f>"
                      << chartSeriesCacheXml(series.categoriesCache()) << "</c:" << refTag << "></c:cat>";
        }
        if (!series.valuesReference().empty())
            seriesXml << "<c:val><c:numRef><c:f>" << xmlEscape(series.valuesReference()) << "</c:f>"
                      << chartSeriesCacheXml(series.valuesCache()) << "</c:numRef></c:val>";
    }
    seriesXml << "</c:ser>";
    return seriesXml.str();
}

std::string combinedChartXml(const xlpp::Chart& chart, bool strict) {
    if (chart.plots().size() < 2) throw std::invalid_argument("Combined chart requires at least two plots");
    bool xyAxes = generatedChartTypeUsesXYAxes(chart.plots().front().type);
    bool hasSecondary = false;
    std::size_t expectedSeries = 0;
    for (const auto& plot : chart.plots()) {
        if (!generatedChartTypeHasAxes(plot.type))
            throw std::invalid_argument("Combined generated charts currently require axis-based chart types");
        if (generatedChartTypeUsesXYAxes(plot.type) != xyAxes)
            throw std::invalid_argument("Cannot combine categorical-axis and XY-axis plots in one generated chart");
        if (plot.firstSeries != expectedSeries || plot.firstSeries + plot.seriesCount > chart.series().size())
            throw std::invalid_argument("Combined chart plot series ranges must be contiguous and in bounds");
        expectedSeries += plot.seriesCount;
        hasSecondary = hasSecondary || plot.usesSecondaryAxes;
    }
    if (expectedSeries != chart.series().size())
        throw std::invalid_argument("Combined chart plots must own every generated series exactly once");

    const std::uint64_t xAxisId = 10;
    const std::uint64_t primaryYAxisId = 100;
    const std::uint64_t secondaryYAxisId = 200;
    std::ostringstream xml;
    xml << "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>";
    xml << "<c:chartSpace xmlns:c=\"" << (strict ? "http://purl.oclc.org/ooxml/drawingml/chart" : "http://schemas.openxmlformats.org/drawingml/2006/chart") << "\"";
    xml << " xmlns:a=\"" << (strict ? "http://purl.oclc.org/ooxml/drawingml/main" : "http://schemas.openxmlformats.org/drawingml/2006/main") << "\">";
    if (!chart.style().empty()) xml << "<c:style val=\"" << xmlEscape(chart.style()) << "\"/>";
    if (chart.pivotSource().present) {
        if (chart.pivotSource().pivotTableName.empty()) throw std::invalid_argument("Pivot chart source name cannot be empty");
        xml << "<c:pivotSource><c:name>" << xmlEscape(chart.pivotSource().pivotTableName)
            << "</c:name><c:fmtId val=\"" << std::max(0, chart.pivotSource().formatId) << "\"/></c:pivotSource>";
    }
    xml << "<c:chart>";
    if (!chart.title().empty())
        xml << "<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang=\"en-US\"/><a:t>" << xmlEscape(chart.title()) << "</a:t></a:r></a:p></c:rich></c:tx><c:overlay val=\"0\"/></c:title>";
    xml << "<c:plotArea><c:layout/>";

    for (const auto& plot : chart.plots()) {
        const auto type = plot.type;
        xml << "<c:" << xlpp::Chart::typeName(type, plot.grouping) << ">";
        if (type == xlpp::Chart::Type::Bar) {
            xml << "<c:barDir val=\"col\"/>";
            const char* grouping = plot.grouping == xlpp::Chart::Grouping::Stacked ? "stacked"
                : (plot.grouping == xlpp::Chart::Grouping::PercentStacked ? "percentStacked" : "clustered");
            xml << "<c:grouping val=\"" << grouping << "\"/><c:varyColors val=\"0\"/>";
        } else if (type == xlpp::Chart::Type::Line || type == xlpp::Chart::Type::Area) {
            const char* grouping = plot.grouping == xlpp::Chart::Grouping::Stacked ? "stacked"
                : (plot.grouping == xlpp::Chart::Grouping::PercentStacked ? "percentStacked" : "standard");
            xml << "<c:grouping val=\"" << grouping << "\"/>";
        } else if (type == xlpp::Chart::Type::Scatter) {
            xml << "<c:scatterStyle val=\"" << xmlEscape(chart.scatterStyle()) << "\"/><c:varyColors val=\"0\"/>";
        } else if (type == xlpp::Chart::Type::Bubble) {
            xml << "<c:varyColors val=\"0\"/>";
        } else {
            throw std::invalid_argument("Unsupported chart type in generated combined chart");
        }
        for (std::size_t i = 0; i < plot.seriesCount; ++i) {
            const std::size_t seriesIndex = plot.firstSeries + i;
            xml << generatedChartSeriesXml(chart.series()[seriesIndex], type, seriesIndex);
        }
        xml << generatedPlotAuxiliaryXml(plot, strict);
        xml << "<c:axId val=\"" << xAxisId << "\"/><c:axId val=\""
            << (plot.usesSecondaryAxes ? secondaryYAxisId : primaryYAxisId) << "\"/>";
        xml << "</c:" << xlpp::Chart::typeName(type, plot.grouping) << ">";
    }

    if (xyAxes) {
        xml << "<c:valAx><c:axId val=\"" << xAxisId << "\"/><c:scaling><c:orientation val=\"minMax\"/></c:scaling><c:delete val=\"0\"/><c:axPos val=\"b\"/>";
        if (!chart.xAxisTitle().empty()) xml << "<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>" << xmlEscape(chart.xAxisTitle()) << "</a:t></a:r></a:p></c:rich></c:tx><c:overlay val=\"0\"/></c:title>";
        xml << "<c:numFmt formatCode=\"General\" sourceLinked=\"1\"/><c:majorTickMark val=\"none\"/><c:minorTickMark val=\"none\"/><c:tickLblPos val=\"nextTo\"/><c:crossAx val=\"" << primaryYAxisId << "\"/><c:crosses val=\"autoZero\"/></c:valAx>";
    } else {
        xml << "<c:catAx><c:axId val=\"" << xAxisId << "\"/><c:scaling><c:orientation val=\"minMax\"/></c:scaling><c:delete val=\"0\"/><c:axPos val=\"b\"/>";
        if (!chart.xAxisTitle().empty()) xml << "<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>" << xmlEscape(chart.xAxisTitle()) << "</a:t></a:r></a:p></c:rich></c:tx><c:overlay val=\"0\"/></c:title>";
        xml << "<c:numFmt formatCode=\"General\" sourceLinked=\"1\"/><c:majorTickMark val=\"none\"/><c:minorTickMark val=\"none\"/><c:tickLblPos val=\"nextTo\"/><c:crossAx val=\"" << primaryYAxisId << "\"/><c:crosses val=\"autoZero\"/><c:auto val=\"1\"/><c:lblAlgn val=\"ctr\"/><c:lblOffset val=\"100\"/></c:catAx>";
    }
    xml << "<c:valAx><c:axId val=\"" << primaryYAxisId << "\"/><c:scaling><c:orientation val=\"minMax\"/></c:scaling><c:delete val=\"0\"/><c:axPos val=\"l\"/>";
    if (!chart.yAxisTitle().empty()) xml << "<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>" << xmlEscape(chart.yAxisTitle()) << "</a:t></a:r></a:p></c:rich></c:tx><c:overlay val=\"0\"/></c:title>";
    xml << "<c:numFmt formatCode=\"General\" sourceLinked=\"1\"/><c:majorGridlines/><c:majorTickMark val=\"none\"/><c:minorTickMark val=\"none\"/><c:tickLblPos val=\"nextTo\"/><c:crossAx val=\"" << xAxisId << "\"/><c:crosses val=\"autoZero\"/><c:crossBetween val=\"between\"/></c:valAx>";
    if (hasSecondary) {
        xml << "<c:valAx><c:axId val=\"" << secondaryYAxisId << "\"/><c:scaling><c:orientation val=\"minMax\"/></c:scaling><c:delete val=\"0\"/><c:axPos val=\"r\"/><c:numFmt formatCode=\"General\" sourceLinked=\"1\"/><c:majorTickMark val=\"none\"/><c:minorTickMark val=\"none\"/><c:tickLblPos val=\"nextTo\"/><c:crossAx val=\"" << xAxisId << "\"/><c:crosses val=\"max\"/><c:crossBetween val=\"between\"/></c:valAx>";
    }
    if (chart.dataTable().present) xml << generatedDataTableXml(chart.dataTable(), strict);
    xml << "</c:plotArea>";
    if (chart.showLegend()) xml << "<c:legend><c:legendPos val=\"" << xmlEscape(chart.legendPosition()) << "\"/></c:legend>";
    xml << "</c:chart></c:chartSpace>";
    return xml.str();
}

std::string chartXml(const xlpp::Chart& chart, bool strict) {
    if (chart.combined()) return combinedChartXml(chart, strict);
    std::ostringstream xml;
    const auto type = chart.type();
    const bool threeAxis = type == xlpp::Chart::Type::Bar3D || type == xlpp::Chart::Type::Line3D ||
                           type == xlpp::Chart::Type::Area3D || type == xlpp::Chart::Type::Surface ||
                           type == xlpp::Chart::Type::Surface3D;
    const bool projectedPie = type == xlpp::Chart::Type::PieOfPie || type == xlpp::Chart::Type::BarOfPie;
    const bool hasAxes = type != xlpp::Chart::Type::Pie && type != xlpp::Chart::Type::Pie3D && type != xlpp::Chart::Type::Doughnut && !projectedPie;
    const bool barLike = type == xlpp::Chart::Type::Bar || type == xlpp::Chart::Type::Bar3D;
    const bool lineOrArea3D = type == xlpp::Chart::Type::Line3D || type == xlpp::Chart::Type::Area3D;
    const bool surface = type == xlpp::Chart::Type::Surface || type == xlpp::Chart::Type::Surface3D;

    xml << "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>";
    xml << "<c:chartSpace xmlns:c=\"" << (strict ? "http://purl.oclc.org/ooxml/drawingml/chart" : "http://schemas.openxmlformats.org/drawingml/2006/chart") << "\"";
    xml << " xmlns:a=\"" << (strict ? "http://purl.oclc.org/ooxml/drawingml/main" : "http://schemas.openxmlformats.org/drawingml/2006/main") << "\">";
    if (!chart.style().empty()) xml << "<c:style val=\"" << xmlEscape(chart.style()) << "\"/>";
    if (chart.pivotSource().present) {
        if (chart.pivotSource().pivotTableName.empty())
            throw std::invalid_argument("Pivot chart source name cannot be empty");
        xml << "<c:pivotSource><c:name>" << xmlEscape(chart.pivotSource().pivotTableName)
            << "</c:name><c:fmtId val=\"" << std::max(0, chart.pivotSource().formatId)
            << "\"/></c:pivotSource>";
    }
    xml << "<c:chart>";
    if (!chart.title().empty())
        xml << "<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang=\"en-US\"/><a:t>" << xmlEscape(chart.title()) << "</a:t></a:r></a:p></c:rich></c:tx><c:overlay val=\"0\"/></c:title>";
    if (chart.view3D().present) xml << chartView3DXml(chart.view3D(), true);
    if (chart.floorFormat().present) xml << generatedChartWallXml("floor", chart.floorFormat(), true);
    if (chart.sideWallFormat().present) xml << generatedChartWallXml("sideWall", chart.sideWallFormat(), true);
    if (chart.backWallFormat().present) xml << generatedChartWallXml("backWall", chart.backWallFormat(), true);
    if (type == xlpp::Chart::Type::Stock && chart.series().size() != 3 && chart.series().size() != 4)
        throw std::invalid_argument("Stock charts require exactly 3 (high-low-close) or 4 (open-high-low-close) series");

    xml << "<c:plotArea><c:layout/>";
    xml << "<c:" << xlpp::Chart::typeName(type, chart.grouping()) << ">";
    if (barLike) {
        xml << "<c:barDir val=\"col\"/>";
        const char* grouping = chart.grouping() == xlpp::Chart::Grouping::Stacked ? "stacked"
            : (chart.grouping() == xlpp::Chart::Grouping::PercentStacked ? "percentStacked" : "clustered");
        xml << "<c:grouping val=\"" << grouping << "\"/><c:varyColors val=\"0\"/>";
    } else if (lineOrArea3D) {
        const char* grouping = chart.grouping() == xlpp::Chart::Grouping::Stacked ? "stacked"
            : (chart.grouping() == xlpp::Chart::Grouping::PercentStacked ? "percentStacked" : "standard");
        xml << "<c:grouping val=\"" << grouping << "\"/>";
    } else if (type == xlpp::Chart::Type::Pie3D || type == xlpp::Chart::Type::Pie || type == xlpp::Chart::Type::Doughnut || projectedPie) {
        xml << "<c:varyColors val=\"1\"/>";
    } else if (type == xlpp::Chart::Type::Scatter) {
        xml << "<c:scatterStyle val=\"" << xmlEscape(chart.scatterStyle()) << "\"/><c:varyColors val=\"0\"/>";
    } else if (type == xlpp::Chart::Type::Bubble) {
        xml << "<c:varyColors val=\"0\"/>";
    } else if (type == xlpp::Chart::Type::Radar) {
        const auto* plot = chart.primaryPlotOrNull();
        xml << "<c:radarStyle val=\"" << xmlEscape(plot && !plot->radarStyle.empty() ? plot->radarStyle : "standard") << "\"/>";
    } else if (surface) {
        const auto* plot = chart.primaryPlotOrNull();
        if (plot && plot->hasWireframe) xml << "<c:wireframe val=\"" << (plot->wireframe ? "1" : "0") << "\"/>";
    }

    for (std::size_t s = 0; s < chart.series().size(); ++s)
        xml << generatedChartSeriesXml(chart.series()[s], type, s);
    if (const auto* generatedPlot = chart.primaryPlotOrNull()) {
        xml << generatedPlotAuxiliaryXml(*generatedPlot, strict);
        if (generatedPlot->hasGapDepth && (type == xlpp::Chart::Type::Bar3D || type == xlpp::Chart::Type::Line3D || type == xlpp::Chart::Type::Area3D))
            xml << "<c:gapDepth val=\"" << generatedPlot->gapDepth << "\"/>";
        if (!generatedPlot->shape.empty() && type == xlpp::Chart::Type::Bar3D)
            xml << "<c:shape val=\"" << xmlEscape(generatedPlot->shape) << "\"/>";
        if ((type == xlpp::Chart::Type::Pie || type == xlpp::Chart::Type::Doughnut) && generatedPlot->hasFirstSliceAngle) {
            if (generatedPlot->firstSliceAngle < 0 || generatedPlot->firstSliceAngle > 360) throw std::invalid_argument("Chart first-slice angle must be between 0 and 360 degrees");
            xml << "<c:firstSliceAng val=\"" << generatedPlot->firstSliceAngle << "\"/>";
        }
        if (type == xlpp::Chart::Type::Doughnut && generatedPlot->hasHoleSize) {
            if (generatedPlot->holeSize < 10 || generatedPlot->holeSize > 90) throw std::invalid_argument("Doughnut hole size must be between 10 and 90 percent");
            xml << "<c:holeSize val=\"" << generatedPlot->holeSize << "\"/>";
        }
        if (projectedPie) {
            auto options = generatedPlot->projectedPie;
            options.present = true; options.ofPieType = type == xlpp::Chart::Type::BarOfPie ? "bar" : "pie";
            if (options.gapWidth < 0 || options.gapWidth > 500 || options.secondPlotSize < 5 || options.secondPlotSize > 200)
                throw std::invalid_argument("Projected-pie gap width or second-plot size is outside the supported OOXML range");
            if (options.splitType != "auto" && options.splitType != "cust" && options.splitType != "percent" && options.splitType != "pos" && options.splitType != "val")
                throw std::invalid_argument("Unsupported projected-pie split type");
            xml << "<c:ofPieType val=\"" << options.ofPieType << "\"/>";
            xml << "<c:gapWidth val=\"" << options.gapWidth << "\"/><c:splitType val=\"" << options.splitType << "\"/>";
            if (options.hasSplitPosition) xml << "<c:splitPos val=\"" << options.splitPosition << "\"/>";
            if (!options.customSplitPoints.empty()) { xml << "<c:custSplit>"; for (const auto point : options.customSplitPoints) { if (point < 0) throw std::invalid_argument("Projected-pie custom split indices cannot be negative"); xml << "<c:secondPiePt val=\"" << point << "\"/>"; } xml << "</c:custSplit>"; }
            xml << "<c:secondPieSize val=\"" << options.secondPlotSize << "\"/>";
            if (options.hasSeriesLines) {
                std::string lines = "<c:serLines></c:serLines>";
                if (options.seriesLinesFormat.present && !patchNestedLineFormat(lines, options.seriesLinesFormat)) throw std::runtime_error("Failed to serialize projected-pie series lines");
                xml << lines;
            }
        }
    } else if (projectedPie) {
        xlpp::ChartProjectedPieOptions options; options.present=true; options.ofPieType=type==xlpp::Chart::Type::BarOfPie?"bar":"pie";
        xml << "<c:ofPieType val=\"" << options.ofPieType << "\"/><c:gapWidth val=\"150\"/><c:splitType val=\"auto\"/><c:secondPieSize val=\"75\"/>";
    }
    if (hasAxes) {
        xml << "<c:axId val=\"1\"/><c:axId val=\"2\"/>";
        if (threeAxis) xml << "<c:axId val=\"3\"/>";
    }
    if ((type == xlpp::Chart::Type::Pie || type == xlpp::Chart::Type::Doughnut) && (!chart.primaryPlotOrNull() || !chart.primaryPlotOrNull()->hasFirstSliceAngle))
        xml << "<c:firstSliceAng val=\"0\"/>";
    if (type == xlpp::Chart::Type::Doughnut && (!chart.primaryPlotOrNull() || !chart.primaryPlotOrNull()->hasHoleSize))
        xml << "<c:holeSize val=\"10\"/>";
    xml << "</c:" << xlpp::Chart::typeName(type, chart.grouping()) << ">";

    if (hasAxes) {
        const bool xyAxes = generatedChartTypeUsesXYAxes(type);
        if (xyAxes) {
            xml << "<c:valAx><c:axId val=\"1\"/><c:scaling><c:orientation val=\"minMax\"/></c:scaling><c:delete val=\"0\"/><c:axPos val=\"b\"/>";
            if (!chart.xAxisTitle().empty()) xml << "<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>" << xmlEscape(chart.xAxisTitle()) << "</a:t></a:r></a:p></c:rich></c:tx><c:overlay val=\"0\"/></c:title>";
            xml << "<c:numFmt formatCode=\"General\" sourceLinked=\"1\"/><c:majorTickMark val=\"none\"/><c:minorTickMark val=\"none\"/><c:tickLblPos val=\"nextTo\"/><c:crossAx val=\"2\"/><c:crosses val=\"autoZero\"/></c:valAx>";
        } else {
            xml << "<c:catAx><c:axId val=\"1\"/><c:scaling><c:orientation val=\"minMax\"/></c:scaling><c:delete val=\"0\"/><c:axPos val=\"b\"/>";
            if (!chart.xAxisTitle().empty()) xml << "<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>" << xmlEscape(chart.xAxisTitle()) << "</a:t></a:r></a:p></c:rich></c:tx><c:overlay val=\"0\"/></c:title>";
            xml << "<c:numFmt formatCode=\"General\" sourceLinked=\"1\"/><c:majorTickMark val=\"none\"/><c:minorTickMark val=\"none\"/><c:tickLblPos val=\"nextTo\"/><c:crossAx val=\"2\"/><c:crosses val=\"autoZero\"/><c:auto val=\"1\"/><c:lblAlgn val=\"ctr\"/><c:lblOffset val=\"100\"/></c:catAx>";
        }
        xml << "<c:valAx><c:axId val=\"2\"/><c:scaling><c:orientation val=\"minMax\"/></c:scaling><c:delete val=\"0\"/><c:axPos val=\"l\"/>";
        if (!chart.yAxisTitle().empty()) xml << "<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>" << xmlEscape(chart.yAxisTitle()) << "</a:t></a:r></a:p></c:rich></c:tx><c:overlay val=\"0\"/></c:title>";
        xml << "<c:numFmt formatCode=\"General\" sourceLinked=\"1\"/><c:majorGridlines/><c:majorTickMark val=\"none\"/><c:minorTickMark val=\"none\"/><c:tickLblPos val=\"nextTo\"/><c:crossAx val=\"1\"/><c:crosses val=\"autoZero\"/><c:crossBetween val=\"between\"/></c:valAx>";
        if (threeAxis)
            xml << "<c:serAx><c:axId val=\"3\"/><c:scaling><c:orientation val=\"minMax\"/></c:scaling><c:delete val=\"0\"/><c:axPos val=\"r\"/><c:majorTickMark val=\"none\"/><c:minorTickMark val=\"none\"/><c:crossAx val=\"1\"/><c:crosses val=\"autoZero\"/></c:serAx>";
    }
    if (chart.dataTable().present) xml << generatedDataTableXml(chart.dataTable(), strict);
    xml << "</c:plotArea>";
    if (chart.showLegend()) xml << "<c:legend><c:legendPos val=\"" << xmlEscape(chart.legendPosition()) << "\"/></c:legend>";
    xml << "</c:chart></c:chartSpace>";
    return xml.str();
}


std::string generatedChartTitleXml(const std::string& title, bool prefixed, bool strict) {
    const auto c = prefixed ? "c:" : "";
    const auto drawingMain = strict ? "http://purl.oclc.org/ooxml/drawingml/main"
                                    : "http://schemas.openxmlformats.org/drawingml/2006/main";
    std::ostringstream xml;
    xml << '<' << c << "title><" << c << "tx><" << c << "rich xmlns:a=\"" << drawingMain
        << "\"><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>" << xmlEscape(title)
        << "</a:t></a:r></a:p></" << c << "rich></" << c << "tx><" << c
        << "overlay val=\"0\"/></" << c << "title>";
    return xml.str();
}

std::string chartSolidFillXml(const xlpp::ChartColor& color, bool declareNamespace = false);

std::string chartRichTextTxXml(const xlpp::ChartRichText& richText, bool prefixed, bool strict) {
    const auto c = prefixed ? "c:" : "";
    const auto drawingMain = strict ? "http://purl.oclc.org/ooxml/drawingml/main"
                                    : "http://schemas.openxmlformats.org/drawingml/2006/main";
    std::string xml = "<" + std::string(c) + "tx><" + std::string(c) + "rich xmlns:a=\"" + drawingMain +
                      "\"><a:bodyPr/><a:lstStyle/><a:p>";
    for (const auto& run : richText.runs) {
        xml += "<a:r>";
        if (run.bold || run.italic || run.fontSizePoints > 0.0 || !run.typeface.empty() || run.color.present()) {
            xml += "<a:rPr";
            if (run.bold) xml += " b=\"1\"";
            if (run.italic) xml += " i=\"1\"";
            if (run.fontSizePoints > 0.0)
                xml += " sz=\"" + std::to_string(static_cast<long long>(std::llround(run.fontSizePoints * 100.0))) + "\"";
            xml += ">";
            if (run.color.present()) xml += chartSolidFillXml(run.color, false);
            if (!run.typeface.empty()) xml += "<a:latin typeface=\"" + xmlEscape(run.typeface) + "\"/>";
            xml += "</a:rPr>";
        }
        xml += "<a:t>" + xmlEscape(run.text) + "</a:t></a:r>";
    }
    xml += "</a:p></" + std::string(c) + "rich></" + std::string(c) + "tx>";
    return xml;
}

bool patchImportedChartTitleRichText(std::string& chartXmlText, const xlpp::ChartRichText& richText) {
    if (!richText.present || richText.runs.empty()) return false;
    const auto chartNodes = drawingTags(chartXmlText, "c:chart", "chart");
    if (chartNodes.empty()) return false;
    const auto originalChart = chartNodes.front();
    auto chartNode = originalChart;
    auto plotPosition = chartNode.find("<c:plotArea");
    if (plotPosition == std::string::npos) plotPosition = chartNode.find("<plotArea");
    if (plotPosition == std::string::npos) return false;
    std::string titleNode;
    for (const auto& candidate : drawingTags(chartNode, "c:title", "title")) {
        const auto position = chartNode.find(candidate);
        if (position < plotPosition) { titleNode = candidate; break; }
    }
    const bool prefixed = chartNode.find("<c:chart") != std::string::npos;
    const bool strict = chartXmlText.find("http://purl.oclc.org/ooxml/drawingml/chart") != std::string::npos;
    const auto txXml = chartRichTextTxXml(richText, prefixed, strict);
    if (titleNode.empty()) {
        const auto c = prefixed ? "c:" : "";
        const auto generated = "<" + std::string(c) + "title>" + txXml + "<" + std::string(c) +
                               "overlay val=\"0\"/></" + std::string(c) + "title>";
        chartNode.insert(plotPosition, generated);
    } else {
        auto patchedTitle = titleNode;
        const auto txNodes = drawingTags(patchedTitle, "c:tx", "tx");
        if (!txNodes.empty()) {
            const auto position = patchedTitle.find(txNodes.front());
            if (position == std::string::npos) return false;
            patchedTitle.replace(position, txNodes.front().size(), txXml);
        } else {
            const auto close = patchedTitle.rfind("</");
            if (close == std::string::npos) return false;
            patchedTitle.insert(close, txXml);
        }
        const auto position = chartNode.find(titleNode);
        if (position == std::string::npos) return false;
        chartNode.replace(position, titleNode.size(), patchedTitle);
    }
    const auto position = chartXmlText.find(originalChart);
    if (position == std::string::npos) return false;
    chartXmlText.replace(position, originalChart.size(), chartNode);
    return true;
}

bool patchImportedChartTitle(std::string& chartXmlText, const std::string& title) {
    const auto chartNodes = drawingTags(chartXmlText, "c:chart", "chart");
    if (chartNodes.empty()) return false;
    const auto originalChart = chartNodes.front();
    auto chartNode = originalChart;
    auto plotPosition = chartNode.find("<c:plotArea");
    if (plotPosition == std::string::npos) plotPosition = chartNode.find("<plotArea");
    if (plotPosition == std::string::npos) return false;

    std::string titleNode;
    for (const auto& candidate : drawingTags(chartNode, "c:title", "title")) {
        const auto position = chartNode.find(candidate);
        if (position < plotPosition) { titleNode = candidate; break; }
    }
    if (title.empty()) {
        if (!titleNode.empty()) {
            const auto position = chartNode.find(titleNode);
            if (position == std::string::npos) return false;
            chartNode.erase(position, titleNode.size());
        }
    } else if (!titleNode.empty()) {
        auto patchedTitle = titleNode;
        if (!replaceSimpleElementText(patchedTitle, "a:t", "t", title) &&
            !replaceSimpleElementText(patchedTitle, "c:v", "v", title)) {
            const auto prefixed = chartNode.find("<c:chart") != std::string::npos;
            const auto strict = chartXmlText.find("http://purl.oclc.org/ooxml/drawingml/chart") != std::string::npos;
            patchedTitle = generatedChartTitleXml(title, prefixed, strict);
        }
        const auto position = chartNode.find(titleNode);
        if (position == std::string::npos) return false;
        chartNode.replace(position, titleNode.size(), patchedTitle);
    } else {
        const auto prefixed = chartNode.find("<c:chart") != std::string::npos;
        const auto strict = chartXmlText.find("http://purl.oclc.org/ooxml/drawingml/chart") != std::string::npos;
        chartNode.insert(plotPosition, generatedChartTitleXml(title, prefixed, strict));
    }
    const auto position = chartXmlText.find(originalChart);
    if (position == std::string::npos) return false;
    chartXmlText.replace(position, originalChart.size(), chartNode);
    return true;
}

bool patchAxisTitleNode(std::string& axis, const std::string& chartXmlText, const std::string& title) {
    const auto titleNodes = drawingTags(axis, "c:title", "title");
    if (title.empty()) {
        if (!titleNodes.empty()) {
            const auto position = axis.find(titleNodes.front());
            if (position == std::string::npos) return false;
            axis.erase(position, titleNodes.front().size());
        }
    } else if (!titleNodes.empty()) {
        auto patched = titleNodes.front();
        if (!replaceSimpleElementText(patched, "a:t", "t", title) &&
            !replaceSimpleElementText(patched, "c:v", "v", title)) {
            const bool prefixed = axis.find("<c:") != std::string::npos;
            const bool strict = chartXmlText.find("http://purl.oclc.org/ooxml/drawingml/chart") != std::string::npos;
            patched = generatedChartTitleXml(title, prefixed, strict);
        }
        const auto position = axis.find(titleNodes.front());
        if (position == std::string::npos) return false;
        axis.replace(position, titleNodes.front().size(), patched);
    } else {
        const bool prefixed = axis.find("<c:") != std::string::npos;
        const bool strict = chartXmlText.find("http://purl.oclc.org/ooxml/drawingml/chart") != std::string::npos;
        const auto generated = generatedChartTitleXml(title, prefixed, strict);
        std::size_t insertion = std::string::npos;
        for (const auto& candidate : std::array<std::pair<const char*, const char*>, 7>{
                 std::pair{"c:numFmt", "numFmt"}, std::pair{"c:majorTickMark", "majorTickMark"},
                 std::pair{"c:minorTickMark", "minorTickMark"}, std::pair{"c:tickLblPos", "tickLblPos"},
                 std::pair{"c:spPr", "spPr"}, std::pair{"c:txPr", "txPr"},
                 std::pair{"c:crossAx", "crossAx"}}) {
            const auto nodes = drawingTags(axis, candidate.first, candidate.second);
            if (nodes.empty()) continue;
            const auto position = axis.find(nodes.front());
            if (position != std::string::npos) insertion = std::min(insertion, position);
        }
        if (insertion == std::string::npos) {
            const auto close = axis.rfind("</");
            if (close == std::string::npos) return false;
            insertion = close;
        }
        axis.insert(insertion, generated);
    }
    return true;
}

bool patchImportedAxisTitle(std::string& chartXmlText,
                            const char* prefixedAxis,
                            const char* localAxis,
                            std::size_t axisIndex,
                            const std::string& title) {
    const auto axes = drawingTags(chartXmlText, prefixedAxis, localAxis);
    if (axisIndex >= axes.size()) return false;
    const auto originalAxis = axes[axisIndex];
    auto axis = originalAxis;
    if (!patchAxisTitleNode(axis, chartXmlText, title)) return false;
    const auto position = chartXmlText.find(originalAxis);
    if (position == std::string::npos) return false;
    chartXmlText.replace(position, originalAxis.size(), axis);
    return true;
}

bool patchImportedAxisTitleById(std::string& chartXmlText, std::uint64_t axisId, const std::string& title) {
    if (axisId == 0) return false;
    const auto tryAxis = [&](const char* prefixedAxis, const char* localAxis) -> bool {
        for (const auto& originalAxis : drawingTags(chartXmlText, prefixedAxis, localAxis)) {
            const auto ids = drawingTags(originalAxis, "c:axId", "axId");
            if (ids.empty()) continue;
            const auto value = xlpp::internal::attribute(ids.front(), "val");
            std::uint64_t parsed = 0;
            try { if (!value.empty()) parsed = std::stoull(value); } catch (...) { continue; }
            if (parsed != axisId) continue;
            auto axis = originalAxis;
            if (!patchAxisTitleNode(axis, chartXmlText, title)) return false;
            const auto position = chartXmlText.find(originalAxis);
            if (position == std::string::npos) return false;
            chartXmlText.replace(position, originalAxis.size(), axis);
            return true;
        }
        return false;
    };
    return tryAxis("c:catAx", "catAx") || tryAxis("c:valAx", "valAx") ||
           tryAxis("c:dateAx", "dateAx") || tryAxis("c:serAx", "serAx");
}

bool patchOpeningTagAttribute(std::string& node, const std::string& name, const std::string& value, bool removeWhenEmpty);
void removeDrawingChild(std::string& container, const char* prefixed, const char* local);
bool patchOrInsertValChild(std::string& container, const char* prefixed, const char* local, const std::string& value, bool insertWhenMissing);
bool patchImportedChartStyle(std::string& chartXmlText, const std::string& style) {
    if (style.empty()) return false;
    const auto existing = drawingTags(chartXmlText, "c:style", "style");
    if (!existing.empty()) {
        auto node = existing.front();
        if (!patchOpeningTagAttribute(node, "val", style, false)) return false;
        const auto pos = chartXmlText.find(existing.front());
        if (pos == std::string::npos) return false;
        chartXmlText.replace(pos, existing.front().size(), node);
        return true;
    }
    const auto charts = drawingTags(chartXmlText, "c:chart", "chart");
    if (charts.empty()) return false;
    const auto pos = chartXmlText.find(charts.front());
    if (pos == std::string::npos) return false;
    const bool prefixed = chartXmlText.find("<c:chartSpace") != std::string::npos;
    chartXmlText.insert(pos, std::string("<") + (prefixed ? "c:" : "") + "style val=\"" + xmlEscape(style) + "\"/>");
    return true;
}

bool patchChartLineFormatInSpPr(std::string& spPr, const xlpp::ChartLineFormat& format);
bool patchChartFillFormatInSpPr(std::string& spPr, const xlpp::ChartFillFormat& format);
bool ensureChartSpPr(std::string& owner, std::string& spPr, const std::string& beforeXml);
bool patchNestedLineFormat(std::string& owner, const xlpp::ChartLineFormat& format);

bool patchAxisNodeById(std::string& chartXmlText, std::uint64_t axisId,
                       const std::function<bool(std::string&)>& patcher) {
    if (axisId == 0) return false;
    for (const auto& pair : std::array<std::pair<const char*, const char*>, 4>{{
             {"c:catAx", "catAx"}, {"c:valAx", "valAx"}, {"c:dateAx", "dateAx"}, {"c:serAx", "serAx"}}}) {
        for (const auto& originalAxis : drawingTags(chartXmlText, pair.first, pair.second)) {
            const auto ids = drawingTags(originalAxis, "c:axId", "axId");
            if (ids.empty()) continue;
            std::uint64_t parsed = 0;
            try { const auto value=xlpp::internal::attribute(ids.front(),"val"); if(!value.empty()) parsed=std::stoull(value); } catch (...) { continue; }
            if (parsed != axisId) continue;
            auto axis = originalAxis;
            if (!patcher(axis)) return false;
            const auto position = chartXmlText.find(originalAxis);
            if (position == std::string::npos) return false;
            chartXmlText.replace(position, originalAxis.size(), axis);
            return true;
        }
    }
    return false;
}

bool patchImportedAxisTitleRichTextById(std::string& chartXmlText, std::uint64_t axisId, const xlpp::ChartRichText& richText) {
    if (!richText.present || richText.runs.empty()) return false;
    return patchAxisNodeById(chartXmlText, axisId, [&](std::string& axis) {
        const bool prefixed = axis.find("<c:") != std::string::npos;
        const bool strict = chartXmlText.find("http://purl.oclc.org/ooxml/drawingml/chart") != std::string::npos;
        const auto c = prefixed ? "c:" : "";
        const auto txXml = chartRichTextTxXml(richText, prefixed, strict);
        const auto titles = drawingTags(axis, "c:title", "title");
        if (!titles.empty()) {
            auto title = titles.front();
            const auto tx = drawingTags(title, "c:tx", "tx");
            if (!tx.empty()) { const auto pos=title.find(tx.front()); if(pos==std::string::npos) return false; title.replace(pos,tx.front().size(),txXml); }
            else { const auto close=title.rfind("</"); if(close==std::string::npos) return false; title.insert(close,txXml); }
            const auto pos=axis.find(titles.front()); if(pos==std::string::npos) return false; axis.replace(pos,titles.front().size(),title);
        } else {
            const auto generated="<"+std::string(c)+"title>"+txXml+"<"+std::string(c)+"overlay val=\"0\"/></"+std::string(c)+"title>";
            std::size_t insertion=axis.rfind("</");
            const auto cross=drawingTags(axis,"c:crossAx","crossAx"); if(!cross.empty()) insertion=axis.find(cross.front());
            if(insertion==std::string::npos) return false; axis.insert(insertion,generated);
        }
        return true;
    });
}

bool patchImportedAxisNumberFormat(std::string& chartXmlText, std::uint64_t axisId, const std::string& formatCode, bool sourceLinked) {
    return patchAxisNodeById(chartXmlText, axisId, [&](std::string& axis) {
        auto nodes=drawingTags(axis,"c:numFmt","numFmt");
        if(!nodes.empty()) { auto node=nodes.front(); if(!patchOpeningTagAttribute(node,"formatCode",formatCode,false) || !patchOpeningTagAttribute(node,"sourceLinked",sourceLinked?"1":"0",false)) return false; const auto pos=axis.find(nodes.front()); if(pos==std::string::npos) return false; axis.replace(pos,nodes.front().size(),node); return true; }
        const bool prefixed=axis.find("<c:")!=std::string::npos; const auto c=prefixed?"c:":"";
        std::size_t insertion=axis.rfind("</"); const auto major=drawingTags(axis,"c:majorTickMark","majorTickMark"); if(!major.empty()) insertion=axis.find(major.front()); const auto cross=drawingTags(axis,"c:crossAx","crossAx"); if(insertion==std::string::npos && !cross.empty()) insertion=axis.find(cross.front()); if(insertion==std::string::npos) return false;
        axis.insert(insertion,"<"+std::string(c)+"numFmt formatCode=\""+xmlEscape(formatCode)+"\" sourceLinked=\""+(sourceLinked?"1":"0")+"\"/>"); return true;
    });
}

bool patchImportedAxisTicks(std::string& chartXmlText, std::uint64_t axisId, const std::string& major, const std::string& minor, const std::string& labelPos) {
    return patchAxisNodeById(chartXmlText, axisId, [&](std::string& axis) {
        if(!major.empty() && !patchOrInsertValChild(axis,"c:majorTickMark","majorTickMark",major,true)) return false;
        if(!minor.empty() && !patchOrInsertValChild(axis,"c:minorTickMark","minorTickMark",minor,true)) return false;
        if(!labelPos.empty() && !patchOrInsertValChild(axis,"c:tickLblPos","tickLblPos",labelPos,true)) return false;
        return true;
    });
}

bool patchImportedAxisUnits(std::string& chartXmlText, std::uint64_t axisId, double majorUnit, double minorUnit) {
    return patchAxisNodeById(chartXmlText, axisId, [&](std::string& axis) {
        std::ostringstream major; major<<std::setprecision(15)<<majorUnit;
        if(!patchOrInsertValChild(axis,"c:majorUnit","majorUnit",major.str(),true)) return false;
        if(minorUnit>0.0) { std::ostringstream minor; minor<<std::setprecision(15)<<minorUnit; if(!patchOrInsertValChild(axis,"c:minorUnit","minorUnit",minor.str(),true)) return false; }
        return true;
    });
}

bool patchScalingValChild(std::string& scaling, const char* prefixed, const char* local, const std::string& value) {
    const auto nodes = drawingTags(scaling, prefixed, local);
    const bool usePrefix = scaling.find("<c:scaling") != std::string::npos;
    const auto c = usePrefix ? "c:" : "";
    const auto generated = "<" + std::string(c) + local + " val=\"" + xmlEscape(value) + "\"/>";
    if (!nodes.empty()) {
        const auto pos = scaling.find(nodes.front());
        if (pos == std::string::npos) return false;
        scaling.replace(pos, nodes.front().size(), generated);
        return true;
    }
    const auto ext = drawingTags(scaling, "c:extLst", "extLst");
    const auto insertion = !ext.empty() ? scaling.find(ext.front()) : scaling.rfind("</");
    if (insertion == std::string::npos) return false;
    scaling.insert(insertion, generated);
    return true;
}

bool patchImportedAxisScaling(std::string& chartXmlText, std::uint64_t axisId, const xlpp::ChartAxisScaling& scalingValue) {
    return patchAxisNodeById(chartXmlText, axisId, [&](std::string& axis) {
        auto nodes = drawingTags(axis, "c:scaling", "scaling");
        std::string scaling;
        if (!nodes.empty()) scaling = nodes.front();
        else {
            const bool prefixed = axis.find("<c:") != std::string::npos;
            const auto c = prefixed ? "c:" : "";
            scaling = "<" + std::string(c) + "scaling></" + std::string(c) + "scaling>";
        }
        if (scalingValue.hasLogBase) {
            std::ostringstream value; value << std::setprecision(15) << scalingValue.logBase;
            if (!patchScalingValChild(scaling, "c:logBase", "logBase", value.str())) return false;
        } else removeDrawingChild(scaling, "c:logBase", "logBase");
        if (!patchScalingValChild(scaling, "c:orientation", "orientation", scalingValue.reverseOrder ? "maxMin" : "minMax")) return false;
        if (scalingValue.hasMaximum) {
            std::ostringstream value; value << std::setprecision(15) << scalingValue.maximum;
            if (!patchScalingValChild(scaling, "c:max", "max", value.str())) return false;
        } else removeDrawingChild(scaling, "c:max", "max");
        if (scalingValue.hasMinimum) {
            std::ostringstream value; value << std::setprecision(15) << scalingValue.minimum;
            if (!patchScalingValChild(scaling, "c:min", "min", value.str())) return false;
        } else removeDrawingChild(scaling, "c:min", "min");
        if (!nodes.empty()) {
            const auto pos = axis.find(nodes.front());
            if (pos == std::string::npos) return false;
            axis.replace(pos, nodes.front().size(), scaling);
        } else {
            const auto ids = drawingTags(axis, "c:axId", "axId");
            if (ids.empty()) return false;
            const auto pos = axis.find(ids.front());
            if (pos == std::string::npos) return false;
            axis.insert(pos + ids.front().size(), scaling);
        }
        return true;
    });
}

bool patchImportedAxisCrossesAt(std::string& chartXmlText, std::uint64_t axisId, double crossesAt, bool clear) {
    return patchAxisNodeById(chartXmlText, axisId, [&](std::string& axis) {
        removeDrawingChild(axis, "c:crossesAt", "crossesAt");
        if (clear) return true;
        removeDrawingChild(axis, "c:crosses", "crosses");
        const bool prefixed = axis.find("<c:") != std::string::npos;
        const auto c = prefixed ? "c:" : "";
        std::ostringstream value; value << std::setprecision(15) << crossesAt;
        const auto generated = "<" + std::string(c) + "crossesAt val=\"" + value.str() + "\"/>";
        std::size_t insertion = axis.rfind("</");
        for (const auto& pair : std::array<std::pair<const char*, const char*>, 4>{{
                 {"c:crossBetween", "crossBetween"}, {"c:majorUnit", "majorUnit"}, {"c:minorUnit", "minorUnit"}, {"c:dispUnits", "dispUnits"}}}) {
            const auto following = drawingTags(axis, pair.first, pair.second);
            if (!following.empty()) { const auto pos=axis.find(following.front()); if(pos!=std::string::npos) insertion=std::min(insertion,pos); }
        }
        const auto ext = drawingTags(axis, "c:extLst", "extLst");
        if (!ext.empty()) { const auto pos=axis.find(ext.front()); if(pos!=std::string::npos) insertion=std::min(insertion,pos); }
        if (insertion == std::string::npos) return false;
        axis.insert(insertion, generated);
        return true;
    });
}

bool patchImportedAxisDisplayUnits(std::string& chartXmlText, std::uint64_t axisId, const xlpp::ChartDisplayUnits* units) {
    return patchAxisNodeById(chartXmlText, axisId, [&](std::string& axis) {
        const auto existing = drawingTags(axis, "c:dispUnits", "dispUnits");
        if (!units) {
            if (!existing.empty()) { const auto pos=axis.find(existing.front()); if(pos==std::string::npos) return false; axis.erase(pos,existing.front().size()); }
            return true;
        }
        const bool prefixed = axis.find("<c:") != std::string::npos;
        const bool strict = chartXmlText.find("http://purl.oclc.org/ooxml/drawingml/chart") != std::string::npos;
        const auto c = prefixed ? "c:" : "";
        std::string generated = "<" + std::string(c) + "dispUnits>";
        if (units->hasCustomUnit) {
            std::ostringstream value; value << std::setprecision(15) << units->customUnit;
            generated += "<" + std::string(c) + "custUnit val=\"" + value.str() + "\"/>";
        } else generated += "<" + std::string(c) + "builtInUnit val=\"" + xmlEscape(units->builtInUnit) + "\"/>";
        if (units->showLabel) {
            generated += "<" + std::string(c) + "dispUnitsLbl>";
            if (units->labelRichText.present && !units->labelRichText.runs.empty()) generated += chartRichTextTxXml(units->labelRichText, prefixed, strict);
            generated += "</" + std::string(c) + "dispUnitsLbl>";
        }
        generated += "</" + std::string(c) + "dispUnits>";
        if (!existing.empty()) {
            const auto pos = axis.find(existing.front()); if(pos==std::string::npos) return false; axis.replace(pos,existing.front().size(),generated); return true;
        }
        const auto ext = drawingTags(axis, "c:extLst", "extLst");
        const auto insertion = !ext.empty() ? axis.find(ext.front()) : axis.rfind("</");
        if (insertion == std::string::npos) return false;
        axis.insert(insertion, generated);
        return true;
    });
}

bool patchImportedAxisCrossing(std::string& chartXmlText, std::uint64_t axisId, const std::string& crosses, const std::string& crossBetween) {
    return patchAxisNodeById(chartXmlText, axisId, [&](std::string& axis) {
        if(!crosses.empty()) removeDrawingChild(axis,"c:crossesAt","crossesAt");
        if(!crosses.empty() && !patchOrInsertValChild(axis,"c:crosses","crosses",crosses,true)) return false;
        if(!crossBetween.empty() && !patchOrInsertValChild(axis,"c:crossBetween","crossBetween",crossBetween,true)) return false;
        return true;
    });
}

bool patchImportedAxisLineFormat(std::string& chartXmlText, std::uint64_t axisId, const xlpp::ChartLineFormat& format) {
    return patchAxisNodeById(chartXmlText, axisId, [&](std::string& axis) {
        auto spPr=axisDirectSpPr(axis); if(spPr.empty() && !ensureChartSpPr(axis,spPr,{})) return false; auto patched=spPr; if(!patchChartLineFormatInSpPr(patched,format)) return false; const auto pos=axis.find(spPr); if(pos==std::string::npos) return false; axis.replace(pos,spPr.size(),patched); return true;
    });
}

bool patchImportedAxisGridlineFormat(std::string& chartXmlText, std::uint64_t axisId, bool major, const xlpp::ChartLineFormat& format) {
    return patchAxisNodeById(chartXmlText, axisId, [&](std::string& axis) {
        const char* pref=major?"c:majorGridlines":"c:minorGridlines"; const char* local=major?"majorGridlines":"minorGridlines";
        auto grids=drawingTags(axis,pref,local); std::string grid;
        if(!grids.empty()) grid=grids.front(); else { const bool prefixed=axis.find("<c:")!=std::string::npos; const auto c=prefixed?"c:":""; grid="<"+std::string(c)+local+"></"+std::string(c)+local+">"; }
        if(!patchNestedLineFormat(grid,format)) return false;
        if(!grids.empty()) { const auto pos=axis.find(grids.front()); if(pos==std::string::npos) return false; axis.replace(pos,grids.front().size(),grid); }
        else { const auto title=drawingTags(axis,"c:title","title"); std::size_t pos=!title.empty()?axis.find(title.front()):axis.rfind("</"); if(pos==std::string::npos) return false; axis.insert(pos,grid); }
        return true;
    });
}

bool removeImportedAxisGridlines(std::string& chartXmlText, std::uint64_t axisId, bool major) {
    return patchAxisNodeById(chartXmlText, axisId, [&](std::string& axis) {
        removeDrawingChild(axis, major ? "c:majorGridlines" : "c:minorGridlines", major ? "majorGridlines" : "minorGridlines");
        return true;
    });
}

std::string chartSpaceDirectSpPr(const std::string& chartXmlText) {
    const auto candidates = drawingTags(chartXmlText, "c:spPr", "spPr");
    if (candidates.empty()) return {};
    const auto charts = drawingTags(chartXmlText, "c:chart", "chart");
    for (const auto& candidate : candidates) {
        if (std::none_of(charts.begin(), charts.end(), [&](const auto& chart) { return chart.find(candidate) != std::string::npos; })) return candidate;
    }
    return {};
}

std::string plotAreaDirectSpPr(const std::string& plotArea) {
    const auto candidates = drawingTags(plotArea, "c:spPr", "spPr");
    if (candidates.empty()) return {};
    std::vector<std::string> nested;
    for (const auto& pair : std::array<std::pair<const char*, const char*>, 20>{{
             {"c:barChart","barChart"},{"c:lineChart","lineChart"},{"c:pieChart","pieChart"},{"c:scatterChart","scatterChart"},
             {"c:doughnutChart","doughnutChart"},{"c:radarChart","radarChart"},{"c:areaChart","areaChart"},{"c:bubbleChart","bubbleChart"},{"c:stockChart","stockChart"},
             {"c:bar3DChart","bar3DChart"},{"c:line3DChart","line3DChart"},{"c:area3DChart","area3DChart"},{"c:pie3DChart","pie3DChart"},
             {"c:surfaceChart","surfaceChart"},{"c:surface3DChart","surface3DChart"},
             {"c:catAx","catAx"},{"c:valAx","valAx"},{"c:dateAx","dateAx"},{"c:serAx","serAx"},{"c:dTable","dTable"}}}) {
        const auto nodes = drawingTags(plotArea, pair.first, pair.second); nested.insert(nested.end(), nodes.begin(), nodes.end());
    }
    for (const auto& candidate : candidates)
        if (std::none_of(nested.begin(), nested.end(), [&](const auto& owner){ return owner.find(candidate) != std::string::npos; })) return candidate;
    return {};
}

bool patchImportedAreaFormat(std::string& chartXmlText, bool chartArea, const xlpp::ChartLineFormat* line, const xlpp::ChartFillFormat* fill) {
    std::string owner;
    std::string originalOwner;
    if (chartArea) owner = chartXmlText;
    else {
        const auto plots = drawingTags(chartXmlText, "c:plotArea", "plotArea");
        if (plots.empty()) return false;
        originalOwner = plots.front(); owner = originalOwner;
    }
    auto spPr = chartArea ? chartSpaceDirectSpPr(owner) : plotAreaDirectSpPr(owner);
    if (spPr.empty()) {
        const bool prefixed = owner.find("<c:") != std::string::npos;
        const auto c = prefixed ? "c:" : "";
        spPr = "<" + std::string(c) + "spPr></" + std::string(c) + "spPr>";
        std::size_t insertion = std::string::npos;
        if (chartArea) {
            const auto charts = drawingTags(owner, "c:chart", "chart");
            if (charts.empty()) return false;
            const auto chartPos = owner.find(charts.front());
            if (chartPos == std::string::npos) return false;
            insertion = chartPos + charts.front().size();
        } else {
            const auto ext = drawingTags(owner, "c:extLst", "extLst");
            insertion = !ext.empty() ? owner.find(ext.back()) : owner.rfind("</");
        }
        if (insertion == std::string::npos) return false;
        owner.insert(insertion, spPr);
    }
    auto patched = spPr;
    if (line && !patchChartLineFormatInSpPr(patched, *line)) return false;
    if (fill && !patchChartFillFormatInSpPr(patched, *fill)) return false;
    const auto spPos = owner.find(spPr);
    if (spPos == std::string::npos) return false;
    owner.replace(spPos, spPr.size(), patched);
    if (chartArea) { chartXmlText = std::move(owner); return true; }
    const auto position = chartXmlText.find(originalOwner);
    if (position == std::string::npos) return false;
    chartXmlText.replace(position, originalOwner.size(), owner);
    return true;
}

std::string chartManualLayoutXml(const xlpp::ChartManualLayout& layout, bool prefixed) {
    const auto c=prefixed?"c:":""; std::ostringstream xml; xml<<"<"<<c<<"layout><"<<c<<"manualLayout>";
    const auto addText=[&](const char* tag,const std::string& value){ if(!value.empty()) xml<<"<"<<c<<tag<<" val=\""<<xmlEscape(value)<<"\"/>"; };
    const auto addNum=[&](const char* tag,bool has,double value){ if(has) xml<<"<"<<c<<tag<<" val=\""<<std::setprecision(15)<<value<<"\"/>"; };
    addText("layoutTarget",layout.target); addText("xMode",layout.xMode); addText("yMode",layout.yMode); addText("wMode",layout.widthMode); addText("hMode",layout.heightMode);
    addNum("x",layout.hasX,layout.x); addNum("y",layout.hasY,layout.y); addNum("w",layout.hasWidth,layout.width); addNum("h",layout.hasHeight,layout.height);
    xml<<"</"<<c<<"manualLayout></"<<c<<"layout>"; return xml.str();
}

bool patchManualLayoutOwner(std::string& owner, const xlpp::ChartManualLayout& layout) {
    if(!layout.present) return false; const bool prefixed=owner.find("<c:")!=std::string::npos; const auto generated=chartManualLayoutXml(layout,prefixed); const auto layouts=drawingTags(owner,"c:layout","layout");
    if(!layouts.empty()) { const auto pos=owner.find(layouts.front()); if(pos==std::string::npos) return false; owner.replace(pos,layouts.front().size(),generated); }
    else { const auto openEnd=owner.find('>'); if(openEnd==std::string::npos) return false; owner.insert(openEnd+1,generated); }
    return true;
}

bool patchImportedPlotAreaLayout(std::string& chartXmlText, const xlpp::ChartManualLayout& layout) {
    const auto plots=drawingTags(chartXmlText,"c:plotArea","plotArea"); if(plots.empty()) return false; auto plot=plots.front(); if(!patchManualLayoutOwner(plot,layout)) return false; const auto pos=chartXmlText.find(plots.front()); if(pos==std::string::npos) return false; chartXmlText.replace(pos,plots.front().size(),plot); return true;
}

bool patchImportedLegendLayout(std::string& chartXmlText, const xlpp::ChartManualLayout& layout) {
    const auto legends=drawingTags(chartXmlText,"c:legend","legend"); if(legends.empty()) return false; auto legend=legends.front(); if(!patchManualLayoutOwner(legend,layout)) return false; const auto pos=chartXmlText.find(legends.front()); if(pos==std::string::npos) return false; chartXmlText.replace(pos,legends.front().size(),legend); return true;
}

bool patchImportedLegendOverlay(std::string& chartXmlText, bool overlay) {
    const auto legends=drawingTags(chartXmlText,"c:legend","legend"); if(legends.empty()) return false; auto legend=legends.front(); if(!patchOrInsertValChild(legend,"c:overlay","overlay",overlay?"1":"0",true)) return false; const auto pos=chartXmlText.find(legends.front()); if(pos==std::string::npos) return false; chartXmlText.replace(pos,legends.front().size(),legend); return true;
}

bool patchImportedLegendFormat(std::string& chartXmlText, const xlpp::ChartLineFormat* line, const xlpp::ChartFillFormat* fill) {
    const auto legends=drawingTags(chartXmlText,"c:legend","legend"); if(legends.empty()) return false; auto legend=legends.front(); auto spPrNodes=drawingTags(legend,"c:spPr","spPr"); std::string spPr=spPrNodes.empty()?std::string{}:spPrNodes.front(); if(spPr.empty() && !ensureChartSpPr(legend,spPr,{})) return false; auto patched=spPr; if(line && !patchChartLineFormatInSpPr(patched,*line)) return false; if(fill && !patchChartFillFormatInSpPr(patched,*fill)) return false; const auto spos=legend.find(spPr); if(spos==std::string::npos) return false; legend.replace(spos,spPr.size(),patched); const auto pos=chartXmlText.find(legends.front()); if(pos==std::string::npos) return false; chartXmlText.replace(pos,legends.front().size(),legend); return true;
}

bool patchImportedChartLegend(std::string& chartXmlText, bool show, const std::string& legendPosition) {
    const auto chartNodes = drawingTags(chartXmlText, "c:chart", "chart");
    if (chartNodes.empty()) return false;
    const auto originalChart = chartNodes.front();
    auto chart = originalChart;
    const auto legends = drawingTags(chart, "c:legend", "legend");
    if (!show) {
        if (!legends.empty()) {
            const auto position = chart.find(legends.front());
            if (position == std::string::npos) return false;
            chart.erase(position, legends.front().size());
        }
    } else if (!legends.empty()) {
        auto legend = legends.front();
        const auto positionNodes = drawingTags(legend, "c:legendPos", "legendPos");
        if (!positionNodes.empty()) {
            auto node = positionNodes.front();
            const auto val = xlpp::internal::attribute(node, "val");
            const auto attr = std::string("val=\"") + val + "\"";
            const auto attrPosition = node.find(attr);
            if (attrPosition == std::string::npos) return false;
            node.replace(attrPosition, attr.size(), "val=\"" + xmlEscape(legendPosition) + "\"");
            const auto nodePosition = legend.find(positionNodes.front());
            if (nodePosition == std::string::npos) return false;
            legend.replace(nodePosition, positionNodes.front().size(), node);
        } else {
            const bool prefixed = legend.find("<c:legend") != std::string::npos;
            const auto openEnd = legend.find('>');
            if (openEnd == std::string::npos) return false;
            const auto c = prefixed ? "c:" : "";
            legend.insert(openEnd + 1, "<" + std::string(c) + "legendPos val=\"" + xmlEscape(legendPosition) + "\"/>");
        }
        const auto position = chart.find(legends.front());
        if (position == std::string::npos) return false;
        chart.replace(position, legends.front().size(), legend);
    } else {
        const bool prefixed = chart.find("<c:chart") != std::string::npos;
        const auto c = prefixed ? "c:" : "";
        const std::string legend = "<" + std::string(c) + "legend><" + std::string(c) +
            "legendPos val=\"" + xmlEscape(legendPosition) + "\"/><" + std::string(c) +
            "layout/><" + std::string(c) + "overlay val=\"0\"/></" + std::string(c) + "legend>";
        const auto plotAreas = drawingTags(chart, "c:plotArea", "plotArea");
        if (plotAreas.empty()) return false;
        const auto plotPosition = chart.find(plotAreas.front());
        if (plotPosition == std::string::npos) return false;
        chart.insert(plotPosition + plotAreas.front().size(), legend);
    }
    const auto position = chartXmlText.find(originalChart);
    if (position == std::string::npos) return false;
    chartXmlText.replace(position, originalChart.size(), chart);
    return true;
}

bool patchValAttribute(std::string& node, const std::string& value) {
    const auto oldValue = xlpp::internal::attribute(node, "val");
    if (!oldValue.empty()) {
        const auto token = std::string("val=\"") + oldValue + "\"";
        const auto position = node.find(token);
        if (position == std::string::npos) return false;
        node.replace(position, token.size(), "val=\"" + xmlEscape(value) + "\"");
        return true;
    }
    const auto close = node.find("/>");
    const auto openEnd = node.find('>');
    const auto insertion = close != std::string::npos ? close : openEnd;
    if (insertion == std::string::npos) return false;
    node.insert(insertion, " val=\"" + xmlEscape(value) + "\"");
    return true;
}

bool patchOrInsertValChild(std::string& container, const char* prefixed, const char* local,
                           const std::string& value, bool insertWhenMissing = true) {
    const auto nodes = drawingTags(container, prefixed, local);
    if (!nodes.empty()) {
        auto patched = nodes.front();
        if (!patchValAttribute(patched, value)) return false;
        const auto position = container.find(nodes.front());
        if (position == std::string::npos) return false;
        container.replace(position, nodes.front().size(), patched);
        return true;
    }
    if (!insertWhenMissing) return true;
    const bool prefixedContainer = container.find("<c:") != std::string::npos;
    const auto c = prefixedContainer ? "c:" : "";
    const auto close = container.rfind("</");
    if (close == std::string::npos) return false;
    container.insert(close, "<" + std::string(c) + local + " val=\"" + xmlEscape(value) + "\"/>");
    return true;
}

std::string dataLabelsAggregateMask(const std::string& dLbls) {
    auto mask = dLbls;
    std::size_t cursor = 0;
    for (const auto& point : drawingTags(dLbls, "c:dLbl", "dLbl")) {
        const auto position = dLbls.find(point, cursor);
        if (position == std::string::npos) continue;
        std::fill(mask.begin() + static_cast<std::ptrdiff_t>(position),
                  mask.begin() + static_cast<std::ptrdiff_t>(position + point.size()), ' ');
        cursor = position + point.size();
    }
    return mask;
}

bool patchOrInsertAggregateDataLabelVal(std::string& dLbls, const char* prefixed, const char* local,
                                        const std::string& value, bool insertWhenMissing = true) {
    const auto mask = dataLabelsAggregateMask(dLbls);
    const auto nodes = drawingTags(mask, prefixed, local);
    if (!nodes.empty()) {
        const auto position = mask.find(nodes.front());
        if (position == std::string::npos || position + nodes.front().size() > dLbls.size()) return false;
        auto patched = dLbls.substr(position, nodes.front().size());
        if (!patchValAttribute(patched, value)) return false;
        dLbls.replace(position, nodes.front().size(), patched);
        return true;
    }
    if (!insertWhenMissing) return true;
    const bool prefixedContainer = dLbls.find("<c:dLbls") != std::string::npos;
    const auto c = prefixedContainer ? "c:" : "";
    const auto close = dLbls.rfind("</");
    if (close == std::string::npos) return false;
    dLbls.insert(close, "<" + std::string(c) + local + " val=\"" + xmlEscape(value) + "\"/>");
    return true;
}

bool patchOrInsertAggregateDataLabelText(std::string& dLbls, const char* prefixed, const char* local,
                                         const std::string& value) {
    const auto mask = dataLabelsAggregateMask(dLbls);
    const auto nodes = drawingTags(mask, prefixed, local);
    const bool prefixedContainer = dLbls.find("<c:dLbls") != std::string::npos;
    const auto c = prefixedContainer ? "c:" : "";
    const auto generated = "<" + std::string(c) + local + ">" + xmlEscape(value) + "</" + std::string(c) + local + ">";
    if (!nodes.empty()) {
        const auto position = mask.find(nodes.front());
        if (position == std::string::npos || position + nodes.front().size() > dLbls.size()) return false;
        dLbls.replace(position, nodes.front().size(), generated);
        return true;
    }
    const auto close = dLbls.rfind("</");
    if (close == std::string::npos) return false;
    dLbls.insert(close, generated);
    return true;
}

void removeDrawingChild(std::string& container, const char* prefixed, const char* local) {
    for (;;) {
        const auto nodes = drawingTags(container, prefixed, local);
        if (nodes.empty()) return;
        const auto position = container.find(nodes.front());
        if (position == std::string::npos) return;
        container.erase(position, nodes.front().size());
    }
}

bool patchOpeningTagAttribute(std::string& node, const std::string& name, const std::string& value, bool removeWhenEmpty = false) {
    const auto openEnd = node.find('>');
    if (openEnd == std::string::npos) return false;
    const auto key = name + "=\"";
    auto position = node.find(key);
    if (position != std::string::npos && position < openEnd) {
        const auto valueStart = position + key.size();
        const auto valueEnd = node.find('"', valueStart);
        if (valueEnd == std::string::npos || valueEnd > openEnd) return false;
        if (removeWhenEmpty && value.empty()) {
            auto eraseStart = position;
            if (eraseStart > 0 && std::isspace(static_cast<unsigned char>(node[eraseStart - 1]))) --eraseStart;
            node.erase(eraseStart, valueEnd + 1 - eraseStart);
        } else node.replace(valueStart, valueEnd - valueStart, xmlEscape(value));
        return true;
    }
    if (removeWhenEmpty && value.empty()) return true;
    const auto insertion = node.find("/>") < openEnd ? node.find("/>") : openEnd;
    node.insert(insertion, " " + name + "=\"" + xmlEscape(value) + "\"");
    return true;
}

const char* chartColorTransformTag(xlpp::ChartColorTransform::Kind kind) {
    using Kind = xlpp::ChartColorTransform::Kind;
    switch (kind) {
    case Kind::Alpha: return "alpha";
    case Kind::AlphaMod: return "alphaMod";
    case Kind::AlphaOff: return "alphaOff";
    case Kind::Tint: return "tint";
    case Kind::Shade: return "shade";
    case Kind::LumMod: return "lumMod";
    case Kind::LumOff: return "lumOff";
    case Kind::SatMod: return "satMod";
    case Kind::SatOff: return "satOff";
    }
    return "alpha";
}

std::string chartColorElement(const xlpp::ChartColor& color) {
    using Kind = xlpp::ChartColor::Kind;
    if (!color.present()) return {};
    const char* tag = "srgbClr";
    switch (color.kind) {
    case Kind::SRgb: tag = "srgbClr"; break;
    case Kind::Scheme: tag = "schemeClr"; break;
    case Kind::System: tag = "sysClr"; break;
    case Kind::Preset: tag = "prstClr"; break;
    case Kind::Unknown: tag = "srgbClr"; break;
    case Kind::None: return {};
    }
    if (color.transforms.empty())
        return "<a:" + std::string(tag) + " val=\"" + xmlEscape(color.value) + "\"/>";
    std::string xml = "<a:" + std::string(tag) + " val=\"" + xmlEscape(color.value) + "\">";
    for (const auto& transform : color.transforms)
        xml += "<a:" + std::string(chartColorTransformTag(transform.kind)) + " val=\"" + std::to_string(transform.value) + "\"/>";
    xml += "</a:" + std::string(tag) + ">";
    return xml;
}

std::string chartSolidFillXml(const xlpp::ChartColor& color, bool declareNamespace) {
    if (!color.present()) return {};
    return "<a:solidFill" + std::string(declareNamespace ? " xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"" : "") +
           ">" + chartColorElement(color) + "</a:solidFill>";
}

std::string chartGradientFillXml(const xlpp::ChartFillFormat& format, bool declareNamespace = false) {
    if (format.gradientStops.empty()) return {};
    std::string xml = "<a:gradFill" + std::string(declareNamespace ? " xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"" : "") + "><a:gsLst>";
    for (const auto& stop : format.gradientStops)
        if (stop.color.present())
            xml += "<a:gs pos=\"" + std::to_string(std::clamp(stop.position, 0, 100000)) + "\">" +
                   chartColorElement(stop.color) + "</a:gs>";
    xml += "</a:gsLst>";
    if (std::isfinite(format.gradientAngleDegrees) && std::abs(format.gradientAngleDegrees) > 1e-12) {
        const auto angle = static_cast<long long>(std::llround(format.gradientAngleDegrees * 60000.0));
        xml += "<a:lin ang=\"" + std::to_string(angle) + "\" scaled=\"1\"/>";
    }
    xml += "</a:gradFill>";
    return xml;
}

std::string chartPatternFillXml(const xlpp::ChartFillFormat& format, bool declareNamespace = false) {
    if (format.pattern.empty()) return {};
    std::string xml = "<a:pattFill" + std::string(declareNamespace ? " xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"" : "") +
                      " prst=\"" + xmlEscape(format.pattern) + "\">";
    if (format.foregroundColor.present()) xml += "<a:fgClr>" + chartColorElement(format.foregroundColor) + "</a:fgClr>";
    if (format.backgroundColor.present()) xml += "<a:bgClr>" + chartColorElement(format.backgroundColor) + "</a:bgClr>";
    xml += "</a:pattFill>";
    return xml;
}

bool replaceFirstChild(std::string& container, const char* prefixed, const char* local, const std::string& replacement) {
    const auto nodes = drawingTags(container, prefixed, local);
    if (nodes.empty()) return false;
    const auto position = container.find(nodes.front());
    if (position == std::string::npos) return false;
    container.replace(position, nodes.front().size(), replacement);
    return true;
}

std::string directSpPrFillNode(const std::string& spPr, const char* prefixed, const char* local) {
    const auto lines = drawingTags(spPr, "a:ln", "ln");
    for (const auto& candidate : drawingTags(spPr, prefixed, local)) {
        if (std::none_of(lines.begin(), lines.end(), [&](const auto& line) { return line.find(candidate) != std::string::npos; }))
            return candidate;
    }
    return {};
}

bool patchChartLineFormatInSpPr(std::string& spPr, const xlpp::ChartLineFormat& format) {
    auto lines = drawingTags(spPr, "a:ln", "ln");
    std::string line;
    if (!lines.empty()) line = lines.front();
    else line = "<a:ln xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"></a:ln>";

    if (format.widthPoints > 0.0) {
        const auto widthEmu = static_cast<long long>(std::llround(format.widthPoints * 12700.0));
        if (!patchOpeningTagAttribute(line, "w", std::to_string(widthEmu))) return false;
    }
    if (!patchOpeningTagAttribute(line, "cap", format.cap, true)) return false;
    if (!patchOpeningTagAttribute(line, "cmpd", format.compound, true)) return false;

    removeDrawingChild(line, "a:noFill", "noFill");
    removeDrawingChild(line, "a:solidFill", "solidFill");
    if (format.noFill) {
        const auto close = line.rfind("</");
        if (close == std::string::npos) return false;
        line.insert(close, "<a:noFill/>");
    } else if (format.color.present()) {
        const auto dashes = drawingTags(line, "a:prstDash", "prstDash");
        const auto customDashes = drawingTags(line, "a:custDash", "custDash");
        std::size_t insertion = line.rfind("</");
        if (!dashes.empty()) insertion = line.find(dashes.front());
        else if (!customDashes.empty()) insertion = line.find(customDashes.front());
        if (insertion == std::string::npos) return false;
        line.insert(insertion, chartSolidFillXml(format.color));
    }

    removeDrawingChild(line, "a:prstDash", "prstDash");
    removeDrawingChild(line, "a:custDash", "custDash");
    if (!format.customDash.empty()) {
        std::string custom = "<a:custDash>";
        for (const auto& stop : format.customDash) {
            const auto d = static_cast<long long>(std::llround(std::max(0.0, stop.dash) * 1000.0));
            const auto sp = static_cast<long long>(std::llround(std::max(0.0, stop.space) * 1000.0));
            custom += "<a:ds d=\"" + std::to_string(d) + "\" sp=\"" + std::to_string(sp) + "\"/>";
        }
        custom += "</a:custDash>";
        const auto close = line.rfind("</");
        if (close == std::string::npos) return false;
        line.insert(close, custom);
    } else if (!format.dash.empty()) {
        const auto close = line.rfind("</");
        if (close == std::string::npos) return false;
        line.insert(close, "<a:prstDash val=\"" + xmlEscape(format.dash) + "\"/>");
    }

    removeDrawingChild(line, "a:round", "round");
    removeDrawingChild(line, "a:bevel", "bevel");
    removeDrawingChild(line, "a:miter", "miter");
    if (!format.join.empty()) {
        const auto close = line.rfind("</");
        if (close == std::string::npos) return false;
        if (format.join == "round") line.insert(close, "<a:round/>");
        else if (format.join == "bevel") line.insert(close, "<a:bevel/>");
        else if (format.join == "miter") line.insert(close, "<a:miter/>");
        else return false;
    }

    if (!lines.empty()) {
        const auto position = spPr.find(lines.front());
        if (position == std::string::npos) return false;
        spPr.replace(position, lines.front().size(), line);
    } else {
        const auto close = spPr.rfind("</");
        if (close == std::string::npos) return false;
        spPr.insert(close, line);
    }
    return true;
}

bool patchChartFillFormatInSpPr(std::string& spPr, const xlpp::ChartFillFormat& format) {
    for (const auto& pair : std::array<std::pair<const char*, const char*>, 4>{{
             {"a:noFill", "noFill"}, {"a:solidFill", "solidFill"}, {"a:gradFill", "gradFill"}, {"a:pattFill", "pattFill"}}}) {
        const auto existing = directSpPrFillNode(spPr, pair.first, pair.second);
        if (!existing.empty()) {
            const auto position = spPr.find(existing);
            if (position != std::string::npos) spPr.erase(position, existing.size());
        }
    }
    std::string generated;
    const auto kind = format.noFill ? xlpp::ChartFillFormat::Kind::NoFill : format.kind;
    if (kind == xlpp::ChartFillFormat::Kind::NoFill)
        generated = "<a:noFill xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"/>";
    else if (kind == xlpp::ChartFillFormat::Kind::Gradient)
        generated = chartGradientFillXml(format, true);
    else if (kind == xlpp::ChartFillFormat::Kind::Pattern)
        generated = chartPatternFillXml(format, true);
    else if (format.color.present())
        generated = chartSolidFillXml(format.color, true);
    if (!generated.empty()) {
        const auto lines = drawingTags(spPr, "a:ln", "ln");
        const auto insertion = !lines.empty() ? spPr.find(lines.front()) : spPr.rfind("</");
        if (insertion == std::string::npos) return false;
        spPr.insert(insertion, generated);
    }
    return true;
}

bool ensureChartSpPr(std::string& owner, std::string& spPr, const std::string& beforeXml = {}) {
    if (!spPr.empty()) return true;
    const bool prefixed = owner.find("<c:") != std::string::npos;
    const auto c = prefixed ? "c:" : "";
    spPr = "<" + std::string(c) + "spPr></" + std::string(c) + "spPr>";
    std::size_t insertion = std::string::npos;
    if (!beforeXml.empty()) insertion = owner.find(beforeXml);
    if (insertion == std::string::npos) insertion = owner.rfind("</");
    if (insertion == std::string::npos) return false;
    owner.insert(insertion, spPr);
    return true;
}

bool patchSeriesLineOrFill(std::string& chartXmlText, std::size_t seriesIndex,
                           const xlpp::ChartLineFormat* line, const xlpp::ChartFillFormat* fill) {
    const auto seriesNodes = drawingTags(chartXmlText, "c:ser", "ser");
    if (seriesIndex >= seriesNodes.size()) return false;
    const auto original = seriesNodes[seriesIndex];
    auto series = original;
    auto spPr = seriesDirectSpPr(series);
    if (spPr.empty()) {
        std::string before;
        for (const auto& pair : std::array<std::pair<const char*, const char*>, 3>{{{"c:marker", "marker"}, {"c:dLbls", "dLbls"}, {"c:trendline", "trendline"}}}) {
            const auto nodes = drawingTags(series, pair.first, pair.second);
            if (!nodes.empty()) { before = nodes.front(); break; }
        }
        if (!ensureChartSpPr(series, spPr, before)) return false;
    }
    auto patched = spPr;
    if (line && !patchChartLineFormatInSpPr(patched, *line)) return false;
    if (fill && !patchChartFillFormatInSpPr(patched, *fill)) return false;
    const auto spPos = series.find(spPr);
    if (spPos == std::string::npos) return false;
    series.replace(spPos, spPr.size(), patched);
    const auto position = chartXmlText.find(original);
    if (position == std::string::npos) return false;
    chartXmlText.replace(position, original.size(), series);
    return true;
}

bool patchSeriesMarkerFormat(std::string& chartXmlText, std::size_t seriesIndex, const xlpp::ChartMarkerFormat& format) {
    const auto seriesNodes = drawingTags(chartXmlText, "c:ser", "ser");
    if (seriesIndex >= seriesNodes.size()) return false;
    const auto original = seriesNodes[seriesIndex];
    auto series = original;
    auto markers = drawingTags(series, "c:marker", "marker");
    std::string marker;
    if (!markers.empty()) marker = markers.front();
    else {
        const bool prefixed = series.find("<c:ser") != std::string::npos;
        const auto c = prefixed ? "c:" : "";
        marker = "<" + std::string(c) + "marker></" + std::string(c) + "marker>";
    }
    if (!format.symbol.empty() && !patchOrInsertValChild(marker, "c:symbol", "symbol", format.symbol)) return false;
    if (format.size > 0 && !patchOrInsertValChild(marker, "c:size", "size", std::to_string(format.size))) return false;
    auto spPrNodes = drawingTags(marker, "c:spPr", "spPr");
    std::string spPr = spPrNodes.empty() ? std::string{} : spPrNodes.front();
    if ((format.line.present || format.fill.present) && spPr.empty()) {
        if (!ensureChartSpPr(marker, spPr)) return false;
    }
    if (!spPr.empty()) {
        auto patched = spPr;
        if (format.line.present && !patchChartLineFormatInSpPr(patched, format.line)) return false;
        if (format.fill.present && !patchChartFillFormatInSpPr(patched, format.fill)) return false;
        const auto position = marker.find(spPr);
        if (position == std::string::npos) return false;
        marker.replace(position, spPr.size(), patched);
    }
    if (!markers.empty()) {
        const auto position = series.find(markers.front());
        if (position == std::string::npos) return false;
        series.replace(position, markers.front().size(), marker);
    } else {
        std::size_t insertion = std::string::npos;
        for (const auto& pair : std::array<std::pair<const char*, const char*>, 3>{{{"c:dLbls", "dLbls"}, {"c:trendline", "trendline"}, {"c:errBars", "errBars"}}}) {
            const auto nodes = drawingTags(series, pair.first, pair.second);
            if (!nodes.empty()) { insertion = series.find(nodes.front()); if (insertion != std::string::npos) break; }
        }
        if (insertion == std::string::npos) insertion = series.rfind("</");
        if (insertion == std::string::npos) return false;
        series.insert(insertion, marker);
    }
    const auto position = chartXmlText.find(original);
    if (position == std::string::npos) return false;
    chartXmlText.replace(position, original.size(), series);
    return true;
}

bool patchNestedLineFormat(std::string& owner, const xlpp::ChartLineFormat& format) {
    auto spPrNodes = drawingTags(owner, "c:spPr", "spPr");
    std::string spPr = spPrNodes.empty() ? std::string{} : spPrNodes.front();
    if (spPr.empty() && !ensureChartSpPr(owner, spPr)) return false;
    auto patched = spPr;
    if (!patchChartLineFormatInSpPr(patched, format)) return false;
    const auto position = owner.find(spPr);
    if (position == std::string::npos) return false;
    owner.replace(position, spPr.size(), patched);
    return true;
}

std::string trendlineTypeValue(xlpp::ChartSeries::TrendlineType type) {
    using T = xlpp::ChartSeries::TrendlineType;
    switch (type) {
    case T::Linear: return "linear";
    case T::Exponential: return "exp";
    case T::Logarithmic: return "log";
    case T::Polynomial: return "poly";
    case T::Power: return "power";
    case T::MovingAverage: return "movingAvg";
    }
    return "linear";
}

std::string errorBarDirectionValue(xlpp::ChartSeries::ErrorBarDirection direction) {
    return direction == xlpp::ChartSeries::ErrorBarDirection::X ? "x" : "y";
}

std::string errorBarTypeValue(xlpp::ChartSeries::ErrorBarType type) {
    using T = xlpp::ChartSeries::ErrorBarType;
    switch (type) {
    case T::Both: return "both";
    case T::Plus: return "plus";
    case T::Minus: return "minus";
    }
    return "both";
}

std::string errorValueTypeValue(xlpp::ChartSeries::ErrorValueType type) {
    using T = xlpp::ChartSeries::ErrorValueType;
    switch (type) {
    case T::FixedValue: return "fixedVal";
    case T::Percentage: return "percentage";
    case T::StandardDeviation: return "stdDev";
    case T::StandardError: return "stdErr";
    case T::Custom: return "cust";
    }
    return "fixedVal";
}

std::string formatChartDouble(double value) {
    std::ostringstream stream;
    stream << std::setprecision(15) << value;
    return stream.str();
}

bool patchTrendlineNode(std::string& trendlineXml, const xlpp::ChartSeries::Trendline& trendline) {
    if (!patchOrInsertValChild(trendlineXml, "c:trendlineType", "trendlineType", trendlineTypeValue(trendline.type))) return false;
    if (trendline.type == xlpp::ChartSeries::TrendlineType::Polynomial) {
        if (!patchOrInsertValChild(trendlineXml, "c:order", "order", std::to_string(trendline.order))) return false;
    } else removeDrawingChild(trendlineXml, "c:order", "order");
    if (trendline.type == xlpp::ChartSeries::TrendlineType::MovingAverage) {
        if (!patchOrInsertValChild(trendlineXml, "c:period", "period", std::to_string(trendline.period))) return false;
    } else removeDrawingChild(trendlineXml, "c:period", "period");
    if (trendline.forward > 0.0) {
        if (!patchOrInsertValChild(trendlineXml, "c:forward", "forward", formatChartDouble(trendline.forward))) return false;
    } else removeDrawingChild(trendlineXml, "c:forward", "forward");
    if (trendline.backward > 0.0) {
        if (!patchOrInsertValChild(trendlineXml, "c:backward", "backward", formatChartDouble(trendline.backward))) return false;
    } else removeDrawingChild(trendlineXml, "c:backward", "backward");
    if (!patchOrInsertValChild(trendlineXml, "c:dispRSqr", "dispRSqr", trendline.displayRSquared ? "1" : "0")) return false;
    if (!patchOrInsertValChild(trendlineXml, "c:dispEq", "dispEq", trendline.displayEquation ? "1" : "0")) return false;
    if (trendline.lineFormat.present && !patchNestedLineFormat(trendlineXml, trendline.lineFormat)) return false;
    return true;
}

std::string makeTrendlineXml(const xlpp::ChartSeries::Trendline& trendline, bool prefixed) {
    const auto c = prefixed ? "c:" : "";
    std::string xml = "<" + std::string(c) + "trendline><" + std::string(c) + "trendlineType val=\"" + trendlineTypeValue(trendline.type) + "\"/>";
    if (trendline.type == xlpp::ChartSeries::TrendlineType::Polynomial)
        xml += "<" + std::string(c) + "order val=\"" + std::to_string(trendline.order) + "\"/>";
    if (trendline.type == xlpp::ChartSeries::TrendlineType::MovingAverage)
        xml += "<" + std::string(c) + "period val=\"" + std::to_string(trendline.period) + "\"/>";
    if (trendline.forward > 0.0) xml += "<" + std::string(c) + "forward val=\"" + formatChartDouble(trendline.forward) + "\"/>";
    if (trendline.backward > 0.0) xml += "<" + std::string(c) + "backward val=\"" + formatChartDouble(trendline.backward) + "\"/>";
    xml += "<" + std::string(c) + "dispRSqr val=\"" + (trendline.displayRSquared ? "1" : "0") + "\"/>";
    xml += "<" + std::string(c) + "dispEq val=\"" + (trendline.displayEquation ? "1" : "0") + "\"/></" + std::string(c) + "trendline>";
    if (trendline.lineFormat.present && !patchNestedLineFormat(xml, trendline.lineFormat)) return {};
    return xml;
}

bool patchImportedChartSeriesTrendline(std::string& chartXmlText, std::size_t seriesIndex,
                                       std::size_t trendlineIndex,
                                       const xlpp::ChartSeries::Trendline* trendline,
                                       bool add) {
    const auto seriesNodes = drawingTags(chartXmlText, "c:ser", "ser");
    if (seriesIndex >= seriesNodes.size()) return false;
    const auto original = seriesNodes[seriesIndex];
    auto series = original;
    const auto nodes = drawingTags(series, "c:trendline", "trendline");
    if (add) {
        const bool prefixed = series.find("<c:ser") != std::string::npos;
        const auto generated = makeTrendlineXml(*trendline, prefixed);
        if (generated.empty()) return false;
        std::size_t insertion = std::string::npos;
        const auto errBars = drawingTags(series, "c:errBars", "errBars");
        if (!errBars.empty()) insertion = series.find(errBars.front());
        if (insertion == std::string::npos) {
            for (const auto& pair : std::array<std::pair<const char*, const char*>, 4>{{{"c:cat", "cat"}, {"c:xVal", "xVal"}, {"c:val", "val"}, {"c:yVal", "yVal"}}}) {
                const auto refs = drawingTags(series, pair.first, pair.second);
                if (!refs.empty()) { insertion = series.find(refs.front()); if (insertion != std::string::npos) break; }
            }
        }
        if (insertion == std::string::npos) insertion = series.rfind("</");
        if (insertion == std::string::npos) return false;
        series.insert(insertion, generated);
    } else {
        if (trendlineIndex >= nodes.size()) return false;
        const auto position = series.find(nodes[trendlineIndex]);
        if (position == std::string::npos) return false;
        if (!trendline) series.erase(position, nodes[trendlineIndex].size());
        else {
            auto patched = nodes[trendlineIndex];
            if (!patchTrendlineNode(patched, *trendline)) return false;
            series.replace(position, nodes[trendlineIndex].size(), patched);
        }
    }
    const auto position = chartXmlText.find(original);
    if (position == std::string::npos) return false;
    chartXmlText.replace(position, original.size(), series);
    return true;
}

bool patchImportedChartSeriesTrendlineLineFormat(std::string& chartXmlText, std::size_t seriesIndex,
                                                   std::size_t trendlineIndex, const xlpp::ChartLineFormat& format) {
    const auto seriesNodes = drawingTags(chartXmlText, "c:ser", "ser");
    if (seriesIndex >= seriesNodes.size()) return false;
    const auto originalSeries = seriesNodes[seriesIndex];
    auto series = originalSeries;
    const auto trendlines = drawingTags(series, "c:trendline", "trendline");
    if (trendlineIndex >= trendlines.size()) return false;
    auto trendline = trendlines[trendlineIndex];
    if (!patchNestedLineFormat(trendline, format)) return false;
    const auto trendPosition = series.find(trendlines[trendlineIndex]);
    if (trendPosition == std::string::npos) return false;
    series.replace(trendPosition, trendlines[trendlineIndex].size(), trendline);
    const auto position = chartXmlText.find(originalSeries);
    if (position == std::string::npos) return false;
    chartXmlText.replace(position, originalSeries.size(), series);
    return true;
}

bool patchImportedChartSeriesErrorBarsLineFormat(std::string& chartXmlText, std::size_t seriesIndex,
                                                  xlpp::ChartSeries::ErrorBarDirection direction,
                                                  const xlpp::ChartLineFormat& format) {
    const auto seriesNodes = drawingTags(chartXmlText, "c:ser", "ser");
    if (seriesIndex >= seriesNodes.size()) return false;
    const auto originalSeries = seriesNodes[seriesIndex];
    auto series = originalSeries;
    const auto nodes = drawingTags(series, "c:errBars", "errBars");
    const auto found = std::find_if(nodes.begin(), nodes.end(), [&](const auto& node) {
        const auto dirs = drawingTags(node, "c:errDir", "errDir");
        const auto actual = dirs.empty() ? std::string("y") : xlpp::internal::attribute(dirs.front(), "val");
        return actual == errorBarDirectionValue(direction);
    });
    if (found == nodes.end()) return false;
    auto errorBars = *found;
    if (!patchNestedLineFormat(errorBars, format)) return false;
    const auto barsPosition = series.find(*found);
    if (barsPosition == std::string::npos) return false;
    series.replace(barsPosition, found->size(), errorBars);
    const auto position = chartXmlText.find(originalSeries);
    if (position == std::string::npos) return false;
    chartXmlText.replace(position, originalSeries.size(), series);
    return true;
}

bool patchErrorBarReference(std::string& errorBarsXml, const char* prefixed, const char* local, const std::string& reference) {
    if (reference.empty()) return false;
    const bool prefixedContainer = errorBarsXml.find("<c:errBars") != std::string::npos;
    const auto c = prefixedContainer ? "c:" : "";
    const auto generated = "<" + std::string(c) + local + "><" + std::string(c) + "numRef><" + std::string(c) + "f>" +
                           xmlEscape(reference) + "</" + std::string(c) + "f></" + std::string(c) + "numRef></" + std::string(c) + local + ">";
    const auto nodes = drawingTags(errorBarsXml, prefixed, local);
    if (!nodes.empty()) {
        auto node = nodes.front();
        const auto refs = drawingTags(node, "c:numRef", "numRef");
        if (!refs.empty()) {
            auto numRef = refs.front();
            const auto formulas = drawingTags(numRef, "c:f", "f");
            const auto generatedFormula = "<" + std::string(c) + "f>" + xmlEscape(reference) + "</" + std::string(c) + "f>";
            if (!formulas.empty()) {
                const auto pos = numRef.find(formulas.front());
                if (pos == std::string::npos) return false;
                numRef.replace(pos, formulas.front().size(), generatedFormula);
            } else {
                const auto close = numRef.rfind("</");
                if (close == std::string::npos) return false;
                numRef.insert(close, generatedFormula);
            }
            const auto refPos = node.find(refs.front());
            if (refPos == std::string::npos) return false;
            node.replace(refPos, refs.front().size(), numRef);
            const auto position = errorBarsXml.find(nodes.front());
            if (position == std::string::npos) return false;
            errorBarsXml.replace(position, nodes.front().size(), node);
            return true;
        }
        const auto position = errorBarsXml.find(nodes.front());
        if (position == std::string::npos) return false;
        errorBarsXml.replace(position, nodes.front().size(), generated);
        return true;
    }
    const auto spPr = drawingTags(errorBarsXml, "c:spPr", "spPr");
    const auto insertion = !spPr.empty() ? errorBarsXml.find(spPr.front()) : errorBarsXml.rfind("</");
    if (insertion == std::string::npos) return false;
    errorBarsXml.insert(insertion, generated);
    return true;
}

bool patchErrorBarsNode(std::string& errorBarsXml, const xlpp::ChartSeries::ErrorBars& errorBars) {
    if (!patchOrInsertValChild(errorBarsXml, "c:errDir", "errDir", errorBarDirectionValue(errorBars.direction))) return false;
    if (!patchOrInsertValChild(errorBarsXml, "c:errBarType", "errBarType", errorBarTypeValue(errorBars.barType))) return false;
    if (!patchOrInsertValChild(errorBarsXml, "c:errValType", "errValType", errorValueTypeValue(errorBars.valueType))) return false;
    if (!patchOrInsertValChild(errorBarsXml, "c:noEndCap", "noEndCap", errorBars.noEndCap ? "1" : "0")) return false;
    if (errorBars.valueType != xlpp::ChartSeries::ErrorValueType::Custom) {
        removeDrawingChild(errorBarsXml, "c:plus", "plus");
        removeDrawingChild(errorBarsXml, "c:minus", "minus");
        if (!patchOrInsertValChild(errorBarsXml, "c:val", "val", formatChartDouble(errorBars.value))) return false;
    } else {
        removeDrawingChild(errorBarsXml, "c:val", "val");
        if (!patchErrorBarReference(errorBarsXml, "c:minus", "minus", errorBars.minusReference)) return false;
        if (!patchErrorBarReference(errorBarsXml, "c:plus", "plus", errorBars.plusReference)) return false;
    }
    if (errorBars.lineFormat.present && !patchNestedLineFormat(errorBarsXml, errorBars.lineFormat)) return false;
    return true;
}

std::string makeErrorBarsXml(const xlpp::ChartSeries::ErrorBars& errorBars, bool prefixed) {
    const auto c = prefixed ? "c:" : "";
    std::string xml = "<" + std::string(c) + "errBars>";
    xml += "<" + std::string(c) + "errDir val=\"" + errorBarDirectionValue(errorBars.direction) + "\"/>";
    xml += "<" + std::string(c) + "errBarType val=\"" + errorBarTypeValue(errorBars.barType) + "\"/>";
    xml += "<" + std::string(c) + "errValType val=\"" + errorValueTypeValue(errorBars.valueType) + "\"/>";
    xml += "<" + std::string(c) + "noEndCap val=\"" + (errorBars.noEndCap ? "1" : "0") + "\"/>";
    if (errorBars.valueType != xlpp::ChartSeries::ErrorValueType::Custom)
        xml += "<" + std::string(c) + "val val=\"" + formatChartDouble(errorBars.value) + "\"/>";
    else {
        xml += "<" + std::string(c) + "minus><" + std::string(c) + "numRef><" + std::string(c) + "f>" +
               xmlEscape(errorBars.minusReference) + "</" + std::string(c) + "f></" + std::string(c) + "numRef></" + std::string(c) + "minus>";
        xml += "<" + std::string(c) + "plus><" + std::string(c) + "numRef><" + std::string(c) + "f>" +
               xmlEscape(errorBars.plusReference) + "</" + std::string(c) + "f></" + std::string(c) + "numRef></" + std::string(c) + "plus>";
    }
    xml += "</" + std::string(c) + "errBars>";
    if (errorBars.lineFormat.present && !patchNestedLineFormat(xml, errorBars.lineFormat)) return {};
    return xml;
}

bool patchImportedChartSeriesErrorBars(std::string& chartXmlText, std::size_t seriesIndex,
                                       xlpp::ChartSeries::ErrorBarDirection direction,
                                       const xlpp::ChartSeries::ErrorBars* errorBars) {
    const auto seriesNodes = drawingTags(chartXmlText, "c:ser", "ser");
    if (seriesIndex >= seriesNodes.size()) return false;
    const auto original = seriesNodes[seriesIndex];
    auto series = original;
    auto matchesDirection = [&](const std::string& node) {
        const auto dirs = drawingTags(node, "c:errDir", "errDir");
        if (dirs.empty()) return direction == xlpp::ChartSeries::ErrorBarDirection::Y;
        return xlpp::internal::attribute(dirs.front(), "val") == errorBarDirectionValue(direction);
    };
    const auto nodes = drawingTags(series, "c:errBars", "errBars");
    const auto found = std::find_if(nodes.begin(), nodes.end(), matchesDirection);
    if (!errorBars) {
        if (found == nodes.end()) return false;
        const auto position = series.find(*found);
        if (position == std::string::npos) return false;
        series.erase(position, found->size());
    } else if (found != nodes.end()) {
        auto patched = *found;
        if (!patchErrorBarsNode(patched, *errorBars)) return false;
        const auto position = series.find(*found);
        if (position == std::string::npos) return false;
        series.replace(position, found->size(), patched);
    } else {
        const bool prefixed = series.find("<c:ser") != std::string::npos;
        const auto generated = makeErrorBarsXml(*errorBars, prefixed);
        if (generated.empty()) return false;
        std::size_t insertion = std::string::npos;
        for (const auto& pair : std::array<std::pair<const char*, const char*>, 4>{{{"c:cat", "cat"}, {"c:xVal", "xVal"}, {"c:val", "val"}, {"c:yVal", "yVal"}}}) {
            const auto refs = drawingTags(series, pair.first, pair.second);
            if (!refs.empty()) { insertion = series.find(refs.front()); if (insertion != std::string::npos) break; }
        }
        if (insertion == std::string::npos) insertion = series.rfind("</");
        if (insertion == std::string::npos) return false;
        series.insert(insertion, generated);
    }
    const auto position = chartXmlText.find(original);
    if (position == std::string::npos) return false;
    chartXmlText.replace(position, original.size(), series);
    return true;
}

std::vector<std::pair<std::size_t, std::string>> patchablePlotNodesInOrder(const std::string& chartXmlText) {
    const auto plotAreas = drawingTags(chartXmlText, "c:plotArea", "plotArea");
    if (plotAreas.empty()) return {};
    const auto& plotArea = plotAreas.front();
    std::vector<std::pair<std::size_t, std::string>> result;
    const auto collect = [&](const char* prefixed, const char* local) {
        std::size_t cursor = 0;
        for (const auto& node : drawingTags(plotArea, prefixed, local)) {
            const auto position = plotArea.find(node, cursor);
            if (position == std::string::npos) continue;
            result.emplace_back(position, node);
            cursor = position + node.size();
        }
    };
    collect("c:barChart", "barChart"); collect("c:lineChart", "lineChart"); collect("c:pieChart", "pieChart");
    collect("c:scatterChart", "scatterChart"); collect("c:doughnutChart", "doughnutChart"); collect("c:radarChart", "radarChart");
    collect("c:areaChart", "areaChart"); collect("c:bubbleChart", "bubbleChart"); collect("c:stockChart", "stockChart");
    collect("c:ofPieChart", "ofPieChart");
    collect("c:bar3DChart", "bar3DChart"); collect("c:line3DChart", "line3DChart"); collect("c:area3DChart", "area3DChart");
    collect("c:pie3DChart", "pie3DChart"); collect("c:surfaceChart", "surfaceChart"); collect("c:surface3DChart", "surface3DChart");
    std::sort(result.begin(), result.end(), [](const auto& a, const auto& b) { return a.first < b.first; });
    return result;
}

bool patchImportedChartSeriesDataLabels(std::string& chartXmlText, std::size_t seriesIndex, const xlpp::Chart::DataLabels& labels) {
    const auto seriesNodes = drawingTags(chartXmlText, "c:ser", "ser");
    if (seriesIndex >= seriesNodes.size()) return false;
    const auto originalSeries = seriesNodes[seriesIndex];
    auto series = originalSeries;
    const auto existingLabels = drawingTags(series, "c:dLbls", "dLbls");
    std::string dLbls;
    if (!existingLabels.empty()) dLbls = existingLabels.front();
    else {
        const bool prefixed = series.find("<c:ser") != std::string::npos;
        const auto c = prefixed ? "c:" : "";
        dLbls = "<" + std::string(c) + "dLbls></" + std::string(c) + "dLbls>";
    }
    if (!labels.position.empty() && !patchOrInsertAggregateDataLabelVal(dLbls, "c:dLblPos", "dLblPos", labels.position)) return false;
    const auto patchFlag = [&](const char* prefixed, const char* local, bool value) {
        const auto mask = dataLabelsAggregateMask(dLbls);
        const bool exists = !drawingTags(mask, prefixed, local).empty();
        return patchOrInsertAggregateDataLabelVal(dLbls, prefixed, local, value ? "1" : "0", value || exists);
    };
    if (!patchFlag("c:showLegendKey", "showLegendKey", labels.showLegendKey) ||
        !patchFlag("c:showVal", "showVal", labels.showValue) ||
        !patchFlag("c:showCatName", "showCatName", labels.showCategoryName) ||
        !patchFlag("c:showSerName", "showSerName", labels.showSeriesName) ||
        !patchFlag("c:showPercent", "showPercent", labels.showPercent) ||
        !patchFlag("c:showBubbleSize", "showBubbleSize", labels.showBubbleSize) ||
        !patchFlag("c:showLeaderLines", "showLeaderLines", labels.showLeaderLines)) return false;
    if (!labels.separator.empty() &&
        !patchOrInsertAggregateDataLabelText(dLbls, "c:separator", "separator", labels.separator)) return false;
    if (!existingLabels.empty()) {
        const auto position = series.find(existingLabels.front());
        if (position == std::string::npos) return false;
        series.replace(position, existingLabels.front().size(), dLbls);
    } else {
        std::size_t insertion = std::string::npos;
        for (const auto& pair : std::array<std::pair<const char*, const char*>, 6>{{
                 {"c:trendline", "trendline"}, {"c:errBars", "errBars"}, {"c:cat", "cat"},
                 {"c:xVal", "xVal"}, {"c:val", "val"}, {"c:yVal", "yVal"}}}) {
            const auto nodes = drawingTags(series, pair.first, pair.second);
            if (!nodes.empty()) { insertion = series.find(nodes.front()); if (insertion != std::string::npos) break; }
        }
        if (insertion == std::string::npos) insertion = series.rfind("</");
        if (insertion == std::string::npos) return false;
        series.insert(insertion, dLbls);
    }
    const auto position = chartXmlText.find(originalSeries);
    if (position == std::string::npos) return false;
    chartXmlText.replace(position, originalSeries.size(), series);
    return true;
}

bool patchImportedChartPlotDataLabels(std::string& chartXmlText, std::size_t plotIndex, const xlpp::Chart::DataLabels& labels) {
    const auto plotNodes = patchablePlotNodesInOrder(chartXmlText);
    if (plotIndex >= plotNodes.size()) return false;
    const auto originalPlot = plotNodes[plotIndex].second;
    auto plot = originalPlot;
    auto existingLabels = drawingTags(plot, "c:dLbls", "dLbls");
    const auto seriesNodes = drawingTags(plot, "c:ser", "ser");
    existingLabels.erase(std::remove_if(existingLabels.begin(), existingLabels.end(), [&](const std::string& node) {
        return std::any_of(seriesNodes.begin(), seriesNodes.end(), [&](const std::string& series) {
            return series.find(node) != std::string::npos;
        });
    }), existingLabels.end());
    std::string dLbls;
    if (!existingLabels.empty()) dLbls = existingLabels.front();
    else {
        const bool prefixed = plot.find("<c:") != std::string::npos;
        const auto c = prefixed ? "c:" : "";
        dLbls = "<" + std::string(c) + "dLbls></" + std::string(c) + "dLbls>";
    }
    if (!labels.position.empty() && !patchOrInsertAggregateDataLabelVal(dLbls, "c:dLblPos", "dLblPos", labels.position)) return false;
    const auto patchFlag = [&](const char* prefixed, const char* local, bool value) {
        const auto mask = dataLabelsAggregateMask(dLbls);
        const bool exists = !drawingTags(mask, prefixed, local).empty();
        return patchOrInsertAggregateDataLabelVal(dLbls, prefixed, local, value ? "1" : "0", value || exists);
    };
    if (!patchFlag("c:showLegendKey", "showLegendKey", labels.showLegendKey) ||
        !patchFlag("c:showVal", "showVal", labels.showValue) ||
        !patchFlag("c:showCatName", "showCatName", labels.showCategoryName) ||
        !patchFlag("c:showSerName", "showSerName", labels.showSeriesName) ||
        !patchFlag("c:showPercent", "showPercent", labels.showPercent) ||
        !patchFlag("c:showBubbleSize", "showBubbleSize", labels.showBubbleSize) ||
        !patchFlag("c:showLeaderLines", "showLeaderLines", labels.showLeaderLines)) return false;
    if (!labels.separator.empty() &&
        !patchOrInsertAggregateDataLabelText(dLbls, "c:separator", "separator", labels.separator)) return false;
    if (!existingLabels.empty()) {
        const auto position = plot.find(existingLabels.front());
        if (position == std::string::npos) return false;
        plot.replace(position, existingLabels.front().size(), dLbls);
    } else {
        const auto axisIds = drawingTags(plot, "c:axId", "axId");
        std::size_t insertion = !axisIds.empty() ? plot.find(axisIds.front()) : plot.rfind("</");
        if (insertion == std::string::npos) return false;
        plot.insert(insertion, dLbls);
    }
    const auto position = chartXmlText.find(originalPlot);
    if (position == std::string::npos) return false;
    chartXmlText.replace(position, originalPlot.size(), plot);
    return true;
}

bool patchShapeOwnerFormat(std::string& owner, const xlpp::ChartLineFormat* line, const xlpp::ChartFillFormat* fill) {
    auto spPrNodes = drawingTags(owner, "c:spPr", "spPr");
    std::string spPr = spPrNodes.empty() ? std::string{} : spPrNodes.front();
    if (spPr.empty() && !ensureChartSpPr(owner, spPr)) return false;
    auto patched = spPr;
    if (line && line->present && !patchChartLineFormatInSpPr(patched, *line)) return false;
    if (fill && fill->present && !patchChartFillFormatInSpPr(patched, *fill)) return false;
    const auto position = owner.find(spPr);
    if (position == std::string::npos) return false;
    owner.replace(position, spPr.size(), patched);
    return true;
}

std::string generatedChartWallXml(const char* localName, const xlpp::ChartWallFormat& format, bool prefixed) {
    if (!format.present) return {};
    const auto c = prefixed ? "c:" : "";
    std::string wall = "<" + std::string(c) + localName + ">";
    if (format.hasThickness)
        wall += "<" + std::string(c) + "thickness val=\"" + std::to_string(format.thickness) + "\"/>";
    wall += "</" + std::string(c) + localName + ">";
    if ((format.line.present || format.fill.present) &&
        !patchShapeOwnerFormat(wall, format.line.present ? &format.line : nullptr, format.fill.present ? &format.fill : nullptr))
        return {};
    return wall;
}

bool patchImportedChartView3D(std::string& chartXmlText, const xlpp::ChartView3D& view) {
    if (!view.present) return false;
    const auto chartNodes = drawingTags(chartXmlText, "c:chart", "chart");
    if (chartNodes.empty()) return false;
    const auto originalChart = chartNodes.front();
    auto chart = originalChart;
    const bool prefixed = chart.find("<c:") != std::string::npos;
    const auto generated = chartView3DXml(view, prefixed);
    const auto existing = drawingTags(chart, "c:view3D", "view3D");
    if (!existing.empty()) {
        const auto pos = chart.find(existing.front());
        if (pos == std::string::npos) return false;
        chart.replace(pos, existing.front().size(), generated);
    } else {
        std::size_t insertion = std::string::npos;
        for (const auto& pair : std::array<std::pair<const char*, const char*>, 4>{{
                 {"c:floor","floor"},{"c:sideWall","sideWall"},{"c:backWall","backWall"},{"c:plotArea","plotArea"}}}) {
            const auto nodes = drawingTags(chart, pair.first, pair.second);
            if (!nodes.empty()) { insertion = chart.find(nodes.front()); if (insertion != std::string::npos) break; }
        }
        if (insertion == std::string::npos) insertion = chart.rfind("</");
        if (insertion == std::string::npos) return false;
        chart.insert(insertion, generated);
    }
    const auto pos = chartXmlText.find(originalChart);
    if (pos == std::string::npos) return false;
    chartXmlText.replace(pos, originalChart.size(), chart);
    return true;
}

bool patchImportedChartWallFormat(std::string& chartXmlText, const char* prefixedName, const char* localName,
                                  const xlpp::ChartWallFormat& format) {
    if (!format.present) return false;
    const auto chartNodes = drawingTags(chartXmlText, "c:chart", "chart");
    if (chartNodes.empty()) return false;
    const auto originalChart = chartNodes.front();
    auto chart = originalChart;
    const bool prefixed = chart.find("<c:") != std::string::npos;
    const auto c = prefixed ? "c:" : "";
    const auto existing = drawingTags(chart, prefixedName, localName);
    std::string wall = existing.empty()
        ? "<" + std::string(c) + localName + "></" + std::string(c) + localName + ">"
        : existing.front();
    if (format.hasThickness && !patchOrInsertValChild(wall, "c:thickness", "thickness", std::to_string(format.thickness), true)) return false;
    if ((format.line.present || format.fill.present) &&
        !patchShapeOwnerFormat(wall, format.line.present ? &format.line : nullptr, format.fill.present ? &format.fill : nullptr)) return false;
    if (!existing.empty()) {
        const auto pos = chart.find(existing.front());
        if (pos == std::string::npos) return false;
        chart.replace(pos, existing.front().size(), wall);
    } else {
        std::vector<std::pair<const char*, const char*>> later;
        const std::string name(localName);
        if (name == "floor") later = {{"c:sideWall","sideWall"},{"c:backWall","backWall"},{"c:plotArea","plotArea"}};
        else if (name == "sideWall") later = {{"c:backWall","backWall"},{"c:plotArea","plotArea"}};
        else later = {{"c:plotArea","plotArea"}};
        std::size_t insertion = std::string::npos;
        for (const auto& pair : later) {
            const auto nodes = drawingTags(chart, pair.first, pair.second);
            if (!nodes.empty()) { insertion = chart.find(nodes.front()); if (insertion != std::string::npos) break; }
        }
        if (insertion == std::string::npos) insertion = chart.rfind("</");
        if (insertion == std::string::npos) return false;
        chart.insert(insertion, wall);
    }
    const auto pos = chartXmlText.find(originalChart);
    if (pos == std::string::npos) return false;
    chartXmlText.replace(pos, originalChart.size(), chart);
    return true;
}

std::string chartTextStyleTxPrXml(const xlpp::ChartTextStyle& style, bool prefixed, bool strict) {
    if (!style.present) return {};
    const auto c = prefixed ? "c:" : "";
    std::string rPr = "<a:defRPr";
    if (style.bold) rPr += " b=\"1\"";
    if (style.italic) rPr += " i=\"1\"";
    if (style.fontSizePoints > 0.0) {
        const auto size = static_cast<long long>(std::llround(style.fontSizePoints * 100.0));
        rPr += " sz=\"" + std::to_string(size) + "\"";
    }
    rPr += ">";
    if (style.color.present()) rPr += chartSolidFillXml(style.color, false);
    if (!style.typeface.empty()) rPr += "<a:latin typeface=\"" + xmlEscape(style.typeface) + "\"/>";
    rPr += "</a:defRPr>";
    const auto drawingNs = strict ? "http://purl.oclc.org/ooxml/drawingml/main"
                                  : "http://schemas.openxmlformats.org/drawingml/2006/main";
    return "<" + std::string(c) + "txPr xmlns:a=\"" + std::string(drawingNs) + "\"><a:bodyPr/><a:lstStyle/><a:p><a:pPr>" + rPr +
           "</a:pPr><a:endParaRPr lang=\"en-US\"/></a:p></" + std::string(c) + "txPr>";
}

std::string generatedDataLabelsXml(const xlpp::ChartDataLabels& labels) {
    if (!labels.present && !labels.hasLeaderLines) return {};
    std::string xml = "<c:dLbls>";
    const auto flag=[&](const char* name,bool value){ xml += "<c:" + std::string(name) + " val=\"" + (value ? "1" : "0") + "\"/>"; };
    flag("showLegendKey", labels.showLegendKey); flag("showVal", labels.showValue); flag("showCatName", labels.showCategoryName);
    flag("showSerName", labels.showSeriesName); flag("showPercent", labels.showPercent); flag("showBubbleSize", labels.showBubbleSize);
    if (!labels.position.empty()) xml += "<c:dLblPos val=\"" + xmlEscape(labels.position) + "\"/>";
    if (!labels.separator.empty()) xml += "<c:separator>" + xmlEscape(labels.separator) + "</c:separator>";
    if (labels.hasLeaderLines) {
        flag("showLeaderLines", true);
        std::string leader = "<c:leaderLines></c:leaderLines>";
        if (labels.leaderLineFormat.present && !patchNestedLineFormat(leader, labels.leaderLineFormat)) return {};
        xml += leader;
    } else if (labels.showLeaderLines) flag("showLeaderLines", true);
    xml += "</c:dLbls>";
    return xml;
}

std::string generatedPlotAuxiliaryXml(const xlpp::Chart::Plot& plot, bool /*strict*/) {
    std::string xml;
    xml += generatedDataLabelsXml(plot.dataLabels);
    const auto lineObject=[&](const char* name, bool present, const xlpp::ChartLineFormat& format) {
        if (!present) return std::string{};
        std::string object = "<c:" + std::string(name) + "></c:" + std::string(name) + ">";
        if (format.present && !patchNestedLineFormat(object, format)) return std::string{};
        return object;
    };
    xml += lineObject("dropLines", plot.hasDropLines, plot.dropLinesFormat);
    xml += lineObject("hiLowLines", plot.hasHighLowLines, plot.highLowLinesFormat);
    if (plot.upDownBars.present) {
        std::string object = "<c:upDownBars><c:gapWidth val=\"" + std::to_string(plot.upDownBars.gapWidth) + "\"/>";
        const auto bar=[&](const char* name, const xlpp::ChartLineFormat& line, const xlpp::ChartFillFormat& fill) {
            std::string owner = "<c:" + std::string(name) + "></c:" + std::string(name) + ">";
            if ((line.present || fill.present) && !patchShapeOwnerFormat(owner, line.present ? &line : nullptr, fill.present ? &fill : nullptr)) return std::string{};
            return owner;
        };
        const auto up = bar("upBars", plot.upDownBars.upLine, plot.upDownBars.upFill);
        const auto down = bar("downBars", plot.upDownBars.downLine, plot.upDownBars.downFill);
        if (up.empty() || down.empty()) return {};
        object += up + down + "</c:upDownBars>";
        xml += object;
    }
    return xml;
}

std::string generatedDataTableXml(const xlpp::ChartDataTable& table, bool strict) {
    if (!table.present) return {};
    std::string xml = "<c:dTable>";
    const auto flag=[&](const char* name,bool value){ xml += "<c:" + std::string(name) + " val=\"" + (value ? "1" : "0") + "\"/>"; };
    flag("showHorzBorder", table.showHorizontalBorder); flag("showVertBorder", table.showVerticalBorder);
    flag("showOutline", table.showOutline); flag("showKeys", table.showLegendKeys);
    xml += "</c:dTable>";
    if ((table.line.present || table.fill.present) &&
        !patchShapeOwnerFormat(xml, table.line.present ? &table.line : nullptr, table.fill.present ? &table.fill : nullptr)) return {};
    if (table.textStyle.present) {
        const auto txPr = chartTextStyleTxPrXml(table.textStyle, true, strict);
        const auto insertion = xml.rfind("</c:dTable>");
        if (insertion == std::string::npos) return {};
        xml.insert(insertion, txPr);
    }
    return xml;
}

bool patchImportedChartDataTable(std::string& chartXmlText, const xlpp::ChartDataTable* table) {
    const auto plotAreas = drawingTags(chartXmlText, "c:plotArea", "plotArea");
    if (plotAreas.empty()) return false;
    const auto originalPlotArea = plotAreas.front();
    auto plotArea = originalPlotArea;
    auto tables = drawingTags(plotArea, "c:dTable", "dTable");
    if (!table) {
        if (!tables.empty()) {
            const auto position = plotArea.find(tables.front());
            if (position == std::string::npos) return false;
            plotArea.erase(position, tables.front().size());
        }
    } else {
        const bool prefixed = plotArea.find("<c:") != std::string::npos;
        const auto c = prefixed ? "c:" : "";
        std::string generated = "<" + std::string(c) + "dTable>";
        const auto addFlag = [&](const char* name, bool value) {
            generated += "<" + std::string(c) + name + " val=\"" + (value ? "1" : "0") + "\"/>";
        };
        addFlag("showHorzBorder", table->showHorizontalBorder);
        addFlag("showVertBorder", table->showVerticalBorder);
        addFlag("showOutline", table->showOutline);
        addFlag("showKeys", table->showLegendKeys);
        generated += "</" + std::string(c) + "dTable>";
        if ((table->line.present || table->fill.present) &&
            !patchShapeOwnerFormat(generated, table->line.present ? &table->line : nullptr,
                                   table->fill.present ? &table->fill : nullptr)) return false;
        if (table->textStyle.present) {
            const bool strict = chartXmlText.find("http://purl.oclc.org/ooxml/drawingml/chart") != std::string::npos;
            const auto txPr = chartTextStyleTxPrXml(table->textStyle, prefixed, strict);
            const auto closing = generated.rfind(prefixed ? "</c:dTable>" : "</dTable>");
            if (closing == std::string::npos) return false;
            generated.insert(closing, txPr);
        }
        if (!tables.empty()) {
            const auto position = plotArea.find(tables.front());
            if (position == std::string::npos) return false;
            plotArea.replace(position, tables.front().size(), generated);
        } else {
            const auto directSpPr = plotAreaDirectSpPr(plotArea);
            std::size_t insertion = !directSpPr.empty() ? plotArea.find(directSpPr) : std::string::npos;
            if (insertion == std::string::npos) {
                const auto ext = drawingTags(plotArea, "c:extLst", "extLst");
                insertion = !ext.empty() ? plotArea.find(ext.front()) : plotArea.rfind("</");
            }
            if (insertion == std::string::npos) return false;
            plotArea.insert(insertion, generated);
        }
    }
    const auto position = chartXmlText.find(originalPlotArea);
    if (position == std::string::npos) return false;
    chartXmlText.replace(position, originalPlotArea.size(), plotArea);
    return true;
}

std::size_t plotAuxiliaryInsertion(const std::string& plot, const char* local) {
    std::vector<std::pair<const char*, const char*>> later;
    const std::string tag(local);
    if (tag == "dropLines") later = {{"c:hiLowLines","hiLowLines"},{"c:upDownBars","upDownBars"},{"c:marker","marker"},{"c:smooth","smooth"},{"c:axId","axId"}};
    else if (tag == "hiLowLines") later = {{"c:upDownBars","upDownBars"},{"c:marker","marker"},{"c:smooth","smooth"},{"c:axId","axId"}};
    else later = {{"c:marker","marker"},{"c:smooth","smooth"},{"c:axId","axId"}};
    for (const auto& item : later) {
        const auto nodes = drawingTags(plot, item.first, item.second);
        if (!nodes.empty()) {
            const auto position = plot.find(nodes.front());
            if (position != std::string::npos) return position;
        }
    }
    return plot.rfind("</");
}

bool patchImportedChartPlotLineObject(std::string& chartXmlText, std::size_t plotIndex,
                                      const char* prefixed, const char* local,
                                      const xlpp::ChartLineFormat* format) {
    const auto plotNodes = patchablePlotNodesInOrder(chartXmlText);
    if (plotIndex >= plotNodes.size()) return false;
    const auto originalPlot = plotNodes[plotIndex].second;
    auto plot = originalPlot;
    auto nodes = drawingTags(plot, prefixed, local);
    if (!format) {
        if (!nodes.empty()) {
            const auto position = plot.find(nodes.front());
            if (position == std::string::npos) return false;
            plot.erase(position, nodes.front().size());
        }
    } else {
        std::string object;
        if (!nodes.empty()) object = nodes.front();
        else {
            const bool hasPrefix = plot.find("<c:") != std::string::npos;
            const auto c = hasPrefix ? "c:" : "";
            object = "<" + std::string(c) + local + "></" + std::string(c) + local + ">";
        }
        if (format->present && !patchNestedLineFormat(object, *format)) return false;
        if (!nodes.empty()) {
            const auto position = plot.find(nodes.front());
            if (position == std::string::npos) return false;
            plot.replace(position, nodes.front().size(), object);
        } else {
            const auto insertion = plotAuxiliaryInsertion(plot, local);
            if (insertion == std::string::npos) return false;
            plot.insert(insertion, object);
        }
    }
    const auto position = chartXmlText.find(originalPlot);
    if (position == std::string::npos) return false;
    chartXmlText.replace(position, originalPlot.size(), plot);
    return true;
}

bool patchImportedChartPlotUpDownBars(std::string& chartXmlText, std::size_t plotIndex,
                                      const xlpp::ChartUpDownBars* bars) {
    const auto plotNodes = patchablePlotNodesInOrder(chartXmlText);
    if (plotIndex >= plotNodes.size()) return false;
    const auto originalPlot = plotNodes[plotIndex].second;
    auto plot = originalPlot;
    auto nodes = drawingTags(plot, "c:upDownBars", "upDownBars");
    if (!bars) {
        if (!nodes.empty()) {
            const auto position = plot.find(nodes.front());
            if (position == std::string::npos) return false;
            plot.erase(position, nodes.front().size());
        }
    } else {
        const bool prefixed = plot.find("<c:") != std::string::npos;
        const auto c = prefixed ? "c:" : "";
        std::string object = "<" + std::string(c) + "upDownBars><" + std::string(c) + "gapWidth val=\"" +
                             std::to_string(bars->gapWidth) + "\"/>";
        auto makeBar = [&](const char* name, const xlpp::ChartLineFormat& line, const xlpp::ChartFillFormat& fill) {
            std::string bar = "<" + std::string(c) + name + "></" + std::string(c) + name + ">";
            if ((line.present || fill.present) && !patchShapeOwnerFormat(bar, line.present ? &line : nullptr, fill.present ? &fill : nullptr)) return std::string{};
            return bar;
        };
        const auto up = makeBar("upBars", bars->upLine, bars->upFill);
        const auto down = makeBar("downBars", bars->downLine, bars->downFill);
        if (up.empty() || down.empty()) return false;
        object += up + down + "</" + std::string(c) + "upDownBars>";
        if (!nodes.empty()) {
            const auto position = plot.find(nodes.front());
            if (position == std::string::npos) return false;
            plot.replace(position, nodes.front().size(), object);
        } else {
            const auto insertion = plotAuxiliaryInsertion(plot, "upDownBars");
            if (insertion == std::string::npos) return false;
            plot.insert(insertion, object);
        }
    }
    const auto position = chartXmlText.find(originalPlot);
    if (position == std::string::npos) return false;
    chartXmlText.replace(position, originalPlot.size(), plot);
    return true;
}

bool patchLeaderLinesInDataLabels(std::string& dLbls, const xlpp::ChartLineFormat* format, bool remove) {
    auto leaderLines = drawingTags(dLbls, "c:leaderLines", "leaderLines");
    if (remove) {
        if (!leaderLines.empty()) {
            const auto position = dLbls.find(leaderLines.front());
            if (position == std::string::npos) return false;
            dLbls.erase(position, leaderLines.front().size());
        }
        const auto mask = dataLabelsAggregateMask(dLbls);
        const bool exists = !drawingTags(mask, "c:showLeaderLines", "showLeaderLines").empty();
        if (exists && !patchOrInsertAggregateDataLabelVal(dLbls, "c:showLeaderLines", "showLeaderLines", "0", true)) return false;
        return true;
    }
    if (!format || !format->present) return false;
    if (!patchOrInsertAggregateDataLabelVal(dLbls, "c:showLeaderLines", "showLeaderLines", "1", true)) return false;
    std::string leader;
    if (!leaderLines.empty()) leader = leaderLines.front();
    else {
        const bool prefixed = dLbls.find("<c:") != std::string::npos;
        const auto c = prefixed ? "c:" : "";
        leader = "<" + std::string(c) + "leaderLines></" + std::string(c) + "leaderLines>";
    }
    if (!patchNestedLineFormat(leader, *format)) return false;
    if (!leaderLines.empty()) {
        const auto position = dLbls.find(leaderLines.front());
        if (position == std::string::npos) return false;
        dLbls.replace(position, leaderLines.front().size(), leader);
    } else {
        const auto ext = drawingTags(dLbls, "c:extLst", "extLst");
        const auto insertion = !ext.empty() ? dLbls.find(ext.front()) : dLbls.rfind("</");
        if (insertion == std::string::npos) return false;
        dLbls.insert(insertion, leader);
    }
    return true;
}

bool patchImportedChartPlotSimpleValue(std::string& chartXmlText, std::size_t plotIndex,
                                       const char* prefixed, const char* local, const std::string& value,
                                       bool beforeSeries) {
    const auto plotNodes=patchablePlotNodesInOrder(chartXmlText); if(plotIndex>=plotNodes.size()) return false;
    const auto original=plotNodes[plotIndex].second; auto plot=original; auto nodes=drawingTags(plot,prefixed,local);
    const bool hasPrefix=plot.find("<c:")!=std::string::npos; const auto c=hasPrefix?"c:":"";
    const std::string generated="<"+std::string(c)+local+" val=\""+xmlEscape(value)+"\"/>";
    if(!nodes.empty()){ const auto pos=plot.find(nodes.front()); if(pos==std::string::npos) return false; plot.replace(pos,nodes.front().size(),generated); }
    else {
        std::size_t insertion=std::string::npos;
        if(beforeSeries){ const auto ser=drawingTags(plot,"c:ser","ser"); if(!ser.empty()) insertion=plot.find(ser.front()); }
        else { const auto ext=drawingTags(plot,"c:extLst","extLst"); insertion=!ext.empty()?plot.find(ext.front()):plot.rfind("</"); }
        if(insertion==std::string::npos) return false; plot.insert(insertion,generated);
    }
    const auto pos=chartXmlText.find(original); if(pos==std::string::npos) return false; chartXmlText.replace(pos,original.size(),plot); return true;
}

bool patchImportedChartProjectedPie(std::string& chartXmlText, std::size_t plotIndex, const xlpp::ChartProjectedPieOptions& options) {
    const auto plotNodes=patchablePlotNodesInOrder(chartXmlText); if(plotIndex>=plotNodes.size()) return false;
    const auto original=plotNodes[plotIndex].second; auto plot=original;
    if(plot.find("ofPieChart")==std::string::npos) return false;
    const bool hasPrefix=plot.find("<c:")!=std::string::npos; const auto c=hasPrefix?"c:":"";
    for(const auto& tag:std::array<std::pair<const char*,const char*>,7>{{
        {"c:ofPieType","ofPieType"},{"c:gapWidth","gapWidth"},{"c:splitType","splitType"},{"c:splitPos","splitPos"},
        {"c:custSplit","custSplit"},{"c:secondPieSize","secondPieSize"},{"c:serLines","serLines"}}}){
        const auto nodes=drawingTags(plot,tag.first,tag.second);
        if(!nodes.empty()){ const auto pos=plot.find(nodes.front()); if(pos==std::string::npos) return false; plot.erase(pos,nodes.front().size()); }
    }
    std::ostringstream block;
    block<<"<"<<c<<"ofPieType val=\""<<xmlEscape(options.ofPieType)<<"\"/>";
    block<<"<"<<c<<"gapWidth val=\""<<options.gapWidth<<"\"/><"<<c<<"splitType val=\""<<xmlEscape(options.splitType)<<"\"/>";
    if(options.hasSplitPosition) block<<"<"<<c<<"splitPos val=\""<<options.splitPosition<<"\"/>";
    if(!options.customSplitPoints.empty()){ block<<"<"<<c<<"custSplit>"; for(const auto point:options.customSplitPoints) block<<"<"<<c<<"secondPiePt val=\""<<point<<"\"/>"; block<<"</"<<c<<"custSplit>"; }
    block<<"<"<<c<<"secondPieSize val=\""<<options.secondPlotSize<<"\"/>";
    if(options.hasSeriesLines){ std::string lines="<"+std::string(c)+"serLines></"+std::string(c)+"serLines>"; if(options.seriesLinesFormat.present&&!patchNestedLineFormat(lines,options.seriesLinesFormat)) return false; block<<lines; }
    const auto ext=drawingTags(plot,"c:extLst","extLst"); const auto insertion=!ext.empty()?plot.find(ext.front()):plot.rfind("</"); if(insertion==std::string::npos) return false; plot.insert(insertion,block.str());
    const auto pos=chartXmlText.find(original); if(pos==std::string::npos) return false; chartXmlText.replace(pos,original.size(),plot); return true;
}

bool patchImportedChartLeaderLines(std::string& chartXmlText, bool plotLevel, std::size_t ownerIndex,
                                   const xlpp::ChartLineFormat* format, bool remove) {
    if (plotLevel) {
        const auto plotNodes = patchablePlotNodesInOrder(chartXmlText);
        if (ownerIndex >= plotNodes.size()) return false;
        const auto originalPlot = plotNodes[ownerIndex].second;
        auto plot = originalPlot;
        auto labels = drawingTags(plot, "c:dLbls", "dLbls");
        const auto seriesNodes = drawingTags(plot, "c:ser", "ser");
        labels.erase(std::remove_if(labels.begin(), labels.end(), [&](const std::string& node) {
            return std::any_of(seriesNodes.begin(), seriesNodes.end(), [&](const std::string& series){ return series.find(node) != std::string::npos; });
        }), labels.end());
        if (labels.empty()) {
            if (remove) return true;
            const bool prefixed = plot.find("<c:") != std::string::npos;
            const auto c = prefixed ? "c:" : "";
            std::string dLbls = "<" + std::string(c) + "dLbls></" + std::string(c) + "dLbls>";
            if (!patchLeaderLinesInDataLabels(dLbls, format, false)) return false;
            const auto axisIds = drawingTags(plot, "c:axId", "axId");
            const auto insertion = !axisIds.empty() ? plot.find(axisIds.front()) : plot.rfind("</");
            if (insertion == std::string::npos) return false;
            plot.insert(insertion, dLbls);
        } else {
            auto dLbls = labels.front();
            if (!patchLeaderLinesInDataLabels(dLbls, format, remove)) return false;
            const auto position = plot.find(labels.front()); if (position == std::string::npos) return false;
            plot.replace(position, labels.front().size(), dLbls);
        }
        const auto position = chartXmlText.find(originalPlot); if (position == std::string::npos) return false;
        chartXmlText.replace(position, originalPlot.size(), plot); return true;
    }
    const auto seriesNodes = drawingTags(chartXmlText, "c:ser", "ser");
    if (ownerIndex >= seriesNodes.size()) return false;
    const auto originalSeries = seriesNodes[ownerIndex];
    auto series = originalSeries;
    auto labels = drawingTags(series, "c:dLbls", "dLbls");
    if (labels.empty()) {
        if (remove) return true;
        const bool prefixed = series.find("<c:ser") != std::string::npos;
        const auto c = prefixed ? "c:" : "";
        std::string dLbls = "<" + std::string(c) + "dLbls></" + std::string(c) + "dLbls>";
        if (!patchLeaderLinesInDataLabels(dLbls, format, false)) return false;
        std::size_t insertion = series.rfind("</");
        for (const auto& item : std::array<std::pair<const char*, const char*>, 6>{{{"c:trendline","trendline"},{"c:errBars","errBars"},{"c:cat","cat"},{"c:xVal","xVal"},{"c:val","val"},{"c:yVal","yVal"}}}) {
            const auto nodes = drawingTags(series, item.first, item.second);
            if (!nodes.empty()) { const auto pos=series.find(nodes.front()); if(pos!=std::string::npos) { insertion=pos; break; } }
        }
        if (insertion == std::string::npos) return false;
        series.insert(insertion, dLbls);
    } else {
        auto dLbls = labels.front();
        if (!patchLeaderLinesInDataLabels(dLbls, format, remove)) return false;
        const auto position = series.find(labels.front()); if(position==std::string::npos) return false;
        series.replace(position, labels.front().size(), dLbls);
    }
    const auto position = chartXmlText.find(originalSeries); if(position==std::string::npos) return false;
    chartXmlText.replace(position, originalSeries.size(), series); return true;
}

std::string dataLabelPointXml(const xlpp::ChartDataLabelPoint& label, bool prefixed) {
    const auto c = prefixed ? "c:" : "";
    std::string xml = "<" + std::string(c) + "dLbl><" + std::string(c) + "idx val=\"" + std::to_string(label.index) + "\"/>";
    if (label.deleted) xml += "<" + std::string(c) + "delete val=\"1\"/>";
    if (!label.position.empty()) xml += "<" + std::string(c) + "dLblPos val=\"" + xmlEscape(label.position) + "\"/>";
    if (label.showLegendKey) xml += "<" + std::string(c) + "showLegendKey val=\"1\"/>";
    if (label.showValue) xml += "<" + std::string(c) + "showVal val=\"1\"/>";
    if (label.showCategoryName) xml += "<" + std::string(c) + "showCatName val=\"1\"/>";
    if (label.showSeriesName) xml += "<" + std::string(c) + "showSerName val=\"1\"/>";
    if (label.showPercent) xml += "<" + std::string(c) + "showPercent val=\"1\"/>";
    if (label.showBubbleSize) xml += "<" + std::string(c) + "showBubbleSize val=\"1\"/>";
    if (label.showLeaderLines) xml += "<" + std::string(c) + "showLeaderLines val=\"1\"/>";
    if (!label.separator.empty()) xml += "<" + std::string(c) + "separator>" + xmlEscape(label.separator) + "</" + std::string(c) + "separator>";
    xml += "</" + std::string(c) + "dLbl>";
    return xml;
}

bool patchDataLabelPointNode(std::string& dLbls, const xlpp::ChartDataLabelPoint& label, bool remove) {
    auto points = drawingTags(dLbls, "c:dLbl", "dLbl");
    const auto found = std::find_if(points.begin(), points.end(), [&](const auto& point) {
        const auto idx = drawingTags(point, "c:idx", "idx");
        if (idx.empty()) return false;
        try { return std::stoull(xlpp::internal::attribute(idx.front(), "val")) == label.index; } catch (...) { return false; }
    });
    if (remove) {
        if (found == points.end()) return false;
        const auto position = dLbls.find(*found);
        if (position == std::string::npos) return false;
        dLbls.erase(position, found->size());
        return true;
    }
    if (found == points.end()) {
        const bool prefixed = dLbls.find("<c:dLbls") != std::string::npos;
        const auto generated = dataLabelPointXml(label, prefixed);
        std::size_t insertion = dLbls.rfind("</");
        for (const auto& pair : std::array<std::pair<const char*, const char*>, 8>{{
                 {"c:delete", "delete"}, {"c:dLblPos", "dLblPos"}, {"c:showLegendKey", "showLegendKey"}, {"c:showVal", "showVal"},
                 {"c:showCatName", "showCatName"}, {"c:showSerName", "showSerName"}, {"c:showPercent", "showPercent"}, {"c:showBubbleSize", "showBubbleSize"}}}) {
            const auto nodes = drawingTags(dLbls, pair.first, pair.second);
            for (const auto& node : nodes) {
                if (std::any_of(points.begin(), points.end(), [&](const auto& point) { return point.find(node) != std::string::npos; })) continue;
                const auto pos = dLbls.find(node);
                if (pos != std::string::npos) { insertion = std::min(insertion, pos); break; }
            }
        }
        if (insertion == std::string::npos) return false;
        dLbls.insert(insertion, generated);
        return true;
    }

    auto patched = *found;
    const auto patchFlag = [&](const char* prefixed, const char* local, bool value) {
        const bool exists = !drawingTags(patched, prefixed, local).empty();
        return patchOrInsertValChild(patched, prefixed, local, value ? "1" : "0", value || exists);
    };
    if (!patchFlag("c:delete", "delete", label.deleted) ||
        !patchFlag("c:showLegendKey", "showLegendKey", label.showLegendKey) ||
        !patchFlag("c:showVal", "showVal", label.showValue) ||
        !patchFlag("c:showCatName", "showCatName", label.showCategoryName) ||
        !patchFlag("c:showSerName", "showSerName", label.showSeriesName) ||
        !patchFlag("c:showPercent", "showPercent", label.showPercent) ||
        !patchFlag("c:showBubbleSize", "showBubbleSize", label.showBubbleSize) ||
        !patchFlag("c:showLeaderLines", "showLeaderLines", label.showLeaderLines)) return false;
    if (!label.position.empty() && !patchOrInsertValChild(patched, "c:dLblPos", "dLblPos", label.position)) return false;
    if (!label.separator.empty()) {
        const auto separators = drawingTags(patched, "c:separator", "separator");
        const bool prefixed = patched.find("<c:dLbl") != std::string::npos;
        const auto c = prefixed ? "c:" : "";
        const auto generated = "<" + std::string(c) + "separator>" + xmlEscape(label.separator) + "</" + std::string(c) + "separator>";
        if (!separators.empty()) {
            const auto pos = patched.find(separators.front()); if (pos == std::string::npos) return false;
            patched.replace(pos, separators.front().size(), generated);
        } else {
            const auto close = patched.rfind("</"); if (close == std::string::npos) return false;
            patched.insert(close, generated);
        }
    }
    const auto position = dLbls.find(*found);
    if (position == std::string::npos) return false;
    dLbls.replace(position, found->size(), patched);
    return true;
}

std::string plotDirectDataLabels(const std::string& plot) {
    auto labels = drawingTags(plot, "c:dLbls", "dLbls");
    const auto series = drawingTags(plot, "c:ser", "ser");
    labels.erase(std::remove_if(labels.begin(), labels.end(), [&](const auto& node) {
        return std::any_of(series.begin(), series.end(), [&](const auto& seriesNode) { return seriesNode.find(node) != std::string::npos; });
    }), labels.end());
    return labels.empty() ? std::string{} : labels.front();
}

bool patchImportedChartSeriesDataLabelPoint(std::string& chartXmlText, std::size_t seriesIndex,
                                            const xlpp::ChartDataLabelPoint& label, bool remove) {
    const auto seriesNodes = drawingTags(chartXmlText, "c:ser", "ser");
    if (seriesIndex >= seriesNodes.size()) return false;
    const auto original = seriesNodes[seriesIndex];
    auto series = original;
    const auto labelNodes = drawingTags(series, "c:dLbls", "dLbls");
    std::string dLbls = labelNodes.empty() ? std::string{} : labelNodes.front();
    if (dLbls.empty()) {
        if (remove) return false;
        const bool prefixed = series.find("<c:ser") != std::string::npos;
        const auto c = prefixed ? "c:" : "";
        dLbls = "<" + std::string(c) + "dLbls></" + std::string(c) + "dLbls>";
    }
    if (!patchDataLabelPointNode(dLbls, label, remove)) return false;
    if (!labelNodes.empty()) {
        const auto position = series.find(labelNodes.front());
        if (position == std::string::npos) return false;
        series.replace(position, labelNodes.front().size(), dLbls);
    } else {
        std::size_t insertion = series.rfind("</");
        for (const auto& pair : std::array<std::pair<const char*, const char*>, 6>{{
                 {"c:trendline", "trendline"}, {"c:errBars", "errBars"}, {"c:cat", "cat"}, {"c:xVal", "xVal"}, {"c:val", "val"}, {"c:yVal", "yVal"}}}) {
            const auto nodes = drawingTags(series, pair.first, pair.second);
            if (!nodes.empty()) { const auto pos = series.find(nodes.front()); if (pos != std::string::npos) insertion = std::min(insertion, pos); }
        }
        if (insertion == std::string::npos) return false;
        series.insert(insertion, dLbls);
    }
    const auto position = chartXmlText.find(original);
    if (position == std::string::npos) return false;
    chartXmlText.replace(position, original.size(), series);
    return true;
}

bool patchImportedChartSeriesDataLabelPointRichText(std::string& chartXmlText, std::size_t seriesIndex,
                                                    std::size_t pointIndex, const xlpp::ChartRichText& richText) {
    if (!richText.present || richText.runs.empty()) return false;
    const auto seriesNodes = drawingTags(chartXmlText, "c:ser", "ser");
    if (seriesIndex >= seriesNodes.size()) return false;
    const auto originalSeries = seriesNodes[seriesIndex];
    auto series = originalSeries;
    auto labelNodes = drawingTags(series, "c:dLbls", "dLbls");
    std::string dLbls;
    if (labelNodes.empty()) {
        const bool prefixed = series.find("<c:ser") != std::string::npos;
        const auto c = prefixed ? "c:" : "";
        dLbls = "<" + std::string(c) + "dLbls></" + std::string(c) + "dLbls>";
        xlpp::ChartDataLabelPoint point;
        point.index = pointIndex;
        if (!patchDataLabelPointNode(dLbls, point, false)) return false;
    } else dLbls = labelNodes.front();

    auto points = drawingTags(dLbls, "c:dLbl", "dLbl");
    auto found = std::find_if(points.begin(), points.end(), [&](const auto& point) {
        const auto indices = drawingTags(point, "c:idx", "idx");
        if (indices.empty()) return false;
        try { return std::stoull(xlpp::internal::attribute(indices.front(), "val")) == pointIndex; } catch (...) { return false; }
    });
    if (found == points.end()) {
        xlpp::ChartDataLabelPoint point;
        point.index = pointIndex;
        if (!patchDataLabelPointNode(dLbls, point, false)) return false;
        points = drawingTags(dLbls, "c:dLbl", "dLbl");
        found = std::find_if(points.begin(), points.end(), [&](const auto& pointXml) {
            const auto indices = drawingTags(pointXml, "c:idx", "idx");
            if (indices.empty()) return false;
            try { return std::stoull(xlpp::internal::attribute(indices.front(), "val")) == pointIndex; } catch (...) { return false; }
        });
        if (found == points.end()) return false;
    }
    auto pointXml = *found;
    const bool prefixed = pointXml.find("<c:dLbl") != std::string::npos;
    const bool strict = chartXmlText.find("http://purl.oclc.org/ooxml/drawingml/chart") != std::string::npos;
    const auto txXml = chartRichTextTxXml(richText, prefixed, strict);
    const auto txNodes = drawingTags(pointXml, "c:tx", "tx");
    if (!txNodes.empty()) {
        const auto pos = pointXml.find(txNodes.front());
        if (pos == std::string::npos) return false;
        pointXml.replace(pos, txNodes.front().size(), txXml);
    } else {
        std::size_t insertion = pointXml.rfind("</");
        const auto positions = drawingTags(pointXml, "c:dLblPos", "dLblPos");
        if (!positions.empty()) insertion = pointXml.find(positions.front());
        if (insertion == std::string::npos) return false;
        pointXml.insert(insertion, txXml);
    }
    const auto pointPos = dLbls.find(*found);
    if (pointPos == std::string::npos) return false;
    dLbls.replace(pointPos, found->size(), pointXml);

    if (!labelNodes.empty()) {
        const auto pos = series.find(labelNodes.front());
        if (pos == std::string::npos) return false;
        series.replace(pos, labelNodes.front().size(), dLbls);
    } else {
        std::size_t insertion = series.rfind("</");
        for (const auto& pair : std::array<std::pair<const char*, const char*>, 6>{{
                 {"c:trendline", "trendline"}, {"c:errBars", "errBars"}, {"c:cat", "cat"}, {"c:xVal", "xVal"}, {"c:val", "val"}, {"c:yVal", "yVal"}}}) {
            const auto nodes = drawingTags(series, pair.first, pair.second);
            if (!nodes.empty()) {
                const auto pos = series.find(nodes.front());
                if (pos != std::string::npos) insertion = std::min(insertion, pos);
            }
        }
        if (insertion == std::string::npos) return false;
        series.insert(insertion, dLbls);
    }
    const auto seriesPos = chartXmlText.find(originalSeries);
    if (seriesPos == std::string::npos) return false;
    chartXmlText.replace(seriesPos, originalSeries.size(), series);
    return true;
}

bool patchMarkerFormatInOwner(std::string& owner, const xlpp::ChartMarkerFormat& format) {
    auto markers = drawingTags(owner, "c:marker", "marker");
    std::string marker;
    if (!markers.empty()) marker = markers.front();
    else {
        const bool prefixed = owner.find("<c:") != std::string::npos;
        const auto c = prefixed ? "c:" : "";
        marker = "<" + std::string(c) + "marker></" + std::string(c) + "marker>";
    }
    if (!format.symbol.empty() && !patchOrInsertValChild(marker, "c:symbol", "symbol", format.symbol)) return false;
    if (format.size > 0 && !patchOrInsertValChild(marker, "c:size", "size", std::to_string(format.size))) return false;
    auto spPrNodes = drawingTags(marker, "c:spPr", "spPr");
    std::string spPr = spPrNodes.empty() ? std::string{} : spPrNodes.front();
    if ((format.line.present || format.fill.present) && spPr.empty()) {
        if (!ensureChartSpPr(marker, spPr)) return false;
    }
    if (!spPr.empty()) {
        auto patched = spPr;
        if (format.line.present && !patchChartLineFormatInSpPr(patched, format.line)) return false;
        if (format.fill.present && !patchChartFillFormatInSpPr(patched, format.fill)) return false;
        const auto position = marker.find(spPr);
        if (position == std::string::npos) return false;
        marker.replace(position, spPr.size(), patched);
    }
    if (!markers.empty()) {
        const auto position = owner.find(markers.front());
        if (position == std::string::npos) return false;
        owner.replace(position, markers.front().size(), marker);
    } else {
        const auto close = owner.rfind("</");
        if (close == std::string::npos) return false;
        owner.insert(close, marker);
    }
    return true;
}

bool patchImportedChartSeriesDataPointFormat(std::string& chartXmlText, std::size_t seriesIndex,
                                             const xlpp::ChartDataPointFormat& format, bool remove) {
    const auto seriesNodes = drawingTags(chartXmlText, "c:ser", "ser");
    if (seriesIndex >= seriesNodes.size()) return false;
    const auto originalSeries = seriesNodes[seriesIndex];
    auto series = originalSeries;
    auto points = drawingTags(series, "c:dPt", "dPt");
    auto found = std::find_if(points.begin(), points.end(), [&](const auto& point) {
        const auto indices = drawingTags(point, "c:idx", "idx");
        if (indices.empty()) return false;
        try { return std::stoull(xlpp::internal::attribute(indices.front(), "val")) == format.index; } catch (...) { return false; }
    });
    if (remove) {
        if (found == points.end()) return false;
        const auto position = series.find(*found);
        if (position == std::string::npos) return false;
        series.erase(position, found->size());
    } else {
        std::string pointXml;
        if (found != points.end()) pointXml = *found;
        else {
            const bool prefixed = series.find("<c:ser") != std::string::npos;
            const auto c = prefixed ? "c:" : "";
            pointXml = "<" + std::string(c) + "dPt><" + std::string(c) + "idx val=\"" +
                       std::to_string(format.index) + "\"/></" + std::string(c) + "dPt>";
        }
        auto spPrNodes = drawingTags(pointXml, "c:spPr", "spPr");
        std::string spPr = spPrNodes.empty() ? std::string{} : spPrNodes.front();
        if ((format.line.present || format.fill.present) && spPr.empty()) {
            std::string before;
            const auto markerNodes = drawingTags(pointXml, "c:marker", "marker");
            if (!markerNodes.empty()) before = markerNodes.front();
            if (!ensureChartSpPr(pointXml, spPr, before)) return false;
        }
        if (!spPr.empty()) {
            auto patched = spPr;
            if (format.line.present && !patchChartLineFormatInSpPr(patched, format.line)) return false;
            if (format.fill.present && !patchChartFillFormatInSpPr(patched, format.fill)) return false;
            const auto spPos = pointXml.find(spPr);
            if (spPos == std::string::npos) return false;
            pointXml.replace(spPos, spPr.size(), patched);
        }
        if (format.marker.present && !patchMarkerFormatInOwner(pointXml, format.marker)) return false;

        if (found != points.end()) {
            const auto position = series.find(*found);
            if (position == std::string::npos) return false;
            series.replace(position, found->size(), pointXml);
        } else {
            std::size_t insertion = series.rfind("</");
            for (const auto& pair : std::array<std::pair<const char*, const char*>, 7>{{
                     {"c:dLbls", "dLbls"}, {"c:trendline", "trendline"}, {"c:errBars", "errBars"},
                     {"c:cat", "cat"}, {"c:xVal", "xVal"}, {"c:val", "val"}, {"c:yVal", "yVal"}}}) {
                const auto nodes = drawingTags(series, pair.first, pair.second);
                if (!nodes.empty()) {
                    const auto pos = series.find(nodes.front());
                    if (pos != std::string::npos) insertion = std::min(insertion, pos);
                }
            }
            if (insertion == std::string::npos) return false;
            series.insert(insertion, pointXml);
        }
    }
    const auto seriesPosition = chartXmlText.find(originalSeries);
    if (seriesPosition == std::string::npos) return false;
    chartXmlText.replace(seriesPosition, originalSeries.size(), series);
    return true;
}

bool patchImportedChartPlotDataLabelPoint(std::string& chartXmlText, std::size_t plotIndex,
                                          const xlpp::ChartDataLabelPoint& label, bool remove) {
    const auto plots = patchablePlotNodesInOrder(chartXmlText);
    if (plotIndex >= plots.size()) return false;
    const auto original = plots[plotIndex].second;
    auto plot = original;
    const auto existing = plotDirectDataLabels(plot);
    std::string dLbls = existing;
    if (dLbls.empty()) {
        if (remove) return false;
        const bool prefixed = plot.find("<c:") != std::string::npos;
        const auto c = prefixed ? "c:" : "";
        dLbls = "<" + std::string(c) + "dLbls></" + std::string(c) + "dLbls>";
    }
    if (!patchDataLabelPointNode(dLbls, label, remove)) return false;
    if (!existing.empty()) {
        const auto pos = plot.find(existing); if (pos == std::string::npos) return false;
        plot.replace(pos, existing.size(), dLbls);
    } else {
        const auto axisIds = drawingTags(plot, "c:axId", "axId");
        const auto insertion = !axisIds.empty() ? plot.find(axisIds.front()) : plot.rfind("</");
        if (insertion == std::string::npos) return false;
        plot.insert(insertion, dLbls);
    }
    const auto position = chartXmlText.find(original);
    if (position == std::string::npos) return false;
    chartXmlText.replace(position, original.size(), plot);
    return true;
}

bool patchImportedChartSeriesTitle(std::string& chartXmlText,
                                   std::size_t seriesIndex,
                                   const std::string& title) {
    const auto seriesNodes = drawingTags(chartXmlText, "c:ser", "ser");
    if (seriesIndex >= seriesNodes.size()) return false;
    const auto original = seriesNodes[seriesIndex];
    auto series = original;
    const auto txNodes = drawingTags(series, "c:tx", "tx");
    if (title.empty()) {
        if (!txNodes.empty()) {
            const auto position = series.find(txNodes.front());
            if (position == std::string::npos) return false;
            series.erase(position, txNodes.front().size());
        }
    } else {
        const bool prefixed = series.find("<c:ser") != std::string::npos;
        const auto c = prefixed ? "c:" : "";
        const std::string generated = "<" + std::string(c) + "tx><" + std::string(c) + "v>" +
            xmlEscape(title) + "</" + std::string(c) + "v></" + std::string(c) + "tx>";
        if (!txNodes.empty()) {
            const auto position = series.find(txNodes.front());
            if (position == std::string::npos) return false;
            series.replace(position, txNodes.front().size(), generated);
        } else {
            const auto idxNodes = drawingTags(series, "c:idx", "idx");
            const auto orderNodes = drawingTags(series, "c:order", "order");
            std::size_t insertion = std::string::npos;
            if (!orderNodes.empty()) {
                const auto pos = series.find(orderNodes.front());
                if (pos != std::string::npos) insertion = pos + orderNodes.front().size();
            } else if (!idxNodes.empty()) {
                const auto pos = series.find(idxNodes.front());
                if (pos != std::string::npos) insertion = pos + idxNodes.front().size();
            }
            if (insertion == std::string::npos) {
                const auto open = series.find('>');
                if (open == std::string::npos) return false;
                insertion = open + 1;
            }
            series.insert(insertion, generated);
        }
    }
    const auto position = chartXmlText.find(original);
    if (position == std::string::npos) return false;
    chartXmlText.replace(position, original.size(), series);
    return true;
}

bool patchCacheInReferenceOwner(std::string& owner, const xlpp::ChartSeriesCache* cache) {
    auto refs = drawingTags(owner, "c:numRef", "numRef");
    bool numericRef = true;
    if (refs.empty()) { refs = drawingTags(owner, "c:strRef", "strRef"); numericRef = false; }
    if (refs.empty()) return cache == nullptr;
    const auto originalRef = refs.front();
    auto ref = originalRef;
    if (!cache) { eraseChartCacheBlocks(ref); }
    else {
        if (!cache->present || cache->numeric != numericRef) return false;
        const auto existingNum = drawingTags(ref, "c:numCache", "numCache");
        const auto existingStr = drawingTags(ref, "c:strCache", "strCache");
        const auto existing = !existingNum.empty() ? existingNum.front() : (!existingStr.empty() ? existingStr.front() : std::string{});
        const bool prefixed = ref.find("<c:") != std::string::npos;
        const auto generated = chartSeriesCacheXml(*cache, prefixed);
        if (!existing.empty()) {
            const auto pos = ref.find(existing); if (pos == std::string::npos) return false; ref.replace(pos, existing.size(), generated);
        } else {
            const auto close = ref.rfind("</"); if (close == std::string::npos) return false; ref.insert(close, generated);
        }
    }
    const auto pos = owner.find(originalRef);
    if (pos == std::string::npos) return false;
    owner.replace(pos, originalRef.size(), ref);
    return true;
}

bool patchImportedChartSeriesCache(std::string& chartXmlText, std::size_t seriesIndex, int kind, const xlpp::ChartSeriesCache* cache) {
    const auto seriesNodes = drawingTags(chartXmlText, "c:ser", "ser");
    if (seriesIndex >= seriesNodes.size()) return false;
    const auto original = seriesNodes[seriesIndex];
    auto series = original;
    if (kind == 3) eraseChartCacheBlocks(series);
    else {
        const char* prefixed = kind == 0 ? "c:cat" : (kind == 1 ? "c:val" : "c:tx");
        const char* local = kind == 0 ? "cat" : (kind == 1 ? "val" : "tx");
        auto owners = drawingTags(series, prefixed, local);
        if (owners.empty() && kind == 0) owners = drawingTags(series, "c:xVal", "xVal");
        if (owners.empty() && kind == 1) owners = drawingTags(series, "c:yVal", "yVal");
        if (owners.empty()) return false;
        auto owner = owners.front();
        if (!patchCacheInReferenceOwner(owner, cache)) return false;
        const auto ownerPos = series.find(owners.front()); if (ownerPos == std::string::npos) return false;
        series.replace(ownerPos, owners.front().size(), owner);
    }
    const auto pos = chartXmlText.find(original); if (pos == std::string::npos) return false;
    chartXmlText.replace(pos, original.size(), series);
    return true;
}

bool patchImportedChartSeriesReferences(std::string& chartXmlText,
                                        std::size_t seriesIndex,
                                        const std::string& categoriesReference,
                                        const std::string& valuesReference) {
    const auto seriesNodes = drawingTags(chartXmlText, "c:ser", "ser");
    if (seriesIndex >= seriesNodes.size()) return false;
    const auto original = seriesNodes[seriesIndex];
    auto patched = original;
    bool categoriesOk = patchSeriesReferenceContainer(patched, "c:cat", "cat", categoriesReference);
    if (!categoriesOk) categoriesOk = patchSeriesReferenceContainer(patched, "c:xVal", "xVal", categoriesReference);
    bool valuesOk = patchSeriesReferenceContainer(patched, "c:val", "val", valuesReference);
    if (!valuesOk) valuesOk = patchSeriesReferenceContainer(patched, "c:yVal", "yVal", valuesReference);
    if (!categoriesOk || !valuesOk) return false;
    const auto position = chartXmlText.find(original);
    if (position == std::string::npos) return false;
    chartXmlText.replace(position, original.size(), patched);
    return true;
}

void suppressExclusivePartClosure(const std::string& rootPart,
                                  const std::vector<xlpp::PreservedRelationship>& allRelationships,
                                  std::set<std::string>& suppressedPreservedParts) {
    std::unordered_set<std::string> closure;
    std::vector<std::string> stack{rootPart};
    while (!stack.empty()) {
        auto part = std::move(stack.back());
        stack.pop_back();
        if (part.empty() || !closure.insert(part).second) continue;
        for (const auto& relationship : allRelationships) {
            if (relationship.sourcePart != part || relationship.targetMode == "External") continue;
            const auto target = resolvePackagePart(part, relationship.target);
            if (!target.empty() && !closure.count(target)) stack.push_back(target);
        }
    }

    std::unordered_set<std::string> protectedParts;
    for (const auto& candidate : closure) {
        if (candidate == rootPart) continue;
        const bool externallyReferenced = std::any_of(allRelationships.begin(), allRelationships.end(), [&](const auto& relationship) {
            if (relationship.targetMode == "External" || closure.count(relationship.sourcePart)) return false;
            return resolvePackagePart(relationship.sourcePart, relationship.target) == candidate;
        });
        if (externallyReferenced) protectedParts.insert(candidate);
    }
    std::vector<std::string> protectStack(protectedParts.begin(), protectedParts.end());
    while (!protectStack.empty()) {
        auto part = std::move(protectStack.back());
        protectStack.pop_back();
        for (const auto& relationship : allRelationships) {
            if (relationship.sourcePart != part || relationship.targetMode == "External") continue;
            const auto target = resolvePackagePart(part, relationship.target);
            if (closure.count(target) && protectedParts.insert(target).second) protectStack.push_back(target);
        }
    }
    for (const auto& part : closure) {
        if (protectedParts.count(part)) continue;
        suppressedPreservedParts.insert(part);
        suppressedPreservedParts.insert(xlpp::internal::RelationshipGraph::relationshipsPartForSource(part));
    }
}

bool applyChartChangesToPreservedDrawing(xlpp::internal::ZipArchive& z,
                                         const xlpp::Worksheet& sheet,
                                         const std::string& sourceSheetPart,
                                         const std::string& sourceSheetXml,
                                         const std::vector<xlpp::PreservedRelationship>& allRelationships,
                                         const std::vector<xlpp::PreservedPart>& preservedParts,
                                         std::size_t& nextChartId,
                                         std::set<std::string>& suppressedPreservedParts) {
    const auto& edits = xlpp::internal::WorkbookDrawingAccess::chartEdits(sheet);
    if (sheet.appendedChartCount() == 0 && edits.empty()) return true;
    if (sourceSheetPart.empty() || sourceSheetXml.empty()) return false;

    const auto sheetRelationships = relationshipsForSource(allRelationships, sourceSheetPart);
    std::vector<std::string> ownedDrawingParts;
    std::unordered_set<std::string> seenDrawingParts;
    for (const auto& drawingNode : xlpp::internal::tags(sourceSheetXml, "drawing")) {
        const auto relationshipId = xlpp::internal::attribute(drawingNode, "r:id");
        const auto relationship = std::find_if(sheetRelationships.begin(), sheetRelationships.end(), [&](const auto& candidate) {
            return candidate.id == relationshipId && relationshipKind(candidate) == "drawing"
                && candidate.targetMode != "External";
        });
        if (relationship == sheetRelationships.end()) continue;
        const auto drawingPart = resolvePackagePart(sourceSheetPart, relationship->target);
        if (!drawingPart.empty() && seenDrawingParts.insert(drawingPart).second) ownedDrawingParts.push_back(drawingPart);
    }
    if (ownedDrawingParts.empty()) return false;
    for (const auto& edit : edits)
        if (!seenDrawingParts.count(edit.sourceDrawingPart)) return false;

    std::string appendDrawingPart = ownedDrawingParts.front();
    const auto& charts = static_cast<const xlpp::Worksheet&>(sheet).charts();
    for (std::size_t index = 0; index < std::min(sheet.loadedChartCount(), charts.size()); ++index) {
        if (seenDrawingParts.count(charts[index].sourceDrawingPart())) {
            appendDrawingPart = charts[index].sourceDrawingPart();
            break;
        }
    }
    if (sheet.loadedChartCount() == 0) {
        const auto& images = static_cast<const xlpp::Worksheet&>(sheet).images();
        for (std::size_t index = 0; index < std::min(sheet.loadedImageCount(), images.size()); ++index) {
            if (seenDrawingParts.count(images[index].sourceDrawingPart())) {
                appendDrawingPart = images[index].sourceDrawingPart();
                break;
            }
        }
    }

    for (const auto& drawingPart : ownedDrawingParts) {
        const bool appendHere = sheet.appendedChartCount() > 0 && drawingPart == appendDrawingPart;
        const bool hasEditsHere = std::any_of(edits.begin(), edits.end(), [&](const auto& edit) {
            return edit.sourceDrawingPart == drawingPart;
        });
        if (!appendHere && !hasEditsHere) continue;

        std::string drawingXmlText;
        if (z.contains(drawingPart)) drawingXmlText = z.get(drawingPart);
        else {
            const auto* raw = findPreservedPart(preservedParts, drawingPart);
            if (!raw) return false;
            drawingXmlText = raw->data;
        }
        const bool sourceStrict = drawingXmlText.find("http://purl.oclc.org/ooxml/drawingml/") != std::string::npos;
        const auto drawingRelsPart = xlpp::internal::RelationshipGraph::relationshipsPartForSource(drawingPart);
        auto drawingRelationships = z.contains(drawingRelsPart)
            ? xlpp::internal::RelationshipGraph::parseRelationshipsXml(drawingPart, z.get(drawingRelsPart))
            : relationshipsForSource(allRelationships, drawingPart);
        bool relationshipsChanged = false;

        std::unordered_map<std::string, std::string> chartWorkingCopies;
        for (const auto& edit : edits) {
            if (edit.sourceDrawingPart != drawingPart) continue;
            auto relationship = std::find_if(drawingRelationships.begin(), drawingRelationships.end(), [&](const auto& candidate) {
                return candidate.id == edit.sourceRelationshipId;
            });
            if (relationship == drawingRelationships.end()) return false;

            if (edit.removed) {
                if (!removeImportedChartAnchor(drawingXmlText, edit.stableId, edit.sourceRelationshipId)) return false;
                if (!drawingReferencesChartRelationship(drawingXmlText, edit.sourceRelationshipId)) {
                    const bool shared = std::any_of(allRelationships.begin(), allRelationships.end(), [&](const auto& candidate) {
                        if (candidate.sourcePart == drawingPart && candidate.id == edit.sourceRelationshipId) return false;
                        return candidate.targetMode != "External" &&
                            resolvePackagePart(candidate.sourcePart, candidate.target) == edit.sourceChartPart;
                    });
                    drawingRelationships.erase(relationship);
                    relationshipsChanged = true;
                    if (!shared) suppressExclusivePartClosure(edit.sourceChartPart, allRelationships, suppressedPreservedParts);
                }
                continue;
            }

            if ((edit.moved || edit.resized) &&
                !patchImportedChartAnchor(drawingXmlText, edit.stableId, edit.sourceRelationshipId,
                                          edit.anchor, edit.moved, edit.resized)) return false;

            const bool requiresChartXml = edit.titleChanged || edit.styleChanged || edit.titleRichTextChanged || edit.xAxisTitleChanged || edit.yAxisTitleChanged ||
                !edit.axisTitleEdits.empty() || !edit.axisRichTitleEdits.empty() || !edit.axisFormatEdits.empty() || !edit.areaFormatEdits.empty() || !edit.layoutEdits.empty() ||
                !edit.dataTableEdits.empty() || !edit.view3DEdits.empty() || !edit.wallFormatEdits.empty() || !edit.plotAuxiliaryEdits.empty() || !edit.plotTypeSpecificEdits.empty() || !edit.leaderLineEdits.empty() ||
                edit.legendChanged || edit.legendOverlayChanged || edit.legendLineFormatChanged || edit.legendFillFormatChanged ||
                !edit.seriesTitleEdits.empty() || !edit.seriesReferenceEdits.empty() || !edit.seriesCacheEdits.empty() ||
                !edit.plotDataLabelsEdits.empty() || !edit.seriesDataLabelsEdits.empty() || !edit.pointDataLabelEdits.empty() ||
                !edit.pointDataLabelRichTextEdits.empty() || !edit.dataPointFormatEdits.empty() ||
                !edit.seriesFormatEdits.empty() || !edit.trendlineFormatEdits.empty() || !edit.errorBarsFormatEdits.empty() ||
                !edit.trendlineEdits.empty() || !edit.errorBarsEdits.empty();
            if (!requiresChartXml) continue;
            auto chartIt = chartWorkingCopies.find(edit.sourceChartPart);
            if (chartIt == chartWorkingCopies.end()) {
                std::string chartXmlText;
                if (z.contains(edit.sourceChartPart)) chartXmlText = z.get(edit.sourceChartPart);
                else {
                    const auto* raw = findPreservedPart(preservedParts, edit.sourceChartPart);
                    if (!raw) return false;
                    chartXmlText = raw->data;
                }
                chartIt = chartWorkingCopies.emplace(edit.sourceChartPart, std::move(chartXmlText)).first;
            }
            if (edit.titleChanged && !patchImportedChartTitle(chartIt->second, edit.title)) return false;
            if (edit.styleChanged && !patchImportedChartStyle(chartIt->second, edit.style)) return false;
            if (edit.titleRichTextChanged && !patchImportedChartTitleRichText(chartIt->second, edit.titleRichText)) return false;
            const bool xyValueAxes = edit.chartType == xlpp::Chart::Type::Scatter || edit.chartType == xlpp::Chart::Type::Bubble;
            if (edit.xAxisTitleChanged) {
                if (edit.primaryXAxisId != 0) {
                    if (!patchImportedAxisTitleById(chartIt->second, edit.primaryXAxisId, edit.xAxisTitle)) return false;
                } else if (!patchImportedAxisTitle(chartIt->second,
                        xyValueAxes ? "c:valAx" : "c:catAx", xyValueAxes ? "valAx" : "catAx", 0, edit.xAxisTitle)) return false;
            }
            if (edit.yAxisTitleChanged) {
                if (edit.primaryYAxisId != 0) {
                    if (!patchImportedAxisTitleById(chartIt->second, edit.primaryYAxisId, edit.yAxisTitle)) return false;
                } else if (!patchImportedAxisTitle(chartIt->second, "c:valAx", "valAx",
                        xyValueAxes ? 1 : 0, edit.yAxisTitle)) return false;
            }
            for (const auto& axisEdit : edit.axisTitleEdits)
                if (!patchImportedAxisTitleById(chartIt->second, axisEdit.axisId, axisEdit.title)) return false;
            for (const auto& richEdit : edit.axisRichTitleEdits)
                if (!patchImportedAxisTitleRichTextById(chartIt->second, richEdit.axisId, richEdit.richText)) return false;
            for (const auto& axisEdit : edit.axisFormatEdits) {
                using Kind = std::decay_t<decltype(axisEdit.kind)>;
                if (axisEdit.kind == Kind::NumberFormat) { if (!patchImportedAxisNumberFormat(chartIt->second, axisEdit.axisId, axisEdit.value1, axisEdit.flag)) return false; }
                else if (axisEdit.kind == Kind::Ticks) { if (!patchImportedAxisTicks(chartIt->second, axisEdit.axisId, axisEdit.value1, axisEdit.value2, axisEdit.value3)) return false; }
                else if (axisEdit.kind == Kind::Units) { if (!patchImportedAxisUnits(chartIt->second, axisEdit.axisId, axisEdit.number1, axisEdit.number2)) return false; }
                else if (axisEdit.kind == Kind::Scaling) { if (!patchImportedAxisScaling(chartIt->second, axisEdit.axisId, axisEdit.scaling)) return false; }
                else if (axisEdit.kind == Kind::Crossing) { if (!patchImportedAxisCrossing(chartIt->second, axisEdit.axisId, axisEdit.value1, axisEdit.value2)) return false; }
                else if (axisEdit.kind == Kind::CrossesAt) { if (!patchImportedAxisCrossesAt(chartIt->second, axisEdit.axisId, axisEdit.number1, false)) return false; }
                else if (axisEdit.kind == Kind::ClearCrossesAt) { if (!patchImportedAxisCrossesAt(chartIt->second, axisEdit.axisId, 0.0, true)) return false; }
                else if (axisEdit.kind == Kind::DisplayUnits) { if (!patchImportedAxisDisplayUnits(chartIt->second, axisEdit.axisId, &axisEdit.displayUnits)) return false; }
                else if (axisEdit.kind == Kind::ClearDisplayUnits) { if (!patchImportedAxisDisplayUnits(chartIt->second, axisEdit.axisId, nullptr)) return false; }
                else if (axisEdit.kind == Kind::Line) { if (!patchImportedAxisLineFormat(chartIt->second, axisEdit.axisId, axisEdit.line)) return false; }
                else if (axisEdit.kind == Kind::MajorGridline) { if (!patchImportedAxisGridlineFormat(chartIt->second, axisEdit.axisId, true, axisEdit.line)) return false; }
                else if (axisEdit.kind == Kind::MinorGridline) { if (!patchImportedAxisGridlineFormat(chartIt->second, axisEdit.axisId, false, axisEdit.line)) return false; }
                else if (axisEdit.kind == Kind::RemoveMajorGridline) { if (!removeImportedAxisGridlines(chartIt->second, axisEdit.axisId, true)) return false; }
                else if (!removeImportedAxisGridlines(chartIt->second, axisEdit.axisId, false)) return false;
            }
            for (const auto& areaEdit : edit.areaFormatEdits) {
                using Owner = std::decay_t<decltype(areaEdit.owner)>;
                using Kind = std::decay_t<decltype(areaEdit.kind)>;
                const bool chartArea = areaEdit.owner == Owner::ChartArea;
                if (areaEdit.kind == Kind::Line) { if (!patchImportedAreaFormat(chartIt->second, chartArea, &areaEdit.line, nullptr)) return false; }
                else if (!patchImportedAreaFormat(chartIt->second, chartArea, nullptr, &areaEdit.fill)) return false;
            }
            for (const auto& layoutEdit : edit.layoutEdits) {
                using Owner = std::decay_t<decltype(layoutEdit.owner)>;
                if (layoutEdit.owner == Owner::PlotArea) { if (!patchImportedPlotAreaLayout(chartIt->second, layoutEdit.layout)) return false; }
                else if (!patchImportedLegendLayout(chartIt->second, layoutEdit.layout)) return false;
            }
            for (const auto& tableEdit : edit.dataTableEdits)
                if (!patchImportedChartDataTable(chartIt->second, tableEdit.remove ? nullptr : &tableEdit.table)) return false;
            for (const auto& viewEdit : edit.view3DEdits)
                if (!patchImportedChartView3D(chartIt->second, viewEdit.view)) return false;
            for (const auto& wallEdit : edit.wallFormatEdits) {
                using Owner = std::decay_t<decltype(wallEdit.owner)>;
                if (wallEdit.owner == Owner::Floor) { if (!patchImportedChartWallFormat(chartIt->second, "c:floor", "floor", wallEdit.format)) return false; }
                else if (wallEdit.owner == Owner::SideWall) { if (!patchImportedChartWallFormat(chartIt->second, "c:sideWall", "sideWall", wallEdit.format)) return false; }
                else if (!patchImportedChartWallFormat(chartIt->second, "c:backWall", "backWall", wallEdit.format)) return false;
            }
            for (const auto& auxEdit : edit.plotAuxiliaryEdits) {
                using Kind = std::decay_t<decltype(auxEdit.kind)>;
                if (auxEdit.kind == Kind::DropLines) {
                    if (!patchImportedChartPlotLineObject(chartIt->second, auxEdit.plotIndex, "c:dropLines", "dropLines", auxEdit.remove ? nullptr : &auxEdit.line)) return false;
                } else if (auxEdit.kind == Kind::HighLowLines) {
                    if (!patchImportedChartPlotLineObject(chartIt->second, auxEdit.plotIndex, "c:hiLowLines", "hiLowLines", auxEdit.remove ? nullptr : &auxEdit.line)) return false;
                } else if (!patchImportedChartPlotUpDownBars(chartIt->second, auxEdit.plotIndex, auxEdit.remove ? nullptr : &auxEdit.upDownBars)) return false;
            }
            for (const auto& typeEdit : edit.plotTypeSpecificEdits) {
                using Kind = std::decay_t<decltype(typeEdit.kind)>;
                if (typeEdit.kind == Kind::FirstSliceAngle) { if (!patchImportedChartPlotSimpleValue(chartIt->second,typeEdit.plotIndex,"c:firstSliceAng","firstSliceAng",std::to_string(typeEdit.integerValue),false)) return false; }
                else if (typeEdit.kind == Kind::DoughnutHoleSize) { if (!patchImportedChartPlotSimpleValue(chartIt->second,typeEdit.plotIndex,"c:holeSize","holeSize",std::to_string(typeEdit.integerValue),false)) return false; }
                else if (typeEdit.kind == Kind::RadarStyle) { if (!patchImportedChartPlotSimpleValue(chartIt->second,typeEdit.plotIndex,"c:radarStyle","radarStyle",typeEdit.textValue,true)) return false; }
                else if (!patchImportedChartProjectedPie(chartIt->second,typeEdit.plotIndex,typeEdit.projectedPie)) return false;
            }
            for (const auto& leaderEdit : edit.leaderLineEdits)
                if (!patchImportedChartLeaderLines(chartIt->second, leaderEdit.plotLevel, leaderEdit.ownerIndex, leaderEdit.remove ? nullptr : &leaderEdit.line, leaderEdit.remove)) return false;
            if (edit.legendChanged && !patchImportedChartLegend(chartIt->second, edit.showLegend, edit.legendPosition)) return false;
            if (edit.legendOverlayChanged && !patchImportedLegendOverlay(chartIt->second, edit.legendOverlay)) return false;
            if (edit.legendLineFormatChanged && !patchImportedLegendFormat(chartIt->second, &edit.legendLineFormat, nullptr)) return false;
            if (edit.legendFillFormatChanged && !patchImportedLegendFormat(chartIt->second, nullptr, &edit.legendFillFormat)) return false;
            for (const auto& seriesEdit : edit.seriesTitleEdits)
                if (!patchImportedChartSeriesTitle(chartIt->second, seriesEdit.seriesIndex, seriesEdit.title)) return false;
            for (const auto& seriesEdit : edit.seriesReferenceEdits) {
                if (!patchImportedChartSeriesReferences(chartIt->second, seriesEdit.seriesIndex,
                                                        seriesEdit.categoriesReference,
                                                        seriesEdit.valuesReference)) return false;
            }
            for (const auto& cacheEdit : edit.seriesCacheEdits) {
                using Kind = std::decay_t<decltype(cacheEdit.kind)>;
                int kind = cacheEdit.kind == Kind::Categories ? 0 : (cacheEdit.kind == Kind::Values ? 1 : (cacheEdit.kind == Kind::Title ? 2 : 3));
                if (!patchImportedChartSeriesCache(chartIt->second, cacheEdit.seriesIndex, kind,
                                                   cacheEdit.kind == Kind::ClearAll ? nullptr : &cacheEdit.cache)) return false;
            }
            for (const auto& labelsEdit : edit.plotDataLabelsEdits)
                if (!patchImportedChartPlotDataLabels(chartIt->second, labelsEdit.plotIndex, labelsEdit.labels)) return false;
            for (const auto& labelsEdit : edit.seriesDataLabelsEdits)
                if (!patchImportedChartSeriesDataLabels(chartIt->second, labelsEdit.seriesIndex, labelsEdit.labels)) return false;
            for (const auto& pointEdit : edit.pointDataLabelEdits) {
                if (pointEdit.plotLevel) {
                    if (!patchImportedChartPlotDataLabelPoint(chartIt->second, pointEdit.ownerIndex, pointEdit.label, pointEdit.remove)) return false;
                } else if (!patchImportedChartSeriesDataLabelPoint(chartIt->second, pointEdit.ownerIndex, pointEdit.label, pointEdit.remove)) return false;
            }
            for (const auto& richEdit : edit.pointDataLabelRichTextEdits)
                if (!patchImportedChartSeriesDataLabelPointRichText(chartIt->second, richEdit.seriesIndex, richEdit.pointIndex, richEdit.richText)) return false;
            for (const auto& pointFormatEdit : edit.dataPointFormatEdits)
                if (!patchImportedChartSeriesDataPointFormat(chartIt->second, pointFormatEdit.seriesIndex, pointFormatEdit.format, pointFormatEdit.remove)) return false;
            for (const auto& formatEdit : edit.seriesFormatEdits) {
                using Kind = std::decay_t<decltype(formatEdit.kind)>;
                if (formatEdit.kind == Kind::Line) {
                    if (!patchSeriesLineOrFill(chartIt->second, formatEdit.seriesIndex, &formatEdit.line, nullptr)) return false;
                } else if (formatEdit.kind == Kind::Fill) {
                    if (!patchSeriesLineOrFill(chartIt->second, formatEdit.seriesIndex, nullptr, &formatEdit.fill)) return false;
                } else if (!patchSeriesMarkerFormat(chartIt->second, formatEdit.seriesIndex, formatEdit.marker)) return false;
            }
            for (const auto& trendlineEdit : edit.trendlineEdits) {
                using Action = std::decay_t<decltype(trendlineEdit.action)>;
                if (trendlineEdit.action == Action::Add) {
                    if (!patchImportedChartSeriesTrendline(chartIt->second, trendlineEdit.seriesIndex, trendlineEdit.trendlineIndex,
                                                           &trendlineEdit.trendline, true)) return false;
                } else if (trendlineEdit.action == Action::Remove) {
                    if (!patchImportedChartSeriesTrendline(chartIt->second, trendlineEdit.seriesIndex, trendlineEdit.trendlineIndex,
                                                           nullptr, false)) return false;
                } else if (!patchImportedChartSeriesTrendline(chartIt->second, trendlineEdit.seriesIndex, trendlineEdit.trendlineIndex,
                                                               &trendlineEdit.trendline, false)) return false;
            }
            for (const auto& barsEdit : edit.errorBarsEdits)
                if (!patchImportedChartSeriesErrorBars(chartIt->second, barsEdit.seriesIndex, barsEdit.direction,
                                                       barsEdit.remove ? nullptr : &barsEdit.errorBars)) return false;
            for (const auto& formatEdit : edit.trendlineFormatEdits)
                if (!patchImportedChartSeriesTrendlineLineFormat(chartIt->second, formatEdit.seriesIndex, formatEdit.trendlineIndex, formatEdit.line)) return false;
            for (const auto& formatEdit : edit.errorBarsFormatEdits)
                if (!patchImportedChartSeriesErrorBarsLineFormat(chartIt->second, formatEdit.seriesIndex, formatEdit.direction, formatEdit.line)) return false;
        }

        if (appendHere) {
            std::set<std::string> usedRelationshipIds;
            for (const auto& relationship : drawingRelationships) usedRelationshipIds.insert(relationship.id);
            std::size_t objectId = maximumDrawingObjectId(drawingXmlText) + 1;
            std::string appendedAnchors;
            for (std::size_t index = sheet.loadedChartCount(); index < charts.size(); ++index) {
                const auto& chart = charts[index];
                const auto relationshipId = allocateRelationshipId(usedRelationshipIds);
                usedRelationshipIds.insert(relationshipId);
                const auto chartId = nextChartId++;
                xlpp::PreservedRelationship relationship;
                relationship.sourcePart = drawingPart;
                relationship.id = relationshipId;
                relationship.type = nsRelsDoc(sourceStrict) + "/chart";
                relationship.target = "../charts/chart" + std::to_string(chartId) + ".xml";
                drawingRelationships.push_back(std::move(relationship));
                relationshipsChanged = true;
                appendedAnchors += appendedChartAnchorXml(chart, relationshipId, objectId++, index, sourceStrict);
                z.add("xl/charts/chart" + std::to_string(chartId) + ".xml", chartXml(chart, sourceStrict));
            }
            if (!appendedAnchors.empty()) {
                const auto closing = drawingXmlText.rfind("</xdr:wsDr>") != std::string::npos
                    ? std::string("</xdr:wsDr>") : std::string("</wsDr>");
                const auto closingPosition = drawingXmlText.rfind(closing);
                if (closingPosition == std::string::npos) return false;
                drawingXmlText.insert(closingPosition, appendedAnchors);
            }
        }

        z.add(drawingPart, std::move(drawingXmlText));
        if (relationshipsChanged)
            z.add(drawingRelsPart, xlpp::internal::RelationshipGraph::serializeRelationships(drawingRelationships, sourceStrict));
        for (auto& [part, data] : chartWorkingCopies) z.add(part, std::move(data));
    }
    return true;
}

bool relationshipTargetsPart(const xlpp::PreservedRelationship& relationship, const std::string& part) {
    return relationship.targetMode != "External"
        && resolvePackagePart(relationship.sourcePart, relationship.target) == part;
}

bool hasOtherRelationshipToPart(const std::vector<xlpp::PreservedRelationship>& relationships,
                                const std::string& sourcePart,
                                const std::string& relationshipId,
                                const std::string& part) {
    return std::any_of(relationships.begin(), relationships.end(), [&](const auto& relationship) {
        if (relationship.sourcePart == sourcePart && relationship.id == relationshipId) return false;
        return relationshipTargetsPart(relationship, part);
    });
}

bool applyImageChangesToPreservedDrawing(xlpp::internal::ZipArchive& z,
                                         const xlpp::Worksheet& sheet,
                                         const std::string& sourceSheetPart,
                                         const std::string& sourceSheetXml,
                                         const std::vector<xlpp::PreservedRelationship>& allRelationships,
                                         const std::vector<xlpp::PreservedPart>& preservedParts,
                                         std::size_t& nextMediaId,
                                         std::set<std::string>& suppressedPreservedParts) {
    const auto& edits = xlpp::internal::WorkbookDrawingAccess::imageEdits(sheet);
    if (sheet.appendedImageCount() == 0 && edits.empty()) return true;
    if (sourceSheetPart.empty()) return false;

    if (sourceSheetXml.empty()) return false;
    const auto sheetRelationships = relationshipsForSource(allRelationships, sourceSheetPart);

    // A normal worksheet owns one DrawingML part, but preserving producer data
    // means we must not assume that.  Resolve every explicit <drawing r:id>
    // owner node and patch only the drawing part that owns each imported image.
    std::vector<std::string> ownedDrawingParts;
    std::unordered_set<std::string> seenDrawingParts;
    for (const auto& drawingNode : xlpp::internal::tags(sourceSheetXml, "drawing")) {
        const auto relationshipId = xlpp::internal::attribute(drawingNode, "r:id");
        const auto relationship = std::find_if(sheetRelationships.begin(), sheetRelationships.end(), [&](const auto& candidate) {
            return candidate.id == relationshipId && relationshipKind(candidate) == "drawing"
                && candidate.targetMode != "External";
        });
        if (relationship == sheetRelationships.end()) continue;
        const auto drawingPart = resolvePackagePart(sourceSheetPart, relationship->target);
        if (!drawingPart.empty() && seenDrawingParts.insert(drawingPart).second)
            ownedDrawingParts.push_back(drawingPart);
    }
    if (ownedDrawingParts.empty()) return false;

    for (const auto& edit : edits)
        if (!seenDrawingParts.count(edit.sourceDrawingPart)) return false;

    // New images are appended to the first preserved drawing by default.  If
    // the sheet already has imported images, prefer the drawing that owns the
    // first one so append behavior stays stable for multi-drawing producers.
    std::string appendDrawingPart = ownedDrawingParts.front();
    for (std::size_t index = 0; index < std::min(sheet.loadedImageCount(), sheet.images().size()); ++index) {
        const auto& candidate = sheet.images()[index];
        if (seenDrawingParts.count(candidate.sourceDrawingPart())) {
            appendDrawingPart = candidate.sourceDrawingPart();
            break;
        }
    }

    for (const auto& drawingPart : ownedDrawingParts) {
        const bool appendHere = sheet.appendedImageCount() > 0 && drawingPart == appendDrawingPart;
        const bool hasEditsHere = std::any_of(edits.begin(), edits.end(), [&](const auto& edit) {
            return edit.sourceDrawingPart == drawingPart;
        });
        if (!appendHere && !hasEditsHere) continue; // keep untouched drawing bytes exactly as loaded

        const auto* rawDrawingPart = findPreservedPart(preservedParts, drawingPart);
        if (!rawDrawingPart) return false;
        auto drawingXmlText = rawDrawingPart->data;
        const bool sourceStrict = drawingXmlText.find("http://purl.oclc.org/ooxml/drawingml/") != std::string::npos;
        auto drawingRelationships = relationshipsForSource(allRelationships, drawingPart);

        for (const auto& edit : edits) {
            if (edit.sourceDrawingPart != drawingPart) continue;
            if (!patchImportedImageAnchor(drawingXmlText, edit.stableId, edit.sourceRelationshipId, edit.anchor,
                                          edit.moved, edit.resized, edit.removed)) return false;
            auto relationship = std::find_if(drawingRelationships.begin(), drawingRelationships.end(), [&](const auto& candidate) {
                return candidate.id == edit.sourceRelationshipId;
            });
            if (relationship == drawingRelationships.end()) return false;

            if (edit.removed) {
                if (!drawingReferencesRelationship(drawingXmlText, edit.sourceRelationshipId)) {
                    const bool shared = hasOtherRelationshipToPart(allRelationships, drawingPart,
                                                                   edit.sourceRelationshipId, edit.sourceMediaPart);
                    drawingRelationships.erase(relationship);
                    if (!shared) suppressedPreservedParts.insert(edit.sourceMediaPart);
                }
                continue;
            }
            if (edit.replaced) {
                const bool shared = hasOtherRelationshipToPart(allRelationships, drawingPart,
                                                               edit.sourceRelationshipId, edit.sourceMediaPart);
                const auto oldExtension = partExtension(edit.sourceMediaPart);
                if (!shared && oldExtension == edit.replacementExtension) {
                    const std::string bytes(reinterpret_cast<const char*>(edit.replacementBytes.data()), edit.replacementBytes.size());
                    z.add(edit.sourceMediaPart, bytes, false);
                } else {
                    const auto mediaId = nextMediaId++;
                    const auto newMediaPart = "xl/media/image" + std::to_string(mediaId) + "." + edit.replacementExtension;
                    relationship->target = "../media/image" + std::to_string(mediaId) + "." + edit.replacementExtension;
                    const std::string bytes(reinterpret_cast<const char*>(edit.replacementBytes.data()), edit.replacementBytes.size());
                    z.add(newMediaPart, bytes, false);
                    if (!shared) suppressedPreservedParts.insert(edit.sourceMediaPart);
                }
            }
        }

        if (appendHere) {
            std::set<std::string> usedRelationshipIds;
            for (const auto& relationship : drawingRelationships) usedRelationshipIds.insert(relationship.id);
            std::size_t objectId = maximumDrawingObjectId(drawingXmlText) + 1;
            std::string appendedAnchors;
            const auto& images = sheet.images();
            for (std::size_t index = sheet.loadedImageCount(); index < images.size(); ++index) {
                const auto& image = images[index];
                const auto relationshipId = allocateRelationshipId(usedRelationshipIds);
                usedRelationshipIds.insert(relationshipId);
                const auto mediaId = nextMediaId++;
                xlpp::PreservedRelationship relationship;
                relationship.sourcePart = drawingPart;
                relationship.id = relationshipId;
                relationship.type = nsRelsDoc(sourceStrict) + "/image";
                relationship.target = "../media/image" + std::to_string(mediaId) + "." + image.extension();
                drawingRelationships.push_back(std::move(relationship));
                appendedAnchors += appendedImageAnchorXml(image, relationshipId, objectId++, sourceStrict);
                const std::string bytes(reinterpret_cast<const char*>(image.bytes().data()), image.bytes().size());
                z.add("xl/media/image" + std::to_string(mediaId) + "." + image.extension(), bytes, false);
            }
            if (!appendedAnchors.empty()) {
                const auto closing = drawingXmlText.rfind("</xdr:wsDr>") != std::string::npos ? std::string("</xdr:wsDr>") : std::string("</wsDr>");
                const auto closingPosition = drawingXmlText.rfind(closing);
                if (closingPosition == std::string::npos) return false;
                drawingXmlText.insert(closingPosition, appendedAnchors);
            }
        }

        z.add(drawingPart, std::move(drawingXmlText));
        const auto slash = drawingPart.find_last_of('/');
        const auto drawingFile = drawingPart.substr(slash + 1);
        const auto drawingRelsPart = drawingPart.substr(0, slash + 1) + "_rels/" + drawingFile + ".rels";
        z.add(drawingRelsPart, xlpp::internal::RelationshipGraph::serializeRelationships(drawingRelationships, sourceStrict));
    }
    return true;
}

std::string mergePivotTablePartsBlocks(const std::vector<std::string>& originalBlocks,
                                      const std::vector<std::string>& generatedBlocks,
                                      bool preserveOriginal) {
    if (!preserveOriginal || originalBlocks.empty()) return joinBlocks(generatedBlocks);
    if (generatedBlocks.empty()) return joinBlocks(originalBlocks);

    std::vector<std::string> nodes;
    for (const auto& block : originalBlocks) {
        const auto items = extractTagBlocks(block, "pivotTablePart");
        nodes.insert(nodes.end(), items.begin(), items.end());
    }
    for (const auto& block : generatedBlocks) {
        const auto items = extractTagBlocks(block, "pivotTablePart");
        nodes.insert(nodes.end(), items.begin(), items.end());
    }
    if (nodes.empty()) return {};
    std::ostringstream xml;
    xml << "<pivotTableParts count=\"" << nodes.size() << "\">";
    for (const auto& node : nodes) xml << node;
    xml << "</pivotTableParts>";
    return xml.str();
}

std::string rebuildWorksheetTail(std::string generated,
                                 const std::string& original,
                                 bool preserveDrawing,
                                 bool preservePivot,
                                 bool preserveTables,
                                 bool preserveComments) {
    static const std::array<const char*, 10> tailTags{
        "drawing", "legacyDrawing", "legacyDrawingHF", "picture", "oleObjects",
        "controls", "webPublishItems", "tableParts", "pivotTableParts", "extLst"
    };
    std::unordered_map<std::string, std::vector<std::string>> generatedBlocks;
    std::unordered_map<std::string, std::vector<std::string>> originalBlocks;
    for (const auto* tag : tailTags) {
        generatedBlocks[tag] = extractTagBlocks(generated, tag);
        originalBlocks[tag] = extractTagBlocks(original, tag);
        eraseTagBlocks(generated, tag);
    }

    std::string tail;
    tail += joinBlocks(preserveDrawing && !originalBlocks["drawing"].empty()
        ? originalBlocks["drawing"] : generatedBlocks["drawing"]);
    tail += joinBlocks(preserveComments && !originalBlocks["legacyDrawing"].empty()
        ? originalBlocks["legacyDrawing"] : generatedBlocks["legacyDrawing"]);
    for (const auto* tag : {"legacyDrawingHF", "picture", "oleObjects", "controls", "webPublishItems"})
        tail += joinBlocks(!originalBlocks[tag].empty() ? originalBlocks[tag] : generatedBlocks[tag]);
    tail += joinBlocks(preserveTables && !originalBlocks["tableParts"].empty()
        ? originalBlocks["tableParts"] : generatedBlocks["tableParts"]);
    tail += mergePivotTablePartsBlocks(originalBlocks["pivotTableParts"], generatedBlocks["pivotTableParts"], preservePivot);
    tail += joinBlocks(!originalBlocks["extLst"].empty() ? originalBlocks["extLst"] : generatedBlocks["extLst"]);
    insertBefore(generated, "</worksheet>", tail);
    return generated;
}

std::string workbookViewsXml(const std::string& sourceWorkbookXml, std::size_t activeTab, std::size_t firstSheet) {
    const auto blocks = extractTagBlocks(sourceWorkbookXml, "bookViews");
    if (!blocks.empty()) {
        auto result = blocks.front();
        const auto views = extractTagBlocks(result, "workbookView");
        if (!views.empty()) {
            auto patched = views.front();
            patchOpeningTagAttribute(patched, "activeTab", std::to_string(activeTab), false);
            patchOpeningTagAttribute(patched, "firstSheet", std::to_string(firstSheet), false);
            const auto position = result.find(views.front());
            if (position != std::string::npos) result.replace(position, views.front().size(), patched);
            return result;
        }
    }
    return "<bookViews><workbookView activeTab=\"" + std::to_string(activeTab)
        + "\" firstSheet=\"" + std::to_string(firstSheet) + "\"/></bookViews>";
}

std::string preserveWorkbookNodes(std::string generated,
                                  const std::string& original,
                                  bool preservePivotCaches) {
    if (original.empty()) return generated;
    const auto preserveBeforeWorkbookPr = [&](const char* tag) {
        const auto blocks = extractTagBlocks(original, tag);
        if (!blocks.empty() && extractTagBlocks(generated, tag).empty())
            insertBefore(generated, "<workbookPr", joinBlocks(blocks));
    };
    preserveBeforeWorkbookPr("fileVersion");
    preserveBeforeWorkbookPr("fileSharing");

    if (extractTagBlocks(generated, "bookViews").empty())
        insertBefore(generated, "<sheets>", joinBlocks(extractTagBlocks(original, "bookViews")));

    const auto sheetsEnd = generated.find("</sheets>");
    if (sheetsEnd != std::string::npos) {
        std::string afterSheets;
        for (const auto* tag : {"functionGroups", "externalReferences"})
            if (extractTagBlocks(generated, tag).empty()) afterSheets += joinBlocks(extractTagBlocks(original, tag));
        generated.insert(sheetsEnd + std::string("</sheets>").size(), afterSheets);
    }

    std::string finalNodes;
    for (const auto* tag : {"oleSize", "customWorkbookViews"})
        if (extractTagBlocks(generated, tag).empty()) finalNodes += joinBlocks(extractTagBlocks(original, tag));
    if (preservePivotCaches && extractTagBlocks(generated, "pivotCaches").empty())
        finalNodes += joinBlocks(extractTagBlocks(original, "pivotCaches"));
    for (const auto* tag : {"webPublishing", "fileRecoveryPr", "webPublishObjects", "extLst"})
        if (extractTagBlocks(generated, tag).empty()) finalNodes += joinBlocks(extractTagBlocks(original, tag));
    insertBefore(generated, "</workbook>", finalNodes);
    return generated;
}

std::string mergeWorkbookPivotCaches(const std::string& originalWorkbookXml,
                                     const std::string& generatedPivotCachesXml) {
    const auto originalContainers = extractTagBlocks(originalWorkbookXml, "pivotCaches");
    if (originalContainers.empty()) return generatedPivotCachesXml;
    if (generatedPivotCachesXml.empty()) return joinBlocks(originalContainers);

    std::vector<std::string> nodes;
    for (const auto& block : originalContainers) {
        const auto items = extractTagBlocks(block, "pivotCache");
        nodes.insert(nodes.end(), items.begin(), items.end());
    }
    const auto generatedContainers = extractTagBlocks(generatedPivotCachesXml, "pivotCaches");
    for (const auto& block : generatedContainers) {
        const auto items = extractTagBlocks(block, "pivotCache");
        nodes.insert(nodes.end(), items.begin(), items.end());
    }
    if (nodes.empty()) return {};
    std::ostringstream xml;
    xml << "<pivotCaches>";
    for (const auto& node : nodes) xml << node;
    xml << "</pivotCaches>";
    return xml.str();
}

std::size_t nextAvailablePivotCacheId(const std::string& workbookXml) {
    std::size_t maximum = 0;
    for (const auto& node : extractTagBlocks(workbookXml, "pivotCache")) {
        const auto value = xlpp::internal::attribute(node, "cacheId");
        if (value.empty() || !std::all_of(value.begin(), value.end(), [](unsigned char c) { return std::isdigit(c); })) continue;
        unsigned long long parsed = 0;
        if (xlpp::internal::tryParseIntegerExact(value, parsed))
            maximum = std::max(maximum, static_cast<std::size_t>(parsed));
    }
    return maximum + 1;
}

std::size_t nextAvailablePartId(const std::vector<xlpp::PreservedPart>& parts,
                                const std::string& prefix,
                                const std::string& suffix) {
    std::size_t maximum = 0;
    for (const auto& part : parts) {
        if (part.name.rfind(prefix, 0) != 0 || part.name.size() <= prefix.size() + suffix.size()) continue;
        if (part.name.compare(part.name.size() - suffix.size(), suffix.size(), suffix) != 0) continue;
        const auto number = part.name.substr(prefix.size(), part.name.size() - prefix.size() - suffix.size());
        if (number.empty() || !std::all_of(number.begin(), number.end(), [](unsigned char c) { return std::isdigit(c); })) continue;
        unsigned long long parsed = 0;
        if (xlpp::internal::tryParseIntegerExact(number, parsed))
            maximum = std::max(maximum, static_cast<std::size_t>(parsed));
    }
    return maximum + 1;
}


std::size_t nextAvailableMediaId(const std::vector<xlpp::PreservedPart>& parts) {
    std::size_t maximum = 0;
    const std::string prefix = "xl/media/image";
    for (const auto& part : parts) {
        if (part.name.rfind(prefix, 0) != 0) continue;
        const auto dot = part.name.find('.', prefix.size());
        if (dot == std::string::npos) continue;
        const auto number = part.name.substr(prefix.size(), dot - prefix.size());
        if (number.empty() || !std::all_of(number.begin(), number.end(), [](unsigned char c) { return std::isdigit(c); })) continue;
        unsigned long long parsed = 0;
        if (xlpp::internal::tryParseIntegerExact(number, parsed))
            maximum = std::max(maximum, static_cast<std::size_t>(parsed));
    }
    return maximum + 1;
}

} // namespace internal
} // namespace xlpp

