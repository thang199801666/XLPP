#include "OOXML/Charts/ChartSerializer.h"
#include "OOXML/Charts/ChartFormatCodec.h"
#include "OOXML/Charts/ChartXmlSupport.h"
#include "OOXML/Common/DrawingXmlSupport.h"
#include "Package/Xml/XmlUtilities.h"
#include <XLPP/Chart/Chart.h>
#include <algorithm>
#include <array>
#include <cmath>
#include <cstddef>
#include <cstdint>
#include <iomanip>
#include <limits>
#include <sstream>
#include <string>
#include <utility>
#include <vector>


namespace xlpp::internal::ooxml {

std::string chartView3DXml(const xlpp::ChartView3D& view, bool prefixed);
std::string generatedChartWallXml(const char* localName, const xlpp::ChartWallFormat& format, bool prefixed);
std::string chartTextStyleTxPrXml(const xlpp::ChartTextStyle& style, bool prefixed, bool strict);
std::string generatedDataLabelsXml(const xlpp::ChartDataLabels& labels);
std::string generatedPlotAuxiliaryXml(const xlpp::Chart::Plot& plot, bool strict);
std::string generatedDataTableXml(const xlpp::ChartDataTable& table, bool strict);
std::string chartSeriesCacheXml(const xlpp::ChartSeriesCache& cache, bool prefixed) {
    if (!cache.present) return {};
    const auto c = prefixed ? "c:" : "";
    const auto local = cache.numeric ? "numCache" : "strCache";
    std::ostringstream xml;
    xml << "<" << c << local << ">";
    if (cache.numeric) xml << "<" << c << "formatCode>" << xmlEscape(cache.formatCode.empty() ? "General" : cache.formatCode) << "</" << c << "formatCode>";
    xml << "<" << c << "ptCount val=\"" << cache.effectivePointCount() << "\"/>";
    auto points = cache.points;
    std::sort(points.begin(), points.end(), [](const auto& a, const auto& b){ return a.index < b.index; });
    for (const auto& point : points)
        xml << "<" << c << "pt idx=\"" << point.index << "\"><" << c << "v>" << xmlEscape(point.value) << "</" << c << "v></" << c << "pt>";
    xml << "</" << c << local << ">";
    return xml.str();
}

std::string chartExFormula(std::string value) {
    if (!value.empty() && value.front() == '=') value.erase(value.begin());
    return value;
}

std::string chartExCacheLevelXml(const xlpp::ChartSeriesCache& cache, bool numeric) {
    if (!cache.present) return {};
    std::ostringstream xml;
    xml << "<cx:lvl ptCount=\"" << cache.effectivePointCount() << "\"";
    if (numeric) xml << " formatCode=\"" << xmlEscape(cache.formatCode.empty() ? "General" : cache.formatCode) << "\"";
    xml << ">";
    auto points = cache.points;
    std::sort(points.begin(), points.end(), [](const auto& a, const auto& b){ return a.index < b.index; });
    for (const auto& point : points)
        xml << "<cx:pt idx=\"" << point.index << "\">" << xmlEscape(point.value) << "</cx:pt>";
    xml << "</cx:lvl>";
    return xml.str();
}

const char* chartExLayoutId(xlpp::Chart::Type type) {
    using T = xlpp::Chart::Type;
    switch (type) {
        case T::BoxWhisker: return "boxWhisker";
        case T::Funnel: return "funnel";
        case T::Pareto: return "clusteredColumn";
        case T::FilledMap: return "regionMap";
        case T::Sunburst: return "sunburst";
        case T::Treemap: return "treemap";
        case T::Waterfall: return "waterfall";
        case T::Histogram: return "clusteredColumn";
        default: return "clusteredColumn";
    }
}

std::string chartExLayoutPropertiesXml(const xlpp::Chart& chart) {
    const auto type = chart.type();
    const auto* plot = chart.primaryPlotOrNull();
    std::ostringstream xml;
    using T = xlpp::Chart::Type;
    if (type == T::Histogram || type == T::Pareto) {
        xml << "<cx:layoutPr><cx:binning intervalClosed=\"l\"";
        if (plot && plot->histogramHasUnderflow) xml << " underflow=\"" << plot->histogramUnderflow << "\"";
        if (plot && plot->histogramHasOverflow) xml << " overflow=\"" << plot->histogramOverflow << "\"";
        xml << ">";
        if (plot && !plot->histogramAutomaticBins && plot->histogramBinWidth > 0.0) xml << "<cx:binSize val=\"" << plot->histogramBinWidth << "\"/>";
        else if (plot && !plot->histogramAutomaticBins && plot->histogramBinCount > 0) xml << "<cx:binCount val=\"" << plot->histogramBinCount << "\"/>";
        xml << "</cx:binning></cx:layoutPr>";
    } else if (type == T::BoxWhisker) {
        xml << "<cx:layoutPr><cx:visibility connectorLines=\"0\" meanLine=\"" << (plot && plot->boxWhiskerShowMeanLine ? 1 : 0)
            << "\" meanMarker=\"" << (plot && plot->boxWhiskerShowMeanMarker ? 1 : 0)
            << "\" nonoutliers=\"" << (plot && plot->boxWhiskerShowInnerPoints ? 1 : 0)
            << "\" outliers=\"" << (!plot || plot->boxWhiskerShowOutlierPoints ? 1 : 0) << "\"/>"
            << "<cx:statistics quartileMethod=\"" << ((plot && plot->boxWhiskerQuartileInclusive) ? "inclusive" : "exclusive") << "\"/></cx:layoutPr>";
    } else if (type == T::Waterfall) {
        xml << "<cx:layoutPr><cx:visibility connectorLines=\"" << (!plot || plot->waterfallShowConnectorLines ? 1 : 0) << "\"/></cx:layoutPr>";
    } else if (type == T::Treemap || type == T::Sunburst) {
        xml << "<cx:layoutPr><cx:parentLabelLayout val=\"banner\"/></cx:layoutPr>";
    }
    return xml.str();
}

std::string serializeChartEx(const xlpp::Chart& chart, bool /*strict*/) {
    if (!xlpp::Chart::isModernType(chart.type())) throw std::invalid_argument("ChartEx serialization requires a modern Excel chart type");
    std::ostringstream xml;
    xml << "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
        << "<cx:chartSpace xmlns:cx=\"http://schemas.microsoft.com/office/drawing/2014/chartex\" "
        << "xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" "
        << "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">";
    xml << "<cx:chartData>";
    for (std::size_t s = 0; s < chart.series().size(); ++s) {
        const auto& series = chart.series()[s];
        xml << "<cx:data id=\"" << s << "\">";
        if (!series.categoriesReference().empty() || series.categoriesCache().present) {
            xml << "<cx:strDim type=\"cat\">";
            if (!series.categoriesReference().empty()) xml << "<cx:f>" << xmlEscape(chartExFormula(series.categoriesReference())) << "</cx:f>";
            const auto level = chartExCacheLevelXml(series.categoriesCache(), false);
            if (!level.empty()) xml << level; else xml << "<cx:lvl ptCount=\"0\"/>";
            xml << "</cx:strDim>";
        }
        if (!series.valuesReference().empty() || series.valuesCache().present) {
            const bool sizeDimension = chart.type() == xlpp::Chart::Type::Treemap || chart.type() == xlpp::Chart::Type::Sunburst;
            xml << "<cx:numDim type=\"" << (sizeDimension ? "size" : "val") << "\">";
            if (!series.valuesReference().empty()) xml << "<cx:f>" << xmlEscape(chartExFormula(series.valuesReference())) << "</cx:f>";
            const auto level = chartExCacheLevelXml(series.valuesCache(), true);
            if (!level.empty()) xml << level; else xml << "<cx:lvl ptCount=\"0\" formatCode=\"General\"/>";
            xml << "</cx:numDim>";
        }
        xml << "</cx:data>";
    }
    xml << "</cx:chartData><cx:chart>";
    if (!chart.title().empty()) {
        xml << "<cx:title pos=\"t\" align=\"ctr\" overlay=\"0\"><cx:tx><cx:txData><cx:v>"
            << xmlEscape(chart.title()) << "</cx:v></cx:txData></cx:tx></cx:title>";
    }
    xml << "<cx:plotArea><cx:plotAreaRegion>";
    for (std::size_t s = 0; s < chart.series().size(); ++s) {
        const auto& series = chart.series()[s];
        const auto writeSeries = [&](const char* layoutId, std::size_t ordinal) {
            xml << "<cx:series layoutId=\"" << layoutId << "\" uniqueId=\"{00000000-0000-0000-0000-"
                << std::setw(12) << std::setfill('0') << (s * 2 + ordinal + 1) << "}\">";
            if (!series.title().empty()) xml << "<cx:tx><cx:txData><cx:v>" << xmlEscape(series.title()) << "</cx:v></cx:txData></cx:tx>";
            xml << "<cx:dataId val=\"" << s << "\"/>";
            if (ordinal == 0) xml << chartExLayoutPropertiesXml(chart);
            xml << "</cx:series>";
        };
        writeSeries(chartExLayoutId(chart.type()), 0);
        if (chart.type() == xlpp::Chart::Type::Pareto) writeSeries("paretoLine", 1);
    }
    xml << "</cx:plotAreaRegion></cx:plotArea>";
    if (chart.showLegend()) xml << "<cx:legend pos=\"" << xmlEscape(chart.legendPosition()) << "\" align=\"ctr\" overlay=\"0\"/>";
    xml << "</cx:chart></cx:chartSpace>";
    return xml.str();
}


namespace {

bool comboScatterLike(xlpp::Chart::Type type) {
    return type == xlpp::Chart::Type::Scatter || type == xlpp::Chart::Type::Bubble;
}

bool comboBarLike(xlpp::Chart::Type type) {
    return type == xlpp::Chart::Type::Bar || type == xlpp::Chart::Type::HorizontalBar;
}

bool comboLineLike(xlpp::Chart::Type type) {
    return type == xlpp::Chart::Type::Line || type == xlpp::Chart::Type::Area;
}

bool comboSupportedType(xlpp::Chart::Type type) {
    return comboBarLike(type) || comboLineLike(type) || comboScatterLike(type) || type == xlpp::Chart::Type::Stock;
}

std::pair<std::uint64_t, std::uint64_t> comboAxisPair(const xlpp::Chart::Plot& plot) {
    if (plot.axisIds.size() >= 2) return {plot.axisIds[0], plot.axisIds[1]};
    return plot.usesSecondaryAxes ? std::pair<std::uint64_t, std::uint64_t>{3, 4}
                                  : std::pair<std::uint64_t, std::uint64_t>{1, 2};
}

std::string comboSeriesXml(const xlpp::Chart& chart, const xlpp::Chart::Plot& plot,
                           xlpp::Chart::Type type, std::size_t seriesIndex) {
    if (seriesIndex >= chart.series().size()) throw std::invalid_argument("Combined-chart plot references a series outside the chart");
    const auto& series = chart.series()[seriesIndex];
    const bool scatter = type == xlpp::Chart::Type::Scatter;
    const bool bubble = type == xlpp::Chart::Type::Bubble;
    std::ostringstream xml;
    xml << "<c:ser><c:idx val=\"" << seriesIndex << "\"/><c:order val=\"" << seriesIndex << "\"/>";
    if (!series.titleReference().empty()) {
        xml << "<c:tx><c:strRef><c:f>" << xmlEscape(series.titleReference()) << "</c:f>"
            << chartSeriesCacheXml(series.titleCache()) << "</c:strRef></c:tx>";
    } else if (!series.title().empty()) {
        xml << "<c:tx><c:v>" << xmlEscape(series.title()) << "</c:v></c:tx>";
    }
    if (scatter || bubble) {
        if (!series.categoriesReference().empty())
            xml << "<c:xVal><c:numRef><c:f>" << xmlEscape(series.categoriesReference()) << "</c:f>"
                << chartSeriesCacheXml(series.categoriesCache()) << "</c:numRef></c:xVal>";
        if (!series.valuesReference().empty())
            xml << "<c:yVal><c:numRef><c:f>" << xmlEscape(series.valuesReference()) << "</c:f>"
                << chartSeriesCacheXml(series.valuesCache()) << "</c:numRef></c:yVal>";
        if (bubble) {
            if (!series.bubbleSizeReference().empty())
                xml << "<c:bubbleSize><c:numRef><c:f>" << xmlEscape(series.bubbleSizeReference()) << "</c:f>"
                    << chartSeriesCacheXml(series.bubbleSizeCache()) << "</c:numRef></c:bubbleSize>";
            xml << "<c:bubble3D val=\"" << (plot.bubble3D ? 1 : 0) << "\"/>";
        }
    } else {
        if (!series.categoriesReference().empty()) {
            const bool numericCategories = type == xlpp::Chart::Type::Stock ||
                (series.categoriesCache().present && series.categoriesCache().numeric);
            const auto refTag = numericCategories ? "numRef" : "strRef";
            xml << "<c:cat><c:" << refTag << "><c:f>" << xmlEscape(series.categoriesReference()) << "</c:f>"
                << chartSeriesCacheXml(series.categoriesCache()) << "</c:" << refTag << "></c:cat>";
        }
        if (!series.valuesReference().empty())
            xml << "<c:val><c:numRef><c:f>" << xmlEscape(series.valuesReference()) << "</c:f>"
                << chartSeriesCacheXml(series.valuesCache()) << "</c:numRef></c:val>";
    }
    if ((type == xlpp::Chart::Type::Line || scatter) && series.hasSmooth())
        xml << "<c:smooth val=\"" << (series.smooth() ? 1 : 0) << "\"/>";
    xml << "</c:ser>";
    return xml.str();
}

std::string comboPlotXml(const xlpp::Chart& chart, const xlpp::Chart::Plot& plot, bool strict) {
    const auto type = plot.type;
    if (!comboSupportedType(type)) throw std::invalid_argument("This chart type cannot participate in a generated Excel combo chart");
    const auto count = plot.seriesCount == 0 ? chart.series().size() - std::min(plot.firstSeries, chart.series().size()) : plot.seriesCount;
    if (plot.firstSeries > chart.series().size() || count > chart.series().size() - plot.firstSeries)
        throw std::invalid_argument("Combined-chart plot series range is invalid");
    if (type == xlpp::Chart::Type::Stock && count != 3 && count != 4)
        throw std::invalid_argument("A stock plot inside a combo chart requires 3 (HLC) or 4 (OHLC) series");

    const auto axes = comboAxisPair(plot);
    std::ostringstream xml;
    xml << "<c:" << xlpp::Chart::typeName(type, plot.grouping) << ">";
    if (comboBarLike(type)) {
        const bool horizontal = type == xlpp::Chart::Type::HorizontalBar || plot.barDirection == xlpp::Chart::BarDirection::Bar;
        xml << "<c:barDir val=\"" << (horizontal ? "bar" : "col") << "\"/>";
        const char* grouping = plot.grouping == xlpp::Chart::Grouping::Stacked ? "stacked"
            : (plot.grouping == xlpp::Chart::Grouping::PercentStacked ? "percentStacked" : "clustered");
        xml << "<c:grouping val=\"" << grouping << "\"/><c:varyColors val=\"0\"/>";
    } else if (comboLineLike(type)) {
        const char* grouping = plot.grouping == xlpp::Chart::Grouping::Stacked ? "stacked"
            : (plot.grouping == xlpp::Chart::Grouping::PercentStacked ? "percentStacked" : "standard");
        xml << "<c:grouping val=\"" << grouping << "\"/>";
        if (type == xlpp::Chart::Type::Line) xml << "<c:varyColors val=\"0\"/>";
    } else if (type == xlpp::Chart::Type::Scatter) {
        const char* value = "marker";
        switch (plot.scatterStyle) {
            case xlpp::Chart::ScatterStyle::None: value = "none"; break;
            case xlpp::Chart::ScatterStyle::Line: value = "line"; break;
            case xlpp::Chart::ScatterStyle::LineMarker: value = "lineMarker"; break;
            case xlpp::Chart::ScatterStyle::Marker: value = "marker"; break;
            case xlpp::Chart::ScatterStyle::Smooth: value = "smooth"; break;
            case xlpp::Chart::ScatterStyle::SmoothMarker: value = "smoothMarker"; break;
        }
        xml << "<c:scatterStyle val=\"" << value << "\"/><c:varyColors val=\"0\"/>";
    } else if (type == xlpp::Chart::Type::Bubble) {
        xml << "<c:varyColors val=\"0\"/>";
    }
    for (std::size_t i = 0; i < count; ++i) xml << comboSeriesXml(chart, plot, type, plot.firstSeries + i);
    xml << generatedPlotAuxiliaryXml(plot, strict);
    if (type == xlpp::Chart::Type::Bubble) {
        if (plot.hasBubbleScale) xml << "<c:bubbleScale val=\"" << std::clamp(plot.bubbleScale, 0, 300) << "\"/>";
        xml << "<c:showNegBubbles val=\"" << (plot.showNegativeBubbles ? 1 : 0) << "\"/>"
            << "<c:sizeRepresents val=\"" << (plot.bubbleSizeRepresents == xlpp::Chart::BubbleSizeRepresents::Width ? "w" : "area") << "\"/>";
    }
    xml << "<c:axId val=\"" << axes.first << "\"/><c:axId val=\"" << axes.second << "\"/>";
    xml << "</c:" << xlpp::Chart::typeName(type, plot.grouping) << ">";
    return xml.str();
}

void comboAxesXml(std::ostringstream& xml, std::uint64_t xId, std::uint64_t yId,
                  bool scatterLike, bool secondary, bool horizontal,
                  const xlpp::Chart& chart) {
    const char* xPos = secondary ? "t" : (horizontal ? "l" : "b");
    const char* yPos = secondary ? "r" : (horizontal ? "b" : "l");
    if (scatterLike) {
        xml << "<c:valAx><c:axId val=\"" << xId << "\"/><c:scaling><c:orientation val=\"minMax\"/></c:scaling><c:delete val=\"0\"/><c:axPos val=\"" << xPos << "\"/>";
        if (!secondary && !chart.xAxisTitle().empty()) xml << "<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>" << xmlEscape(chart.xAxisTitle()) << "</a:t></a:r></a:p></c:rich></c:tx><c:overlay val=\"0\"/></c:title>";
        xml << "<c:numFmt formatCode=\"General\" sourceLinked=\"1\"/><c:majorTickMark val=\"none\"/><c:minorTickMark val=\"none\"/><c:tickLblPos val=\"nextTo\"/><c:crossAx val=\"" << yId << "\"/><c:crosses val=\"autoZero\"/><c:crossBetween val=\"midCat\"/></c:valAx>";
    } else {
        xml << "<c:catAx><c:axId val=\"" << xId << "\"/><c:scaling><c:orientation val=\"minMax\"/></c:scaling><c:delete val=\"0\"/><c:axPos val=\"" << xPos << "\"/>";
        if (!secondary && !chart.xAxisTitle().empty()) xml << "<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>" << xmlEscape(chart.xAxisTitle()) << "</a:t></a:r></a:p></c:rich></c:tx><c:overlay val=\"0\"/></c:title>";
        xml << "<c:numFmt formatCode=\"General\" sourceLinked=\"1\"/><c:majorTickMark val=\"none\"/><c:minorTickMark val=\"none\"/><c:tickLblPos val=\"nextTo\"/><c:crossAx val=\"" << yId << "\"/><c:crosses val=\"autoZero\"/><c:auto val=\"1\"/><c:lblAlgn val=\"ctr\"/><c:lblOffset val=\"100\"/></c:catAx>";
    }
    xml << "<c:valAx><c:axId val=\"" << yId << "\"/><c:scaling><c:orientation val=\"minMax\"/></c:scaling><c:delete val=\"0\"/><c:axPos val=\"" << yPos << "\"/>";
    if (!secondary && !chart.yAxisTitle().empty()) xml << "<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>" << xmlEscape(chart.yAxisTitle()) << "</a:t></a:r></a:p></c:rich></c:tx><c:overlay val=\"0\"/></c:title>";
    xml << "<c:numFmt formatCode=\"General\" sourceLinked=\"1\"/><c:majorGridlines/><c:majorTickMark val=\"none\"/><c:minorTickMark val=\"none\"/><c:tickLblPos val=\"nextTo\"/><c:crossAx val=\"" << xId << "\"/><c:crosses val=\"autoZero\"/><c:crossBetween val=\"" << (scatterLike ? "midCat" : "between") << "\"/></c:valAx>";
}

std::string serializeCombinedChart(const xlpp::Chart& chart, bool strict) {
    if (chart.plots().size() < 2) throw std::invalid_argument("Combined chart requires at least two plots");
    bool primarySeen = false, secondarySeen = false;
    bool primaryScatter = false, secondaryScatter = false;
    bool primaryHorizontal = false, secondaryHorizontal = false;
    std::pair<std::uint64_t, std::uint64_t> primaryAxes{1,2}, secondaryAxes{3,4};
    for (const auto& plot : chart.plots()) {
        if (!comboSupportedType(plot.type)) throw std::invalid_argument("Unsupported generated Excel combo-chart plot type");
        const bool scatter = comboScatterLike(plot.type);
        const bool horizontal = plot.type == xlpp::Chart::Type::HorizontalBar ||
            (comboBarLike(plot.type) && plot.barDirection == xlpp::Chart::BarDirection::Bar);
        auto& seen = plot.usesSecondaryAxes ? secondarySeen : primarySeen;
        auto& family = plot.usesSecondaryAxes ? secondaryScatter : primaryScatter;
        auto& horiz = plot.usesSecondaryAxes ? secondaryHorizontal : primaryHorizontal;
        auto& axes = plot.usesSecondaryAxes ? secondaryAxes : primaryAxes;
        if (!seen) { seen = true; family = scatter; horiz = horizontal; axes = comboAxisPair(plot); }
        else {
            if (family != scatter) throw std::invalid_argument("Scatter/bubble plots must use a separate axis group from category-based combo plots");
            if (horiz != horizontal && (horizontal || horiz)) throw std::invalid_argument("Horizontal bar plots cannot share a combo axis group with vertical plots");
        }
    }
    if (!primarySeen) throw std::invalid_argument("Combined chart must contain at least one primary-axis plot");

    std::ostringstream xml;
    xml << "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>";
    xml << "<c:chartSpace xmlns:c=\"" << (strict ? "http://purl.oclc.org/ooxml/drawingml/chart" : "http://schemas.openxmlformats.org/drawingml/2006/chart") << "\""
        << " xmlns:a=\"" << (strict ? "http://purl.oclc.org/ooxml/drawingml/main" : "http://schemas.openxmlformats.org/drawingml/2006/main") << "\">";
    if (!chart.style().empty()) xml << "<c:style val=\"" << xmlEscape(chart.style()) << "\"/>";
    xml << "<c:chart>";
    if (!chart.title().empty()) xml << "<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang=\"en-US\"/><a:t>" << xmlEscape(chart.title()) << "</a:t></a:r></a:p></c:rich></c:tx><c:overlay val=\"0\"/></c:title>";
    xml << "<c:plotArea><c:layout/>";
    for (const auto& plot : chart.plots()) xml << comboPlotXml(chart, plot, strict);
    comboAxesXml(xml, primaryAxes.first, primaryAxes.second, primaryScatter, false, primaryHorizontal, chart);
    if (secondarySeen) comboAxesXml(xml, secondaryAxes.first, secondaryAxes.second, secondaryScatter, true, secondaryHorizontal, chart);
    if (chart.dataTable().present) xml << generatedDataTableXml(chart.dataTable(), strict);
    xml << "</c:plotArea>";
    if (chart.showLegend()) xml << "<c:legend><c:legendPos val=\"" << xmlEscape(chart.legendPosition()) << "\"/></c:legend>";
    xml << "</c:chart></c:chartSpace>";
    return xml.str();
}

} // namespace

std::string serializeChart(const xlpp::Chart& chart, bool strict) {
    if (chart.combined()) return serializeCombinedChart(chart, strict);
    if (chart.modern()) return serializeChartEx(chart, strict);
    std::ostringstream xml;
    const auto type = chart.type();
    const auto* plot = chart.primaryPlotOrNull();
    const bool scatter = type == xlpp::Chart::Type::Scatter;
    const bool bubble = type == xlpp::Chart::Type::Bubble;
    const bool threeAxis = type == xlpp::Chart::Type::Bar3D || type == xlpp::Chart::Type::Line3D ||
                           type == xlpp::Chart::Type::Area3D || type == xlpp::Chart::Type::Surface ||
                           type == xlpp::Chart::Type::Surface3D;
    const bool projectedPie = type == xlpp::Chart::Type::PieOfPie || type == xlpp::Chart::Type::BarOfPie;
    const bool hasAxes = type != xlpp::Chart::Type::Pie && type != xlpp::Chart::Type::Pie3D && type != xlpp::Chart::Type::Doughnut && !projectedPie;
    const bool barLike = type == xlpp::Chart::Type::Bar || type == xlpp::Chart::Type::HorizontalBar || type == xlpp::Chart::Type::Bar3D;
    const bool lineLike = type == xlpp::Chart::Type::Line || type == xlpp::Chart::Type::Area || type == xlpp::Chart::Type::Line3D || type == xlpp::Chart::Type::Area3D;
    const bool surface = type == xlpp::Chart::Type::Surface || type == xlpp::Chart::Type::Surface3D;

    xml << "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>";
    xml << "<c:chartSpace xmlns:c=\"" << (strict ? "http://purl.oclc.org/ooxml/drawingml/chart" : "http://schemas.openxmlformats.org/drawingml/2006/chart") << "\"";
    xml << " xmlns:a=\"" << (strict ? "http://purl.oclc.org/ooxml/drawingml/main" : "http://schemas.openxmlformats.org/drawingml/2006/main") << "\">";
    if (!chart.style().empty()) xml << "<c:style val=\"" << xmlEscape(chart.style()) << "\"/>";
    xml << "<c:chart>";
    if (!chart.title().empty())
        xml << "<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang=\"en-US\"/><a:t>" << xmlEscape(chart.title()) << "</a:t></a:r></a:p></c:rich></c:tx><c:overlay val=\"0\"/></c:title>";
    if (chart.view3D().present) xml << chartView3DXml(chart.view3D(), true);
    if (chart.floorFormat().present) xml << generatedChartWallXml("floor", chart.floorFormat(), true);
    if (chart.sideWallFormat().present) xml << generatedChartWallXml("sideWall", chart.sideWallFormat(), true);
    if (chart.backWallFormat().present) xml << generatedChartWallXml("backWall", chart.backWallFormat(), true);
    if (type == xlpp::Chart::Type::Stock && chart.series().size() != 3 && chart.series().size() != 4)
        throw std::invalid_argument("Stock charts require exactly 3 (high-low-close) or 4 (open-high-low-close) series");

    xml << "<c:plotArea><c:layout/><c:" << xlpp::Chart::typeName(type, chart.grouping()) << ">";
    if (barLike) {
        const bool horizontal = type == xlpp::Chart::Type::HorizontalBar || (plot && plot->barDirection == xlpp::Chart::BarDirection::Bar);
        xml << "<c:barDir val=\"" << (horizontal ? "bar" : "col") << "\"/>";
        const char* grouping = chart.grouping() == xlpp::Chart::Grouping::Stacked ? "stacked"
            : (chart.grouping() == xlpp::Chart::Grouping::PercentStacked ? "percentStacked" : "clustered");
        xml << "<c:grouping val=\"" << grouping << "\"/><c:varyColors val=\"0\"/>";
    } else if (lineLike) {
        const char* grouping = chart.grouping() == xlpp::Chart::Grouping::Stacked ? "stacked"
            : (chart.grouping() == xlpp::Chart::Grouping::PercentStacked ? "percentStacked" : "standard");
        xml << "<c:grouping val=\"" << grouping << "\"/>";
        if (type == xlpp::Chart::Type::Line || type == xlpp::Chart::Type::Line3D) xml << "<c:varyColors val=\"0\"/>";
    } else if (scatter) {
        const auto style = plot ? plot->scatterStyle : xlpp::Chart::ScatterStyle::Marker;
        const char* value = "marker";
        switch (style) {
            case xlpp::Chart::ScatterStyle::None: value = "none"; break;
            case xlpp::Chart::ScatterStyle::Line: value = "line"; break;
            case xlpp::Chart::ScatterStyle::LineMarker: value = "lineMarker"; break;
            case xlpp::Chart::ScatterStyle::Marker: value = "marker"; break;
            case xlpp::Chart::ScatterStyle::Smooth: value = "smooth"; break;
            case xlpp::Chart::ScatterStyle::SmoothMarker: value = "smoothMarker"; break;
        }
        xml << "<c:scatterStyle val=\"" << value << "\"/><c:varyColors val=\"0\"/>";
    } else if (bubble) {
        xml << "<c:varyColors val=\"0\"/>";
    } else if (type == xlpp::Chart::Type::Pie3D || type == xlpp::Chart::Type::Pie || type == xlpp::Chart::Type::Doughnut || projectedPie) {
        xml << "<c:varyColors val=\"1\"/>";
    } else if (type == xlpp::Chart::Type::Radar) {
        xml << "<c:radarStyle val=\"" << xmlEscape(plot && !plot->radarStyle.empty() ? plot->radarStyle : "standard") << "\"/>";
    } else if (surface && plot && plot->hasWireframe) {
        xml << "<c:wireframe val=\"" << (plot->wireframe ? "1" : "0") << "\"/>";
    }

    for (std::size_t s = 0; s < chart.series().size(); ++s) {
        const auto& series = chart.series()[s];
        std::ostringstream seriesXml;
        seriesXml << "<c:ser><c:idx val=\"" << s << "\"/><c:order val=\"" << s << "\"/>";
        if (!series.titleReference().empty()) {
            seriesXml << "<c:tx><c:strRef><c:f>" << xmlEscape(series.titleReference()) << "</c:f>"
                      << chartSeriesCacheXml(series.titleCache()) << "</c:strRef></c:tx>";
        } else if (!series.title().empty()) seriesXml << "<c:tx><c:v>" << xmlEscape(series.title()) << "</c:v></c:tx>";
        if (type == xlpp::Chart::Type::Radar && series.markerFormat().present) {
            std::string markerOwner = "<c:owner></c:owner>";
            if (!patchMarkerFormatInOwner(markerOwner, series.markerFormat())) throw std::runtime_error("Failed to serialize radar marker format");
            const auto markerNodes = drawingTags(markerOwner, "c:marker", "marker");
            if (!markerNodes.empty()) seriesXml << markerNodes.front();
        }
        if (scatter || bubble) {
            if (!series.categoriesReference().empty())
                seriesXml << "<c:xVal><c:numRef><c:f>" << xmlEscape(series.categoriesReference()) << "</c:f>" << chartSeriesCacheXml(series.categoriesCache()) << "</c:numRef></c:xVal>";
            if (!series.valuesReference().empty())
                seriesXml << "<c:yVal><c:numRef><c:f>" << xmlEscape(series.valuesReference()) << "</c:f>" << chartSeriesCacheXml(series.valuesCache()) << "</c:numRef></c:yVal>";
            if (bubble) {
                if (!series.bubbleSizeReference().empty())
                    seriesXml << "<c:bubbleSize><c:numRef><c:f>" << xmlEscape(series.bubbleSizeReference()) << "</c:f>" << chartSeriesCacheXml(series.bubbleSizeCache()) << "</c:numRef></c:bubbleSize>";
                seriesXml << "<c:bubble3D val=\"" << ((plot && plot->bubble3D) ? 1 : 0) << "\"/>";
            }
        } else {
            if (!series.categoriesReference().empty()) {
                const bool numericCategories = type == xlpp::Chart::Type::Stock || (series.categoriesCache().present && series.categoriesCache().numeric);
                const auto refTag = numericCategories ? "numRef" : "strRef";
                seriesXml << "<c:cat><c:" << refTag << "><c:f>" << xmlEscape(series.categoriesReference()) << "</c:f>"
                          << chartSeriesCacheXml(series.categoriesCache()) << "</c:" << refTag << "></c:cat>";
            }
            if (!series.valuesReference().empty()) seriesXml << "<c:val><c:numRef><c:f>" << xmlEscape(series.valuesReference()) << "</c:f>" << chartSeriesCacheXml(series.valuesCache()) << "</c:numRef></c:val>";
        }
        if ((type == xlpp::Chart::Type::Line || scatter) && series.hasSmooth()) seriesXml << "<c:smooth val=\"" << (series.smooth() ? 1 : 0) << "\"/>";
        seriesXml << "</c:ser>";
        xml << seriesXml.str();
    }
    if (plot) {
        xml << generatedPlotAuxiliaryXml(*plot, strict);
        if (plot->hasGapDepth && (type == xlpp::Chart::Type::Bar3D || type == xlpp::Chart::Type::Line3D || type == xlpp::Chart::Type::Area3D)) xml << "<c:gapDepth val=\"" << plot->gapDepth << "\"/>";
        if (!plot->shape.empty() && type == xlpp::Chart::Type::Bar3D) xml << "<c:shape val=\"" << xmlEscape(plot->shape) << "\"/>";
        if ((type == xlpp::Chart::Type::Pie || type == xlpp::Chart::Type::Doughnut) && plot->hasFirstSliceAngle) {
            if (plot->firstSliceAngle < 0 || plot->firstSliceAngle > 360) throw std::invalid_argument("Chart first-slice angle must be between 0 and 360 degrees");
            xml << "<c:firstSliceAng val=\"" << plot->firstSliceAngle << "\"/>";
        }
        if (type == xlpp::Chart::Type::Doughnut && plot->hasHoleSize) {
            if (plot->holeSize < 10 || plot->holeSize > 90) throw std::invalid_argument("Doughnut hole size must be between 10 and 90 percent");
            xml << "<c:holeSize val=\"" << plot->holeSize << "\"/>";
        }
        if (bubble) {
            if (plot->hasBubbleScale) xml << "<c:bubbleScale val=\"" << std::clamp(plot->bubbleScale, 0, 300) << "\"/>";
            xml << "<c:showNegBubbles val=\"" << (plot->showNegativeBubbles ? 1 : 0) << "\"/>"
                << "<c:sizeRepresents val=\"" << (plot->bubbleSizeRepresents == xlpp::Chart::BubbleSizeRepresents::Width ? "w" : "area") << "\"/>";
        }
        if (projectedPie) {
            auto options = plot->projectedPie; options.present = true; options.ofPieType = type == xlpp::Chart::Type::BarOfPie ? "bar" : "pie";
            if (options.gapWidth < 0 || options.gapWidth > 500 || options.secondPlotSize < 5 || options.secondPlotSize > 200) throw std::invalid_argument("Projected-pie gap width or second-plot size is outside the supported OOXML range");
            xml << "<c:ofPieType val=\"" << options.ofPieType << "\"/><c:gapWidth val=\"" << options.gapWidth << "\"/><c:splitType val=\"" << options.splitType << "\"/>";
            if (options.hasSplitPosition) xml << "<c:splitPos val=\"" << options.splitPosition << "\"/>";
            if (!options.customSplitPoints.empty()) { xml << "<c:custSplit>"; for (const auto point : options.customSplitPoints) xml << "<c:secondPiePt val=\"" << point << "\"/>"; xml << "</c:custSplit>"; }
            xml << "<c:secondPieSize val=\"" << options.secondPlotSize << "\"/>";
        }
    } else if (projectedPie) {
        xml << "<c:ofPieType val=\"" << (type == xlpp::Chart::Type::BarOfPie ? "bar" : "pie") << "\"/><c:gapWidth val=\"150\"/><c:splitType val=\"auto\"/><c:secondPieSize val=\"75\"/>";
    }
    if (hasAxes) {
        xml << "<c:axId val=\"1\"/><c:axId val=\"2\"/>";
        if (threeAxis) xml << "<c:axId val=\"3\"/>";
    }
    if ((type == xlpp::Chart::Type::Pie || type == xlpp::Chart::Type::Doughnut) && (!plot || !plot->hasFirstSliceAngle)) xml << "<c:firstSliceAng val=\"0\"/>";
    if (type == xlpp::Chart::Type::Doughnut && (!plot || !plot->hasHoleSize)) xml << "<c:holeSize val=\"10\"/>";
    xml << "</c:" << xlpp::Chart::typeName(type, chart.grouping()) << ">";

    if (hasAxes) {
        const bool horizontal = type == xlpp::Chart::Type::HorizontalBar || (barLike && plot && plot->barDirection == xlpp::Chart::BarDirection::Bar);
        if (scatter || bubble) {
            xml << "<c:valAx><c:axId val=\"1\"/><c:scaling><c:orientation val=\"minMax\"/></c:scaling><c:delete val=\"0\"/><c:axPos val=\"b\"/>";
            if (!chart.xAxisTitle().empty()) xml << "<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>" << xmlEscape(chart.xAxisTitle()) << "</a:t></a:r></a:p></c:rich></c:tx><c:overlay val=\"0\"/></c:title>";
            xml << "<c:numFmt formatCode=\"General\" sourceLinked=\"1\"/><c:majorTickMark val=\"none\"/><c:minorTickMark val=\"none\"/><c:tickLblPos val=\"nextTo\"/><c:crossAx val=\"2\"/><c:crosses val=\"autoZero\"/><c:crossBetween val=\"midCat\"/></c:valAx>";
            xml << "<c:valAx><c:axId val=\"2\"/><c:scaling><c:orientation val=\"minMax\"/></c:scaling><c:delete val=\"0\"/><c:axPos val=\"l\"/>";
            if (!chart.yAxisTitle().empty()) xml << "<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>" << xmlEscape(chart.yAxisTitle()) << "</a:t></a:r></a:p></c:rich></c:tx><c:overlay val=\"0\"/></c:title>";
            xml << "<c:numFmt formatCode=\"General\" sourceLinked=\"1\"/><c:majorGridlines/><c:majorTickMark val=\"none\"/><c:minorTickMark val=\"none\"/><c:tickLblPos val=\"nextTo\"/><c:crossAx val=\"1\"/><c:crosses val=\"autoZero\"/><c:crossBetween val=\"midCat\"/></c:valAx>";
        } else {
            xml << "<c:catAx><c:axId val=\"1\"/><c:scaling><c:orientation val=\"minMax\"/></c:scaling><c:delete val=\"0\"/><c:axPos val=\"" << (horizontal ? "l" : "b") << "\"/>";
            if (!chart.xAxisTitle().empty()) xml << "<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>" << xmlEscape(chart.xAxisTitle()) << "</a:t></a:r></a:p></c:rich></c:tx><c:overlay val=\"0\"/></c:title>";
            xml << "<c:numFmt formatCode=\"General\" sourceLinked=\"1\"/><c:majorTickMark val=\"none\"/><c:minorTickMark val=\"none\"/><c:tickLblPos val=\"nextTo\"/><c:crossAx val=\"2\"/><c:crosses val=\"autoZero\"/><c:auto val=\"1\"/><c:lblAlgn val=\"ctr\"/><c:lblOffset val=\"100\"/></c:catAx>";
            xml << "<c:valAx><c:axId val=\"2\"/><c:scaling><c:orientation val=\"minMax\"/></c:scaling><c:delete val=\"0\"/><c:axPos val=\"" << (horizontal ? "b" : "l") << "\"/>";
            if (!chart.yAxisTitle().empty()) xml << "<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>" << xmlEscape(chart.yAxisTitle()) << "</a:t></a:r></a:p></c:rich></c:tx><c:overlay val=\"0\"/></c:title>";
            xml << "<c:numFmt formatCode=\"General\" sourceLinked=\"1\"/><c:majorGridlines/><c:majorTickMark val=\"none\"/><c:minorTickMark val=\"none\"/><c:tickLblPos val=\"nextTo\"/><c:crossAx val=\"1\"/><c:crosses val=\"autoZero\"/><c:crossBetween val=\"between\"/></c:valAx>";
        }
        if (threeAxis) xml << "<c:serAx><c:axId val=\"3\"/><c:scaling><c:orientation val=\"minMax\"/></c:scaling><c:delete val=\"0\"/><c:axPos val=\"r\"/><c:majorTickMark val=\"none\"/><c:minorTickMark val=\"none\"/><c:crossAx val=\"1\"/><c:crosses val=\"autoZero\"/></c:serAx>";
    }
    if (chart.dataTable().present) xml << generatedDataTableXml(chart.dataTable(), strict);
    xml << "</c:plotArea>";
    if (chart.showLegend()) xml << "<c:legend><c:legendPos val=\"" << xmlEscape(chart.legendPosition()) << "\"/></c:legend>";
    xml << "</c:chart></c:chartSpace>";
    return xml.str();
}

std::string chartView3DXml(const xlpp::ChartView3D& view, bool prefixed) {
    const auto c = prefixed ? "c:" : "";
    std::ostringstream xml;
    xml << "<" << c << "view3D>";
    if (view.hasRotationX) xml << "<" << c << "rotX val=\"" << view.rotationX << "\"/>";
    if (view.hasHeightPercent) xml << "<" << c << "hPercent val=\"" << view.heightPercent << "\"/>";
    if (view.hasRotationY) xml << "<" << c << "rotY val=\"" << view.rotationY << "\"/>";
    if (view.hasDepthPercent) xml << "<" << c << "depthPercent val=\"" << view.depthPercent << "\"/>";
    if (view.hasRightAngleAxes) xml << "<" << c << "rAngAx val=\"" << (view.rightAngleAxes ? "1" : "0") << "\"/>";
    if (view.hasPerspective) xml << "<" << c << "perspective val=\"" << view.perspective << "\"/>";
    xml << "</" << c << "view3D>";
    return xml.str();
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

} // namespace xlpp::internal::ooxml
