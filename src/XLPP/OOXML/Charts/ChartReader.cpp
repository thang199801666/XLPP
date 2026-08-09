#include "OOXML/Charts/ChartReader.h"
#include "OOXML/Common/DrawingXmlSupport.h"
#include "OOXML/Charts/ChartXmlSupport.h"
#include "OOXML/Common/PackageRelationships.h"
#include "Package/Xml/XmlUtilities.h"
#include "Package/Zip/ZipArchive.h"
#include <XLPP/Worksheet/Worksheet.h>

#include <algorithm>
#include <array>
#include <cmath>
#include <cstddef>
#include <cstdint>
#include <string>
#include <utility>
#include <vector>

namespace xlpp::internal::ooxml {

struct ParsedChartPlotNode {
    std::size_t position{0};
    xlpp::Chart::Type type{xlpp::Chart::Type::Bar};
    std::string xml;
};

std::vector<ParsedChartPlotNode> chartPlotNodesInOrder(const std::string& chartXmlText) {
    const auto plotAreas = drawingTags(chartXmlText, "c:plotArea", "plotArea");
    if (plotAreas.empty()) return {};
    const auto& plotArea = plotAreas.front();
    std::vector<ParsedChartPlotNode> result;
    const auto collect = [&](const char* prefixed, const char* local, xlpp::Chart::Type type) {
        std::size_t cursor = 0;
        for (const auto& node : drawingTags(plotArea, prefixed, local)) {
            const auto position = plotArea.find(node, cursor);
            if (position == std::string::npos) continue;
            result.push_back({position, type, node});
            cursor = position + node.size();
        }
    };
    collect("c:barChart", "barChart", xlpp::Chart::Type::Bar);
    collect("c:lineChart", "lineChart", xlpp::Chart::Type::Line);
    collect("c:pieChart", "pieChart", xlpp::Chart::Type::Pie);
    collect("c:scatterChart", "scatterChart", xlpp::Chart::Type::Scatter);
    collect("c:doughnutChart", "doughnutChart", xlpp::Chart::Type::Doughnut);
    collect("c:radarChart", "radarChart", xlpp::Chart::Type::Radar);
    collect("c:areaChart", "areaChart", xlpp::Chart::Type::Area);
    collect("c:bubbleChart", "bubbleChart", xlpp::Chart::Type::Bubble);
    collect("c:stockChart", "stockChart", xlpp::Chart::Type::Stock);
    {
        std::size_t cursor = 0;
        for (const auto& node : drawingTags(plotArea, "c:ofPieChart", "ofPieChart")) {
            const auto position = plotArea.find(node, cursor);
            if (position == std::string::npos) continue;
            const auto kinds = drawingTags(node, "c:ofPieType", "ofPieType");
            const auto value = kinds.empty() ? std::string{"pie"} : xlpp::internal::attribute(kinds.front(), "val");
            result.push_back({position, value == "bar" ? xlpp::Chart::Type::BarOfPie : xlpp::Chart::Type::PieOfPie, node});
            cursor = position + node.size();
        }
    }
    collect("c:bar3DChart", "bar3DChart", xlpp::Chart::Type::Bar3D);
    collect("c:line3DChart", "line3DChart", xlpp::Chart::Type::Line3D);
    collect("c:area3DChart", "area3DChart", xlpp::Chart::Type::Area3D);
    collect("c:pie3DChart", "pie3DChart", xlpp::Chart::Type::Pie3D);
    collect("c:surfaceChart", "surfaceChart", xlpp::Chart::Type::Surface);
    collect("c:surface3DChart", "surface3DChart", xlpp::Chart::Type::Surface3D);
    std::sort(result.begin(), result.end(), [](const auto& a, const auto& b) { return a.position < b.position; });
    return result;
}

xlpp::Chart::Grouping parseChartGroupingNode(const std::string& chartNode) {
    const auto groupingNodes = drawingTags(chartNode, "c:grouping", "grouping");
    if (groupingNodes.empty()) return xlpp::Chart::Grouping::Standard;
    const auto value = xlpp::internal::attribute(groupingNodes.front(), "val");
    if (value == "stacked") return xlpp::Chart::Grouping::Stacked;
    if (value == "percentStacked") return xlpp::Chart::Grouping::PercentStacked;
    if (value == "clustered") return xlpp::Chart::Grouping::Clustered;
    return xlpp::Chart::Grouping::Standard;
}

std::vector<std::uint64_t> chartAxisIds(const std::string& chartNode) {
    std::vector<std::uint64_t> result;
    for (const auto& node : drawingTags(chartNode, "c:axId", "axId")) {
        const auto value = xlpp::internal::attribute(node, "val");
        if (value.empty()) continue;
        try { result.push_back(std::stoull(value)); } catch (...) {}
    }
    return result;
}

bool chartBoolValue(const std::string& container, const char* prefixed, const char* local, bool fallback = false) {
    const auto nodes = drawingTags(container, prefixed, local);
    if (nodes.empty()) return fallback;
    const auto value = xlpp::internal::attribute(nodes.front(), "val");
    return value == "1" || value == "true" || value == "True";
}

xlpp::ChartColor parseChartColor(const std::string& container) {
    using Kind = xlpp::ChartColor::Kind;
    using TransformKind = xlpp::ChartColorTransform::Kind;
    for (const auto& entry : std::array<std::tuple<const char*, const char*, Kind>, 4>{
             std::tuple{"a:srgbClr", "srgbClr", Kind::SRgb},
             std::tuple{"a:schemeClr", "schemeClr", Kind::Scheme},
             std::tuple{"a:sysClr", "sysClr", Kind::System},
             std::tuple{"a:prstClr", "prstClr", Kind::Preset}}) {
        const auto nodes = drawingTags(container, std::get<0>(entry), std::get<1>(entry));
        if (nodes.empty()) continue;
        const auto value = xlpp::internal::attribute(nodes.front(), "val");
        if (value.empty()) continue;
        xlpp::ChartColor color;
        color.kind = std::get<2>(entry);
        color.value = value;
        const auto& colorXml = nodes.front();
        for (const auto& transform : std::array<std::tuple<const char*, const char*, TransformKind>, 9>{
                 std::tuple{"a:alpha", "alpha", TransformKind::Alpha},
                 std::tuple{"a:alphaMod", "alphaMod", TransformKind::AlphaMod},
                 std::tuple{"a:alphaOff", "alphaOff", TransformKind::AlphaOff},
                 std::tuple{"a:tint", "tint", TransformKind::Tint},
                 std::tuple{"a:shade", "shade", TransformKind::Shade},
                 std::tuple{"a:lumMod", "lumMod", TransformKind::LumMod},
                 std::tuple{"a:lumOff", "lumOff", TransformKind::LumOff},
                 std::tuple{"a:satMod", "satMod", TransformKind::SatMod},
                 std::tuple{"a:satOff", "satOff", TransformKind::SatOff}}) {
            for (const auto& node : drawingTags(colorXml, std::get<0>(transform), std::get<1>(transform))) {
                try {
                    color.transforms.push_back({std::get<2>(transform), std::stoi(xlpp::internal::attribute(node, "val"))});
                } catch (...) {}
            }
        }
        return color;
    }
    return {};
}


xlpp::ChartSeriesCache parseChartSeriesCache(const std::string& container) {
    xlpp::ChartSeriesCache result;
    auto caches = drawingTags(container, "c:numCache", "numCache");
    if (!caches.empty()) result.numeric = true;
    else { caches = drawingTags(container, "c:strCache", "strCache"); result.numeric = false; }
    if (caches.empty()) return result;
    result.present = true;
    const auto& cache = caches.front();
    result.formatCode = drawingTagText(cache, "c:formatCode", "formatCode");
    const auto counts = drawingTags(cache, "c:ptCount", "ptCount");
    if (!counts.empty()) { try { result.pointCount = static_cast<std::size_t>(std::stoull(xlpp::internal::attribute(counts.front(), "val"))); } catch (...) {} }
    for (const auto& point : drawingTags(cache, "c:pt", "pt")) {
        const auto idx = xlpp::internal::attribute(point, "idx");
        if (idx.empty()) continue;
        try { result.points.push_back({static_cast<std::size_t>(std::stoull(idx)), drawingTagText(point, "c:v", "v")}); } catch (...) {}
    }
    return result;
}

xlpp::ChartThemePalette parseChartThemePalette(const xlpp::internal::ZipArchive& z) {
    xlpp::ChartThemePalette palette;
    if (!z.contains("xl/theme/theme1.xml")) return palette;
    const auto xml = z.get("xl/theme/theme1.xml");
    for (const auto& name : std::array<const char*, 12>{"dk1","lt1","dk2","lt2","accent1","accent2","accent3","accent4","accent5","accent6","hlink","folHlink"}) {
        const auto prefixed = std::string("a:") + name;
        const auto nodes = drawingTags(xml, prefixed.c_str(), name);
        if (nodes.empty()) continue;
        std::string value;
        const auto srgb = drawingTags(nodes.front(), "a:srgbClr", "srgbClr");
        if (!srgb.empty()) value = xlpp::internal::attribute(srgb.front(), "val");
        if (value.empty()) {
            const auto sys = drawingTags(nodes.front(), "a:sysClr", "sysClr");
            if (!sys.empty()) { value = xlpp::internal::attribute(sys.front(), "lastClr"); if (value.empty()) value = xlpp::internal::attribute(sys.front(), "val"); }
        }
        if (!value.empty()) palette.colors.push_back({name, value});
    }
    const auto fontSchemes = drawingTags(xml, "a:fontScheme", "fontScheme");
    if (!fontSchemes.empty()) {
        palette.fontScheme.present = true;
        palette.fontScheme.name = xlpp::internal::attribute(fontSchemes.front(), "name");
        const auto major = drawingTags(fontSchemes.front(), "a:majorFont", "majorFont");
        if (!major.empty()) {
            const auto latin = drawingTags(major.front(), "a:latin", "latin");
            if (!latin.empty()) palette.fontScheme.majorLatinTypeface = xlpp::internal::attribute(latin.front(), "typeface");
        }
        const auto minor = drawingTags(fontSchemes.front(), "a:minorFont", "minorFont");
        if (!minor.empty()) {
            const auto latin = drawingTags(minor.front(), "a:latin", "latin");
            if (!latin.empty()) palette.fontScheme.minorLatinTypeface = xlpp::internal::attribute(latin.front(), "typeface");
        }
    }
    const auto fmtSchemes = drawingTags(xml, "a:fmtScheme", "fmtScheme");
    if (!fmtSchemes.empty()) {
        auto& effects = palette.effectScheme; effects.present = true; effects.name = xlpp::internal::attribute(fmtSchemes.front(), "name");
        const auto countFills = [&](const std::string& list) {
            std::size_t count = 0;
            for (const auto* tag : {"solidFill", "gradFill", "pattFill", "noFill"}) count += drawingTags(list, (std::string("a:") + tag).c_str(), tag).size();
            return count;
        };
        const auto fills = drawingTags(fmtSchemes.front(), "a:fillStyleLst", "fillStyleLst"); if (!fills.empty()) effects.fillStyleCount = countFills(fills.front());
        const auto lines = drawingTags(fmtSchemes.front(), "a:lnStyleLst", "lnStyleLst"); if (!lines.empty()) effects.lineStyleCount = drawingTags(lines.front(), "a:ln", "ln").size();
        const auto effectStyles = drawingTags(fmtSchemes.front(), "a:effectStyleLst", "effectStyleLst"); if (!effectStyles.empty()) effects.effectStyleCount = drawingTags(effectStyles.front(), "a:effectStyle", "effectStyle").size();
        const auto bgFills = drawingTags(fmtSchemes.front(), "a:bgFillStyleLst", "bgFillStyleLst"); if (!bgFills.empty()) effects.backgroundFillStyleCount = countFills(bgFills.front());
    }
    palette.present = !palette.colors.empty() || palette.fontScheme.present || palette.effectScheme.present;
    return palette;
}

xlpp::ChartStyleResources parseChartStyleResources(const xlpp::internal::ZipArchive& z, const std::string& chartPart) {
    xlpp::ChartStyleResources resources;
    const auto slash = chartPart.find_last_of('/');
    if (slash == std::string::npos) return resources;
    const auto file = chartPart.substr(slash + 1);
    const auto relsPart = chartPart.substr(0, slash + 1) + "_rels/" + file + ".rels";
    if (!z.contains(relsPart)) return resources;
    for (const auto& rel : xlpp::internal::tags(z.get(relsPart), "Relationship")) {
        const auto type = xlpp::internal::attribute(rel, "Type");
        const auto target = xlpp::internal::attribute(rel, "Target");
        if (target.empty()) continue;
        if (type.find("/chartStyle") != std::string::npos) { resources.chartStylePresent = true; resources.chartStylePart = resolvePackagePart(chartPart, target); }
        else if (type.find("/chartColorStyle") != std::string::npos) { resources.colorStylePresent = true; resources.colorStylePart = resolvePackagePart(chartPart, target); }
    }
    return resources;
}

xlpp::ChartLineFormat parseChartLineFormat(const std::string& container) {
    xlpp::ChartLineFormat format;
    const auto lines = drawingTags(container, "a:ln", "ln");
    if (lines.empty()) return format;
    const auto& line = lines.front();
    format.present = true;
    const auto width = xlpp::internal::attribute(line, "w");
    if (!width.empty()) {
        try { format.widthPoints = std::stod(width) / 12700.0; } catch (...) {}
    }
    format.cap = xlpp::internal::attribute(line, "cap");
    format.compound = xlpp::internal::attribute(line, "cmpd");
    format.noFill = !drawingTags(line, "a:noFill", "noFill").empty();
    const auto fills = drawingTags(line, "a:solidFill", "solidFill");
    if (!fills.empty()) format.color = parseChartColor(fills.front());
    const auto dashes = drawingTags(line, "a:prstDash", "prstDash");
    if (!dashes.empty()) format.dash = xlpp::internal::attribute(dashes.front(), "val");
    const auto custom = drawingTags(line, "a:custDash", "custDash");
    if (!custom.empty()) {
        for (const auto& ds : drawingTags(custom.front(), "a:ds", "ds")) {
            try {
                format.customDash.push_back({
                    std::stod(xlpp::internal::attribute(ds, "d")) / 1000.0,
                    std::stod(xlpp::internal::attribute(ds, "sp")) / 1000.0
                });
            } catch (...) {}
        }
    }
    if (!drawingTags(line, "a:round", "round").empty()) format.join = "round";
    else if (!drawingTags(line, "a:bevel", "bevel").empty()) format.join = "bevel";
    else if (!drawingTags(line, "a:miter", "miter").empty()) format.join = "miter";
    return format;
}

xlpp::ChartFillFormat parseChartFillFormat(const std::string& container) {
    xlpp::ChartFillFormat format;
    auto scope = container;
    for (const auto& line : drawingTags(scope, "a:ln", "ln")) {
        const auto position = scope.find(line);
        if (position != std::string::npos) scope.erase(position, line.size());
    }
    if (!drawingTags(scope, "a:noFill", "noFill").empty()) {
        format.present = true;
        format.noFill = true;
        format.kind = xlpp::ChartFillFormat::Kind::NoFill;
        return format;
    }
    const auto fills = drawingTags(scope, "a:solidFill", "solidFill");
    if (!fills.empty()) {
        format.present = true;
        format.kind = xlpp::ChartFillFormat::Kind::Solid;
        format.color = parseChartColor(fills.front());
        return format;
    }
    const auto gradients = drawingTags(scope, "a:gradFill", "gradFill");
    if (!gradients.empty()) {
        format.present = true;
        format.kind = xlpp::ChartFillFormat::Kind::Gradient;
        for (const auto& gs : drawingTags(gradients.front(), "a:gs", "gs")) {
            xlpp::ChartGradientStop stop;
            try { stop.position = std::stoi(xlpp::internal::attribute(gs, "pos")); } catch (...) {}
            stop.color = parseChartColor(gs);
            if (stop.color.present()) format.gradientStops.push_back(std::move(stop));
        }
        const auto linear = drawingTags(gradients.front(), "a:lin", "lin");
        if (!linear.empty()) {
            try { format.gradientAngleDegrees = std::stod(xlpp::internal::attribute(linear.front(), "ang")) / 60000.0; } catch (...) {}
        }
        return format;
    }
    const auto patterns = drawingTags(scope, "a:pattFill", "pattFill");
    if (!patterns.empty()) {
        format.present = true;
        format.kind = xlpp::ChartFillFormat::Kind::Pattern;
        format.pattern = xlpp::internal::attribute(patterns.front(), "prst");
        const auto foreground = drawingTags(patterns.front(), "a:fgClr", "fgClr");
        if (!foreground.empty()) format.foregroundColor = parseChartColor(foreground.front());
        const auto background = drawingTags(patterns.front(), "a:bgClr", "bgClr");
        if (!background.empty()) format.backgroundColor = parseChartColor(background.front());
    }
    return format;
}

xlpp::ChartMarkerFormat parseChartMarkerFormat(const std::string& seriesXml) {
    xlpp::ChartMarkerFormat format;
    const auto markers = drawingTags(seriesXml, "c:marker", "marker");
    if (markers.empty()) return format;
    const auto& marker = markers.front();
    format.present = true;
    const auto symbols = drawingTags(marker, "c:symbol", "symbol");
    if (!symbols.empty()) format.symbol = xlpp::internal::attribute(symbols.front(), "val");
    const auto sizes = drawingTags(marker, "c:size", "size");
    if (!sizes.empty()) { try { format.size = std::stoi(xlpp::internal::attribute(sizes.front(), "val")); } catch (...) {} }
    const auto spPr = drawingTags(marker, "c:spPr", "spPr");
    if (!spPr.empty()) {
        format.line = parseChartLineFormat(spPr.front());
        format.fill = parseChartFillFormat(spPr.front());
    }
    return format;
}

xlpp::ChartRichText parseChartRichText(const std::string& owner) {
    xlpp::ChartRichText result;
    const auto richNodes = drawingTags(owner, "c:rich", "rich");
    if (richNodes.empty()) return result;
    result.present = true;
    const auto& rich = richNodes.front();
    for (const auto& runXml : drawingTags(rich, "a:r", "r")) {
        xlpp::ChartTextRun run;
        run.text = drawingTagText(runXml, "a:t", "t");
        const auto properties = drawingTags(runXml, "a:rPr", "rPr");
        if (!properties.empty()) {
            const auto& rPr = properties.front();
            const auto bold = xlpp::internal::attribute(rPr, "b");
            const auto italic = xlpp::internal::attribute(rPr, "i");
            run.bold = bold == "1" || bold == "true";
            run.italic = italic == "1" || italic == "true";
            const auto size = xlpp::internal::attribute(rPr, "sz");
            if (!size.empty()) { try { run.fontSizePoints = std::stod(size) / 100.0; } catch (...) {} }
            const auto latin = drawingTags(rPr, "a:latin", "latin");
            if (!latin.empty()) run.typeface = xlpp::internal::attribute(latin.front(), "typeface");
            run.color = parseChartColor(rPr);
        }
        result.runs.push_back(std::move(run));
    }
    if (result.runs.empty()) {
        const auto text = drawingTagText(rich, "a:t", "t");
        if (!text.empty()) {
            xlpp::ChartTextRun run;
            run.text = text;
            result.runs.push_back(std::move(run));
        }
    }
    return result;
}

xlpp::ChartTextStyle parseChartTextStyle(const std::string& owner) {
    xlpp::ChartTextStyle style;
    const auto txPr = drawingTags(owner, "c:txPr", "txPr");
    if (txPr.empty()) return style;
    std::string properties;
    const auto defaults = drawingTags(txPr.front(), "a:defRPr", "defRPr");
    if (!defaults.empty()) properties = defaults.front();
    else {
        const auto ends = drawingTags(txPr.front(), "a:endParaRPr", "endParaRPr");
        if (!ends.empty()) properties = ends.front();
        else {
            const auto runs = drawingTags(txPr.front(), "a:rPr", "rPr");
            if (!runs.empty()) properties = runs.front();
        }
    }
    if (properties.empty()) { style.present = true; return style; }
    style.present = true;
    const auto bold = xlpp::internal::attribute(properties, "b");
    const auto italic = xlpp::internal::attribute(properties, "i");
    style.bold = bold == "1" || bold == "true";
    style.italic = italic == "1" || italic == "true";
    const auto size = xlpp::internal::attribute(properties, "sz");
    if (!size.empty()) { try { style.fontSizePoints = std::stod(size) / 100.0; } catch (...) {} }
    const auto latin = drawingTags(properties, "a:latin", "latin");
    if (!latin.empty()) style.typeface = xlpp::internal::attribute(latin.front(), "typeface");
    style.color = parseChartColor(properties);
    return style;
}

std::vector<xlpp::ChartDataPointFormat> parseChartDataPoints(const std::string& seriesXml) {
    std::vector<xlpp::ChartDataPointFormat> result;
    for (const auto& node : drawingTags(seriesXml, "c:dPt", "dPt")) {
        xlpp::ChartDataPointFormat point;
        const auto indices = drawingTags(node, "c:idx", "idx");
        if (!indices.empty()) { try { point.index = static_cast<std::size_t>(std::stoull(xlpp::internal::attribute(indices.front(), "val"))); } catch (...) {} }
        const auto spPr = drawingTags(node, "c:spPr", "spPr");
        if (!spPr.empty()) {
            point.line = parseChartLineFormat(spPr.front());
            point.fill = parseChartFillFormat(spPr.front());
        }
        point.marker = parseChartMarkerFormat(node);
        result.push_back(std::move(point));
    }
    return result;
}

xlpp::ChartDataLabelPoint parseChartDataLabelPoint(const std::string& xml) {
    xlpp::ChartDataLabelPoint point;
    const auto idx = drawingTags(xml, "c:idx", "idx");
    if (!idx.empty()) { try { point.index = static_cast<std::size_t>(std::stoull(xlpp::internal::attribute(idx.front(), "val"))); } catch (...) {} }
    point.deleted = chartBoolValue(xml, "c:delete", "delete");
    point.showLegendKey = chartBoolValue(xml, "c:showLegendKey", "showLegendKey");
    point.showValue = chartBoolValue(xml, "c:showVal", "showVal");
    point.showCategoryName = chartBoolValue(xml, "c:showCatName", "showCatName");
    point.showSeriesName = chartBoolValue(xml, "c:showSerName", "showSerName");
    point.showPercent = chartBoolValue(xml, "c:showPercent", "showPercent");
    point.showBubbleSize = chartBoolValue(xml, "c:showBubbleSize", "showBubbleSize");
    point.showLeaderLines = chartBoolValue(xml, "c:showLeaderLines", "showLeaderLines");
    const auto positions = drawingTags(xml, "c:dLblPos", "dLblPos");
    if (!positions.empty()) point.position = xlpp::internal::attribute(positions.front(), "val");
    point.separator = drawingTagText(xml, "c:separator", "separator");
    point.richText = parseChartRichText(xml);
    return point;
}

xlpp::Chart::DataLabels parseChartDataLabels(const std::string& plotXml, bool directPlotChild = false) {
    xlpp::Chart::DataLabels labels;
    auto nodes = drawingTags(plotXml, "c:dLbls", "dLbls");
    if (directPlotChild && !nodes.empty()) {
        const auto seriesNodes = drawingTags(plotXml, "c:ser", "ser");
        nodes.erase(std::remove_if(nodes.begin(), nodes.end(), [&](const std::string& node) {
            return std::any_of(seriesNodes.begin(), seriesNodes.end(), [&](const std::string& series) {
                return series.find(node) != std::string::npos;
            });
        }), nodes.end());
    }
    if (nodes.empty()) return labels;
    labels.present = true;
    const auto& xml = nodes.front();
    auto aggregateXml = xml;
    for (const auto& pointNode : drawingTags(xml, "c:dLbl", "dLbl")) {
        labels.points.push_back(parseChartDataLabelPoint(pointNode));
        const auto position = aggregateXml.find(pointNode);
        if (position != std::string::npos) aggregateXml.erase(position, pointNode.size());
    }
    labels.showLegendKey = chartBoolValue(aggregateXml, "c:showLegendKey", "showLegendKey");
    labels.showValue = chartBoolValue(aggregateXml, "c:showVal", "showVal");
    labels.showCategoryName = chartBoolValue(aggregateXml, "c:showCatName", "showCatName");
    labels.showSeriesName = chartBoolValue(aggregateXml, "c:showSerName", "showSerName");
    labels.showPercent = chartBoolValue(aggregateXml, "c:showPercent", "showPercent");
    labels.showBubbleSize = chartBoolValue(aggregateXml, "c:showBubbleSize", "showBubbleSize");
    labels.showLeaderLines = chartBoolValue(aggregateXml, "c:showLeaderLines", "showLeaderLines");
    const auto leaderLines = drawingTags(aggregateXml, "c:leaderLines", "leaderLines");
    if (!leaderLines.empty()) {
        labels.hasLeaderLines = true;
        const auto spPr = drawingTags(leaderLines.front(), "c:spPr", "spPr");
        if (!spPr.empty()) labels.leaderLineFormat = parseChartLineFormat(spPr.front());
    }
    const auto positions = drawingTags(aggregateXml, "c:dLblPos", "dLblPos");
    if (!positions.empty()) labels.position = xlpp::internal::attribute(positions.front(), "val");
    labels.separator = drawingTagText(aggregateXml, "c:separator", "separator");
    return labels;
}

xlpp::ChartSeries::TrendlineType parseTrendlineTypeValue(const std::string& value) {
    using T = xlpp::ChartSeries::TrendlineType;
    if (value == "exp") return T::Exponential;
    if (value == "log") return T::Logarithmic;
    if (value == "poly") return T::Polynomial;
    if (value == "power") return T::Power;
    if (value == "movingAvg") return T::MovingAverage;
    return T::Linear;
}

std::vector<xlpp::ChartSeries::Trendline> parseChartTrendlines(const std::string& seriesXml) {
    std::vector<xlpp::ChartSeries::Trendline> result;
    for (const auto& node : drawingTags(seriesXml, "c:trendline", "trendline")) {
        xlpp::ChartSeries::Trendline trendline;
        const auto types = drawingTags(node, "c:trendlineType", "trendlineType");
        if (!types.empty()) trendline.type = parseTrendlineTypeValue(xlpp::internal::attribute(types.front(), "val"));
        const auto order = drawingTags(node, "c:order", "order");
        if (!order.empty()) { try { trendline.order = std::stoi(xlpp::internal::attribute(order.front(), "val")); } catch (...) {} }
        const auto period = drawingTags(node, "c:period", "period");
        if (!period.empty()) { try { trendline.period = std::stoi(xlpp::internal::attribute(period.front(), "val")); } catch (...) {} }
        const auto forward = drawingTags(node, "c:forward", "forward");
        if (!forward.empty()) { try { trendline.forward = std::stod(xlpp::internal::attribute(forward.front(), "val")); } catch (...) {} }
        const auto backward = drawingTags(node, "c:backward", "backward");
        if (!backward.empty()) { try { trendline.backward = std::stod(xlpp::internal::attribute(backward.front(), "val")); } catch (...) {} }
        trendline.displayEquation = chartBoolValue(node, "c:dispEq", "dispEq");
        trendline.displayRSquared = chartBoolValue(node, "c:dispRSqr", "dispRSqr");
        const auto spPr = drawingTags(node, "c:spPr", "spPr");
        if (!spPr.empty()) trendline.lineFormat = parseChartLineFormat(spPr.front());
        result.push_back(std::move(trendline));
    }
    return result;
}

xlpp::ChartSeries::ErrorBarDirection parseErrorBarDirectionValue(const std::string& value) {
    return value == "x" ? xlpp::ChartSeries::ErrorBarDirection::X : xlpp::ChartSeries::ErrorBarDirection::Y;
}

xlpp::ChartSeries::ErrorBarType parseErrorBarTypeValue(const std::string& value) {
    using T = xlpp::ChartSeries::ErrorBarType;
    if (value == "plus") return T::Plus;
    if (value == "minus") return T::Minus;
    return T::Both;
}

xlpp::ChartSeries::ErrorValueType parseErrorValueTypeValue(const std::string& value) {
    using T = xlpp::ChartSeries::ErrorValueType;
    if (value == "percentage") return T::Percentage;
    if (value == "stdDev") return T::StandardDeviation;
    if (value == "stdErr") return T::StandardError;
    if (value == "cust") return T::Custom;
    return T::FixedValue;
}

std::vector<xlpp::ChartSeries::ErrorBars> parseChartErrorBars(const std::string& seriesXml) {
    std::vector<xlpp::ChartSeries::ErrorBars> result;
    for (const auto& node : drawingTags(seriesXml, "c:errBars", "errBars")) {
        xlpp::ChartSeries::ErrorBars bars;
        const auto dirs = drawingTags(node, "c:errDir", "errDir");
        if (!dirs.empty()) bars.direction = parseErrorBarDirectionValue(xlpp::internal::attribute(dirs.front(), "val"));
        const auto types = drawingTags(node, "c:errBarType", "errBarType");
        if (!types.empty()) bars.barType = parseErrorBarTypeValue(xlpp::internal::attribute(types.front(), "val"));
        const auto valueTypes = drawingTags(node, "c:errValType", "errValType");
        if (!valueTypes.empty()) bars.valueType = parseErrorValueTypeValue(xlpp::internal::attribute(valueTypes.front(), "val"));
        bars.noEndCap = chartBoolValue(node, "c:noEndCap", "noEndCap");
        const auto values = drawingTags(node, "c:val", "val");
        if (!values.empty()) { try { bars.value = std::stod(xlpp::internal::attribute(values.front(), "val")); } catch (...) {} }
        const auto plus = drawingTags(node, "c:plus", "plus");
        if (!plus.empty()) bars.plusReference = drawingTagText(plus.front(), "c:f", "f");
        const auto minus = drawingTags(node, "c:minus", "minus");
        if (!minus.empty()) bars.minusReference = drawingTagText(minus.front(), "c:f", "f");
        const auto spPr = drawingTags(node, "c:spPr", "spPr");
        if (!spPr.empty()) bars.lineFormat = parseChartLineFormat(spPr.front());
        result.push_back(std::move(bars));
    }
    return result;
}

xlpp::ChartDataTable parseChartDataTable(const std::string& plotArea) {
    xlpp::ChartDataTable table;
    const auto nodes = drawingTags(plotArea, "c:dTable", "dTable");
    if (nodes.empty()) return table;
    table.present = true;
    const auto& xml = nodes.front();
    table.showHorizontalBorder = chartBoolValue(xml, "c:showHorzBorder", "showHorzBorder");
    table.showVerticalBorder = chartBoolValue(xml, "c:showVertBorder", "showVertBorder");
    table.showOutline = chartBoolValue(xml, "c:showOutline", "showOutline");
    table.showLegendKeys = chartBoolValue(xml, "c:showKeys", "showKeys");
    const auto spPr = drawingTags(xml, "c:spPr", "spPr");
    if (!spPr.empty()) {
        table.line = parseChartLineFormat(spPr.front());
        table.fill = parseChartFillFormat(spPr.front());
    }
    table.textStyle = parseChartTextStyle(xml);
    return table;
}

xlpp::ChartLineFormat parsePlotLineObject(const std::string& plotXml, const char* prefixed, const char* local, bool& present) {
    present = false;
    const auto nodes = drawingTags(plotXml, prefixed, local);
    if (nodes.empty()) return {};
    present = true;
    const auto spPr = drawingTags(nodes.front(), "c:spPr", "spPr");
    return spPr.empty() ? xlpp::ChartLineFormat{} : parseChartLineFormat(spPr.front());
}

xlpp::ChartUpDownBars parseChartUpDownBars(const std::string& plotXml) {
    xlpp::ChartUpDownBars bars;
    const auto nodes = drawingTags(plotXml, "c:upDownBars", "upDownBars");
    if (nodes.empty()) return bars;
    bars.present = true;
    const auto& xml = nodes.front();
    const auto gaps = drawingTags(xml, "c:gapWidth", "gapWidth");
    if (!gaps.empty()) { try { bars.gapWidth = std::stoi(xlpp::internal::attribute(gaps.front(), "val")); } catch (...) {} }
    const auto parseBar = [&](const char* prefixed, const char* local, xlpp::ChartFillFormat& fill, xlpp::ChartLineFormat& line) {
        const auto barNodes = drawingTags(xml, prefixed, local);
        if (barNodes.empty()) return;
        const auto spPr = drawingTags(barNodes.front(), "c:spPr", "spPr");
        if (!spPr.empty()) { fill = parseChartFillFormat(spPr.front()); line = parseChartLineFormat(spPr.front()); }
    };
    parseBar("c:upBars", "upBars", bars.upFill, bars.upLine);
    parseBar("c:downBars", "downBars", bars.downFill, bars.downLine);
    return bars;
}

std::vector<xlpp::Chart::Plot> parseChartPlots(const std::string& chartXmlText) {
    std::vector<xlpp::Chart::Plot> plots;
    std::size_t firstSeries = 0;
    for (const auto& parsed : chartPlotNodesInOrder(chartXmlText)) {
        xlpp::Chart::Plot plot;
        plot.type = parsed.type;
        if (parsed.type == xlpp::Chart::Type::Bar) {
            const auto directions = drawingTags(parsed.xml, "c:barDir", "barDir");
            if (!directions.empty() && xlpp::internal::attribute(directions.front(), "val") == "bar") {
                plot.type = xlpp::Chart::Type::HorizontalBar;
                plot.barDirection = xlpp::Chart::BarDirection::Bar;
            }
        }
        if (parsed.type == xlpp::Chart::Type::Scatter) {
            const auto styles = drawingTags(parsed.xml, "c:scatterStyle", "scatterStyle");
            const auto value = styles.empty() ? std::string{"marker"} : xlpp::internal::attribute(styles.front(), "val");
            if (value == "none") plot.scatterStyle = xlpp::Chart::ScatterStyle::None;
            else if (value == "line") plot.scatterStyle = xlpp::Chart::ScatterStyle::Line;
            else if (value == "lineMarker") plot.scatterStyle = xlpp::Chart::ScatterStyle::LineMarker;
            else if (value == "smooth") plot.scatterStyle = xlpp::Chart::ScatterStyle::Smooth;
            else if (value == "smoothMarker") plot.scatterStyle = xlpp::Chart::ScatterStyle::SmoothMarker;
            else plot.scatterStyle = xlpp::Chart::ScatterStyle::Marker;
        }
        if (parsed.type == xlpp::Chart::Type::Bubble) {
            const auto scale = drawingTags(parsed.xml, "c:bubbleScale", "bubbleScale");
            if (!scale.empty()) { try { plot.bubbleScale = std::stoi(xlpp::internal::attribute(scale.front(), "val")); plot.hasBubbleScale = true; } catch (...) {} }
            plot.showNegativeBubbles = chartBoolValue(parsed.xml, "c:showNegBubbles", "showNegBubbles", false);
            const auto represents = drawingTags(parsed.xml, "c:sizeRepresents", "sizeRepresents");
            if (!represents.empty() && xlpp::internal::attribute(represents.front(), "val") == "w") plot.bubbleSizeRepresents = xlpp::Chart::BubbleSizeRepresents::Width;
            plot.bubble3D = chartBoolValue(parsed.xml, "c:bubble3D", "bubble3D", false);
        }
        plot.grouping = parseChartGroupingNode(parsed.xml);
        plot.axisIds = chartAxisIds(parsed.xml);
        plot.firstSeries = firstSeries;
        plot.seriesCount = drawingTags(parsed.xml, "c:ser", "ser").size();
        plot.dataLabels = parseChartDataLabels(parsed.xml, true);
        plot.dropLinesFormat = parsePlotLineObject(parsed.xml, "c:dropLines", "dropLines", plot.hasDropLines);
        plot.highLowLinesFormat = parsePlotLineObject(parsed.xml, "c:hiLowLines", "hiLowLines", plot.hasHighLowLines);
        plot.upDownBars = parseChartUpDownBars(parsed.xml);
        const auto gapDepth = drawingTags(parsed.xml, "c:gapDepth", "gapDepth");
        if (!gapDepth.empty()) { try { plot.gapDepth=std::stoi(xlpp::internal::attribute(gapDepth.front(),"val")); plot.hasGapDepth=true; } catch (...) {} }
        const auto wireframe = drawingTags(parsed.xml, "c:wireframe", "wireframe");
        if (!wireframe.empty()) { plot.hasWireframe=true; const auto value=xlpp::internal::attribute(wireframe.front(),"val"); plot.wireframe=value=="1"||value=="true"||value=="True"; }
        const auto shape = drawingTags(parsed.xml, "c:shape", "shape");
        if (!shape.empty()) plot.shape=xlpp::internal::attribute(shape.front(),"val");
        const auto firstSlice = drawingTags(parsed.xml, "c:firstSliceAng", "firstSliceAng");
        if (!firstSlice.empty()) { try { plot.firstSliceAngle=std::stoi(xlpp::internal::attribute(firstSlice.front(),"val")); plot.hasFirstSliceAngle=true; } catch (...) {} }
        const auto holeSize = drawingTags(parsed.xml, "c:holeSize", "holeSize");
        if (!holeSize.empty()) { try { plot.holeSize=std::stoi(xlpp::internal::attribute(holeSize.front(),"val")); plot.hasHoleSize=true; } catch (...) {} }
        const auto radarStyle = drawingTags(parsed.xml, "c:radarStyle", "radarStyle");
        if (!radarStyle.empty()) plot.radarStyle=xlpp::internal::attribute(radarStyle.front(),"val");
        if (parsed.type == xlpp::Chart::Type::PieOfPie || parsed.type == xlpp::Chart::Type::BarOfPie) {
            auto& options=plot.projectedPie; options.present=true; options.ofPieType=parsed.type==xlpp::Chart::Type::BarOfPie?"bar":"pie";
            const auto gapWidth=drawingTags(parsed.xml,"c:gapWidth","gapWidth"); if(!gapWidth.empty()) { try { options.gapWidth=std::stoi(xlpp::internal::attribute(gapWidth.front(),"val")); } catch (...) {} }
            const auto splitType=drawingTags(parsed.xml,"c:splitType","splitType"); if(!splitType.empty()) options.splitType=xlpp::internal::attribute(splitType.front(),"val");
            const auto splitPos=drawingTags(parsed.xml,"c:splitPos","splitPos"); if(!splitPos.empty()) { try { options.splitPosition=std::stod(xlpp::internal::attribute(splitPos.front(),"val")); options.hasSplitPosition=true; } catch (...) {} }
            const auto secondSize=drawingTags(parsed.xml,"c:secondPieSize","secondPieSize"); if(!secondSize.empty()) { try { options.secondPlotSize=std::stoi(xlpp::internal::attribute(secondSize.front(),"val")); } catch (...) {} }
            const auto custom=drawingTags(parsed.xml,"c:custSplit","custSplit"); if(!custom.empty()) for(const auto& pt:drawingTags(custom.front(),"c:secondPiePt","secondPiePt")) { try { options.customSplitPoints.push_back(std::stoi(xlpp::internal::attribute(pt,"val"))); } catch (...) {} }
            options.seriesLinesFormat=parsePlotLineObject(parsed.xml,"c:serLines","serLines",options.hasSeriesLines);
        }
        firstSeries += plot.seriesCount;
        plots.push_back(std::move(plot));
    }
    return plots;
}

xlpp::ChartManualLayout parseChartManualLayout(const std::string& owner) {
    xlpp::ChartManualLayout layout;
    const auto layouts = drawingTags(owner, "c:layout", "layout");
    if (layouts.empty()) return layout;
    const auto manuals = drawingTags(layouts.front(), "c:manualLayout", "manualLayout");
    if (manuals.empty()) return layout;
    layout.present = true;
    const auto& xml = manuals.front();
    const auto readString=[&](const char* p,const char* l){ const auto n=drawingTags(xml,p,l); return n.empty()?std::string{}:xlpp::internal::attribute(n.front(),"val"); };
    const auto readDouble=[&](const char* p,const char* l,bool& has,double& value){ const auto n=drawingTags(xml,p,l); if(n.empty()) return; const auto v=xlpp::internal::attribute(n.front(),"val"); if(v.empty()) return; try{ value=std::stod(v); has=true; }catch(...){} };
    layout.target=readString("c:layoutTarget","layoutTarget");
    layout.xMode=readString("c:xMode","xMode"); layout.yMode=readString("c:yMode","yMode");
    layout.widthMode=readString("c:wMode","wMode"); layout.heightMode=readString("c:hMode","hMode");
    readDouble("c:x","x",layout.hasX,layout.x); readDouble("c:y","y",layout.hasY,layout.y);
    readDouble("c:w","w",layout.hasWidth,layout.width); readDouble("c:h","h",layout.hasHeight,layout.height);
    return layout;
}

xlpp::ChartView3D parseChartView3D(const std::string& chartXmlText) {
    xlpp::ChartView3D view;
    const auto nodes = drawingTags(chartXmlText, "c:view3D", "view3D");
    if (nodes.empty()) return view;
    view.present = true;
    const auto& xml = nodes.front();
    const auto readInt = [&](const char* prefixed, const char* local, bool& has, int& out) {
        const auto values = drawingTags(xml, prefixed, local);
        if (values.empty()) return;
        const auto value = xlpp::internal::attribute(values.front(), "val");
        if (value.empty()) return;
        try { out = std::stoi(value); has = true; } catch (...) {}
    };
    readInt("c:rotX", "rotX", view.hasRotationX, view.rotationX);
    readInt("c:rotY", "rotY", view.hasRotationY, view.rotationY);
    readInt("c:hPercent", "hPercent", view.hasHeightPercent, view.heightPercent);
    readInt("c:depthPercent", "depthPercent", view.hasDepthPercent, view.depthPercent);
    readInt("c:perspective", "perspective", view.hasPerspective, view.perspective);
    const auto rAng = drawingTags(xml, "c:rAngAx", "rAngAx");
    if (!rAng.empty()) {
        view.hasRightAngleAxes = true;
        const auto value = xlpp::internal::attribute(rAng.front(), "val");
        view.rightAngleAxes = value.empty() || value == "1" || value == "true" || value == "True";
    }
    return view;
}

xlpp::ChartWallFormat parseChartWallFormat(const std::string& chartXmlText, const char* prefixed, const char* local) {
    xlpp::ChartWallFormat wall;
    const auto nodes = drawingTags(chartXmlText, prefixed, local);
    if (nodes.empty()) return wall;
    wall.present = true;
    const auto& xml = nodes.front();
    const auto thickness = drawingTags(xml, "c:thickness", "thickness");
    if (!thickness.empty()) {
        try {
            const auto value = xlpp::internal::attribute(thickness.front(), "val");
            if (!value.empty()) { wall.thickness = std::stoi(value); wall.hasThickness = true; }
        } catch (...) {}
    }
    const auto spPr = drawingTags(xml, "c:spPr", "spPr");
    if (!spPr.empty()) {
        wall.line = parseChartLineFormat(spPr.front());
        wall.fill = parseChartFillFormat(spPr.front());
    }
    return wall;
}

std::vector<xlpp::Chart::Axis> parseChartAxes(const std::string& chartXmlText,
                                               const std::vector<xlpp::Chart::Plot>& plots) {
    const auto plotAreas = drawingTags(chartXmlText, "c:plotArea", "plotArea");
    if (plotAreas.empty()) return {};
    const auto& plotArea = plotAreas.front();
    struct AxisNode { std::size_t position; xlpp::Chart::AxisKind kind; std::string xml; };
    std::vector<AxisNode> nodes;
    const auto collect = [&](const char* prefixed, const char* local, xlpp::Chart::AxisKind kind) {
        std::size_t cursor = 0;
        for (const auto& node : drawingTags(plotArea, prefixed, local)) {
            const auto position = plotArea.find(node, cursor);
            if (position == std::string::npos) continue;
            nodes.push_back({position, kind, node});
            cursor = position + node.size();
        }
    };
    collect("c:catAx", "catAx", xlpp::Chart::AxisKind::Category);
    collect("c:valAx", "valAx", xlpp::Chart::AxisKind::Value);
    collect("c:dateAx", "dateAx", xlpp::Chart::AxisKind::Date);
    collect("c:serAx", "serAx", xlpp::Chart::AxisKind::Series);
    std::sort(nodes.begin(), nodes.end(), [](const auto& a, const auto& b) { return a.position < b.position; });

    std::set<std::uint64_t> primaryAxisIds;
    if (!plots.empty()) primaryAxisIds.insert(plots.front().axisIds.begin(), plots.front().axisIds.end());
    std::set<std::uint64_t> laterPlotAxisIds;
    for (std::size_t i = 1; i < plots.size(); ++i)
        laterPlotAxisIds.insert(plots[i].axisIds.begin(), plots[i].axisIds.end());

    std::vector<xlpp::Chart::Axis> axes;
    for (const auto& parsed : nodes) {
        xlpp::Chart::Axis axis;
        axis.kind = parsed.kind;
        const auto ids = drawingTags(parsed.xml, "c:axId", "axId");
        if (!ids.empty()) {
            const auto value = xlpp::internal::attribute(ids.front(), "val");
            try { if (!value.empty()) axis.id = std::stoull(value); } catch (...) {}
        }
        const auto crosses = drawingTags(parsed.xml, "c:crossAx", "crossAx");
        if (!crosses.empty()) {
            const auto value = xlpp::internal::attribute(crosses.front(), "val");
            try { if (!value.empty()) axis.crossAxisId = std::stoull(value); } catch (...) {}
        }
        const auto positions = drawingTags(parsed.xml, "c:axPos", "axPos");
        if (!positions.empty()) axis.position = xlpp::internal::attribute(positions.front(), "val");
        const auto scalings = drawingTags(parsed.xml, "c:scaling", "scaling");
        if (!scalings.empty()) {
            const auto& scaling = scalings.front();
            const auto minimum = drawingTags(scaling, "c:min", "min");
            if (!minimum.empty()) { try { axis.scaling.minimum=std::stod(xlpp::internal::attribute(minimum.front(),"val")); axis.scaling.hasMinimum=true; } catch (...) {} }
            const auto maximum = drawingTags(scaling, "c:max", "max");
            if (!maximum.empty()) { try { axis.scaling.maximum=std::stod(xlpp::internal::attribute(maximum.front(),"val")); axis.scaling.hasMaximum=true; } catch (...) {} }
            const auto logBase = drawingTags(scaling, "c:logBase", "logBase");
            if (!logBase.empty()) { try { axis.scaling.logBase=std::stod(xlpp::internal::attribute(logBase.front(),"val")); axis.scaling.hasLogBase=true; } catch (...) {} }
            const auto orientation = drawingTags(scaling, "c:orientation", "orientation");
            if (!orientation.empty()) axis.scaling.reverseOrder = xlpp::internal::attribute(orientation.front(),"val") == "maxMin";
        }
        const auto titles = drawingTags(parsed.xml, "c:title", "title");
        if (!titles.empty()) {
            axis.title = drawingTagText(titles.front(), "a:t", "t");
            if (axis.title.empty()) axis.title = drawingTagText(titles.front(), "c:v", "v");
            if (axis.title.empty()) axis.title = drawingTagText(titles.front(), "c:f", "f");
            axis.titleRichText = parseChartRichText(titles.front());
            if (axis.title.empty() && axis.titleRichText.present) axis.title = axis.titleRichText.plainText();
        }
        const auto numFmt = drawingTags(parsed.xml, "c:numFmt", "numFmt");
        if (!numFmt.empty()) {
            axis.numberFormat = xlpp::internal::attribute(numFmt.front(), "formatCode");
            const auto linked = xlpp::internal::attribute(numFmt.front(), "sourceLinked");
            axis.numberFormatSourceLinked = linked.empty() || linked == "1" || linked == "true";
        }
        const auto majorTick = drawingTags(parsed.xml, "c:majorTickMark", "majorTickMark");
        if (!majorTick.empty()) axis.majorTickMark = xlpp::internal::attribute(majorTick.front(), "val");
        const auto minorTick = drawingTags(parsed.xml, "c:minorTickMark", "minorTickMark");
        if (!minorTick.empty()) axis.minorTickMark = xlpp::internal::attribute(minorTick.front(), "val");
        const auto tickPos = drawingTags(parsed.xml, "c:tickLblPos", "tickLblPos");
        if (!tickPos.empty()) axis.tickLabelPosition = xlpp::internal::attribute(tickPos.front(), "val");
        const auto majorUnit = drawingTags(parsed.xml, "c:majorUnit", "majorUnit");
        if (!majorUnit.empty()) { try { axis.majorUnit = std::stod(xlpp::internal::attribute(majorUnit.front(), "val")); axis.hasMajorUnit = true; } catch (...) {} }
        const auto minorUnit = drawingTags(parsed.xml, "c:minorUnit", "minorUnit");
        if (!minorUnit.empty()) { try { axis.minorUnit = std::stod(xlpp::internal::attribute(minorUnit.front(), "val")); axis.hasMinorUnit = true; } catch (...) {} }
        const auto crossesNode = drawingTags(parsed.xml, "c:crosses", "crosses");
        if (!crossesNode.empty()) axis.crosses = xlpp::internal::attribute(crossesNode.front(), "val");
        const auto crossesAtNode = drawingTags(parsed.xml, "c:crossesAt", "crossesAt");
        if (!crossesAtNode.empty()) { try { axis.crossesAt=std::stod(xlpp::internal::attribute(crossesAtNode.front(),"val")); axis.hasCrossesAt=true; } catch (...) {} }
        const auto crossBetweenNode = drawingTags(parsed.xml, "c:crossBetween", "crossBetween");
        if (!crossBetweenNode.empty()) axis.crossBetween = xlpp::internal::attribute(crossBetweenNode.front(), "val");
        const auto displayUnits = drawingTags(parsed.xml, "c:dispUnits", "dispUnits");
        if (!displayUnits.empty()) {
            axis.displayUnits.present = true;
            const auto builtIn = drawingTags(displayUnits.front(), "c:builtInUnit", "builtInUnit");
            if (!builtIn.empty()) axis.displayUnits.builtInUnit = xlpp::internal::attribute(builtIn.front(), "val");
            const auto custom = drawingTags(displayUnits.front(), "c:custUnit", "custUnit");
            if (!custom.empty()) { try { axis.displayUnits.customUnit=std::stod(xlpp::internal::attribute(custom.front(),"val")); axis.displayUnits.hasCustomUnit=true; } catch (...) {} }
            const auto labels = drawingTags(displayUnits.front(), "c:dispUnitsLbl", "dispUnitsLbl");
            if (!labels.empty()) { axis.displayUnits.showLabel=true; axis.displayUnits.labelRichText=parseChartRichText(labels.front()); }
        }
        const auto axisSpPr = axisDirectSpPr(parsed.xml);
        if (!axisSpPr.empty()) axis.lineFormat = parseChartLineFormat(axisSpPr);
        const auto majorGrid = drawingTags(parsed.xml, "c:majorGridlines", "majorGridlines");
        axis.hasMajorGridlines = !majorGrid.empty();
        if (!majorGrid.empty()) { const auto spPr=drawingTags(majorGrid.front(),"c:spPr","spPr"); if(!spPr.empty()) axis.majorGridlineFormat=parseChartLineFormat(spPr.front()); }
        const auto minorGrid = drawingTags(parsed.xml, "c:minorGridlines", "minorGridlines");
        axis.hasMinorGridlines = !minorGrid.empty();
        if (!minorGrid.empty()) { const auto spPr=drawingTags(minorGrid.front(),"c:spPr","spPr"); if(!spPr.empty()) axis.minorGridlineFormat=parseChartLineFormat(spPr.front()); }
        axis.secondary = axis.id != 0 && primaryAxisIds.find(axis.id) == primaryAxisIds.end() &&
            laterPlotAxisIds.find(axis.id) != laterPlotAxisIds.end();
        axes.push_back(std::move(axis));
    }
    return axes;
}

std::string axisTitleById(const std::vector<xlpp::Chart::Axis>& axes, std::uint64_t axisId) {
    const auto it = std::find_if(axes.begin(), axes.end(), [&](const auto& axis) { return axis.id == axisId; });
    return it == axes.end() ? std::string{} : it->title;
}

xlpp::Chart::Type parseChartType(const std::string& chartXmlText) {
    using Type = xlpp::Chart::Type;
    const auto bars = drawingTags(chartXmlText, "c:barChart", "barChart");
    if (!bars.empty()) {
        const auto directions = drawingTags(bars.front(), "c:barDir", "barDir");
        if (!directions.empty() && xlpp::internal::attribute(directions.front(), "val") == "bar") return Type::HorizontalBar;
        return Type::Bar;
    }
    if (!drawingTags(chartXmlText, "c:lineChart", "lineChart").empty()) return Type::Line;
    if (!drawingTags(chartXmlText, "c:pieChart", "pieChart").empty()) return Type::Pie;
    if (!drawingTags(chartXmlText, "c:scatterChart", "scatterChart").empty()) return Type::Scatter;
    if (!drawingTags(chartXmlText, "c:doughnutChart", "doughnutChart").empty()) return Type::Doughnut;
    if (!drawingTags(chartXmlText, "c:radarChart", "radarChart").empty()) return Type::Radar;
    if (!drawingTags(chartXmlText, "c:areaChart", "areaChart").empty()) return Type::Area;
    if (!drawingTags(chartXmlText, "c:bubbleChart", "bubbleChart").empty()) return Type::Bubble;
    if (!drawingTags(chartXmlText, "c:stockChart", "stockChart").empty()) return Type::Stock;
    const auto projectedPie = drawingTags(chartXmlText, "c:ofPieChart", "ofPieChart");
    if (!projectedPie.empty()) {
        const auto kinds=drawingTags(projectedPie.front(),"c:ofPieType","ofPieType");
        return !kinds.empty() && xlpp::internal::attribute(kinds.front(),"val")=="bar" ? Type::BarOfPie : Type::PieOfPie;
    }
    if (!drawingTags(chartXmlText, "c:bar3DChart", "bar3DChart").empty()) return Type::Bar3D;
    if (!drawingTags(chartXmlText, "c:line3DChart", "line3DChart").empty()) return Type::Line3D;
    if (!drawingTags(chartXmlText, "c:area3DChart", "area3DChart").empty()) return Type::Area3D;
    if (!drawingTags(chartXmlText, "c:pie3DChart", "pie3DChart").empty()) return Type::Pie3D;
    if (!drawingTags(chartXmlText, "c:surfaceChart", "surfaceChart").empty()) return Type::Surface;
    if (!drawingTags(chartXmlText, "c:surface3DChart", "surface3DChart").empty()) return Type::Surface3D;
    return Type::Bar;
}

xlpp::Chart::Grouping parseChartGrouping(const std::string& chartXmlText, xlpp::Chart::Type type) {
    const char* prefixed = "c:barChart";
    const char* local = "barChart";
    switch (type) {
        case xlpp::Chart::Type::Line: prefixed = "c:lineChart"; local = "lineChart"; break;
        case xlpp::Chart::Type::Area: prefixed = "c:areaChart"; local = "areaChart"; break;
        case xlpp::Chart::Type::Bar3D: prefixed = "c:bar3DChart"; local = "bar3DChart"; break;
        case xlpp::Chart::Type::Line3D: prefixed = "c:line3DChart"; local = "line3DChart"; break;
        case xlpp::Chart::Type::Area3D: prefixed = "c:area3DChart"; local = "area3DChart"; break;
        default: break;
    }
    const auto chartTypeNodes = drawingTags(chartXmlText, prefixed, local);
    if (chartTypeNodes.empty()) return xlpp::Chart::Grouping::Standard;
    const auto groupingNodes = drawingTags(chartTypeNodes.front(), "c:grouping", "grouping");
    if (groupingNodes.empty()) return xlpp::Chart::Grouping::Standard;
    const auto value = xlpp::internal::attribute(groupingNodes.front(), "val");
    if (value == "stacked") return xlpp::Chart::Grouping::Stacked;
    if (value == "percentStacked") return xlpp::Chart::Grouping::PercentStacked;
    if (value == "clustered") return xlpp::Chart::Grouping::Clustered;
    return xlpp::Chart::Grouping::Standard;
}

std::string chartTitleText(const std::string& chartXmlText) {
    const auto chartNodes = drawingTags(chartXmlText, "c:chart", "chart");
    if (chartNodes.empty()) return {};
    const auto& chartNode = chartNodes.front();
    auto plotPosition = chartNode.find("<c:plotArea");
    if (plotPosition == std::string::npos) plotPosition = chartNode.find("<plotArea");
    for (const auto& titleNode : drawingTags(chartNode, "c:title", "title")) {
        const auto titlePosition = chartNode.find(titleNode);
        if (plotPosition != std::string::npos && titlePosition > plotPosition) continue;
        auto value = drawingTagText(titleNode, "a:t", "t");
        if (value.empty()) value = drawingTagText(titleNode, "c:v", "v");
        if (value.empty()) value = drawingTagText(titleNode, "c:f", "f");
        return value;
    }
    return {};
}

std::string axisTitleText(const std::string& chartXmlText, const char* prefixedAxis, const char* localAxis, std::size_t axisIndex = 0) {
    const auto axes = drawingTags(chartXmlText, prefixedAxis, localAxis);
    if (axisIndex >= axes.size()) return {};
    const auto titles = drawingTags(axes[axisIndex], "c:title", "title");
    if (titles.empty()) return {};
    auto value = drawingTagText(titles.front(), "a:t", "t");
    if (value.empty()) value = drawingTagText(titles.front(), "c:v", "v");
    if (value.empty()) value = drawingTagText(titles.front(), "c:f", "f");
    return value;
}


std::string chartExInnerText(const std::string& node) {
    const auto begin = node.find('>');
    if (begin == std::string::npos) return {};
    const auto end = node.rfind("</");
    if (end == std::string::npos || end <= begin) return {};
    return xlpp::internal::xmlUnescape(std::string_view(node).substr(begin + 1, end - begin - 1));
}

xlpp::ChartSeriesCache parseChartExSeriesCache(const std::string& dimension, bool numeric) {
    xlpp::ChartSeriesCache result;
    const auto levels = drawingTags(dimension, "cx:lvl", "lvl");
    if (levels.empty()) return result;
    result.present = true; result.numeric = numeric;
    result.formatCode = xlpp::internal::attribute(levels.front(), "formatCode");
    if (const auto count = xlpp::internal::attribute(levels.front(), "ptCount"); !count.empty()) { try { result.pointCount = static_cast<std::size_t>(std::stoull(count)); } catch (...) {} }
    for (const auto& point : drawingTags(levels.front(), "cx:pt", "pt")) {
        const auto index = xlpp::internal::attribute(point, "idx");
        if (index.empty()) continue;
        try { result.points.push_back({static_cast<std::size_t>(std::stoull(index)), chartExInnerText(point)}); } catch (...) {}
    }
    return result;
}

xlpp::Chart::Type parseChartExType(const std::string& chartXmlText) {
    using T = xlpp::Chart::Type;
    const auto series = drawingTags(chartXmlText, "cx:series", "series");
    if (series.empty()) return T::Histogram;
    bool hasParetoLine = false;
    for (const auto& node : series) if (xlpp::internal::attribute(node, "layoutId") == "paretoLine") hasParetoLine = true;
    const auto layout = xlpp::internal::attribute(series.front(), "layoutId");
    if (layout == "boxWhisker") return T::BoxWhisker;
    if (layout == "funnel") return T::Funnel;
    if (layout == "regionMap") return T::FilledMap;
    if (layout == "sunburst") return T::Sunburst;
    if (layout == "treemap") return T::Treemap;
    if (layout == "waterfall") return T::Waterfall;
    if (hasParetoLine) return T::Pareto;
    return T::Histogram;
}

xlpp::Chart::Plot parseChartExPlot(const std::string& chartXmlText, xlpp::Chart::Type type, std::size_t seriesCount) {
    xlpp::Chart::Plot plot; plot.type = type; plot.firstSeries = 0; plot.seriesCount = seriesCount;
    const auto layoutProperties = drawingTags(chartXmlText, "cx:layoutPr", "layoutPr");
    if (layoutProperties.empty()) return plot;
    const auto& layout = layoutProperties.front();
    if (type == xlpp::Chart::Type::Histogram || type == xlpp::Chart::Type::Pareto) {
        const auto bins = drawingTags(layout, "cx:binning", "binning");
        if (!bins.empty()) {
            const auto size = drawingTags(bins.front(), "cx:binSize", "binSize");
            const auto count = drawingTags(bins.front(), "cx:binCount", "binCount");
            if (!size.empty()) { try { plot.histogramBinWidth=std::stod(xlpp::internal::attribute(size.front(),"val")); plot.histogramAutomaticBins=false; } catch (...) {} }
            else if (!count.empty()) { try { plot.histogramBinCount=std::stoi(xlpp::internal::attribute(count.front(),"val")); plot.histogramAutomaticBins=false; } catch (...) {} }
            if (const auto v=xlpp::internal::attribute(bins.front(),"underflow"); !v.empty()) { try { plot.histogramUnderflow=std::stod(v); plot.histogramHasUnderflow=true; } catch (...) {} }
            if (const auto v=xlpp::internal::attribute(bins.front(),"overflow"); !v.empty()) { try { plot.histogramOverflow=std::stod(v); plot.histogramHasOverflow=true; } catch (...) {} }
        }
    } else if (type == xlpp::Chart::Type::BoxWhisker) {
        const auto visibility = drawingTags(layout,"cx:visibility","visibility");
        if (!visibility.empty()) {
            const auto truth=[](const std::string& v,bool fallback){ return v.empty()?fallback:(v=="1"||v=="true"||v=="True"); };
            plot.boxWhiskerShowMeanLine=truth(xlpp::internal::attribute(visibility.front(),"meanLine"),false);
            plot.boxWhiskerShowMeanMarker=truth(xlpp::internal::attribute(visibility.front(),"meanMarker"),true);
            plot.boxWhiskerShowInnerPoints=truth(xlpp::internal::attribute(visibility.front(),"nonoutliers"),true);
            plot.boxWhiskerShowOutlierPoints=truth(xlpp::internal::attribute(visibility.front(),"outliers"),true);
        }
        const auto stats=drawingTags(layout,"cx:statistics","statistics");
        if(!stats.empty()) plot.boxWhiskerQuartileInclusive=xlpp::internal::attribute(stats.front(),"quartileMethod")=="inclusive";
    } else if (type == xlpp::Chart::Type::Waterfall) {
        const auto visibility=drawingTags(layout,"cx:visibility","visibility");
        if(!visibility.empty()) { const auto v=xlpp::internal::attribute(visibility.front(),"connectorLines"); plot.waterfallShowConnectorLines=v.empty()||v=="1"||v=="true"||v=="True"; }
    }
    return plot;
}

void populateChartExModel(xlpp::Chart& chart, const std::string& chartText) {
    const auto dataNodes = drawingTags(chartText, "cx:data", "data");
    std::unordered_map<std::string, std::string> titlesByDataId;
    for (const auto& seriesNode : drawingTags(chartText, "cx:series", "series")) {
        const auto dataIds = drawingTags(seriesNode, "cx:dataId", "dataId");
        if (dataIds.empty()) continue;
        const auto id = xlpp::internal::attribute(dataIds.front(), "val");
        if (titlesByDataId.count(id)) continue;
        const auto tx = drawingTags(seriesNode, "cx:tx", "tx");
        if (!tx.empty()) titlesByDataId[id] = drawingTagText(tx.front(), "cx:v", "v");
    }
    for (const auto& dataNode : dataNodes) {
        xlpp::ChartSeries series;
        const auto id = xlpp::internal::attribute(dataNode, "id");
        if (const auto it=titlesByDataId.find(id); it!=titlesByDataId.end()) series.setTitle(it->second);
        const auto categories = drawingTags(dataNode, "cx:strDim", "strDim");
        if (!categories.empty()) { series.setCategoriesReference(drawingTagText(categories.front(),"cx:f","f")); series.setCategoriesCache(parseChartExSeriesCache(categories.front(), false)); }
        const auto values = drawingTags(dataNode, "cx:numDim", "numDim");
        if (!values.empty()) { series.setValuesReference(drawingTagText(values.front(),"cx:f","f")); series.setValuesCache(parseChartExSeriesCache(values.front(), true)); }
        chart.addSeries(std::move(series));
    }
    chart.setPlots({parseChartExPlot(chartText, chart.type(), chart.series().size())});
    const auto titles=drawingTags(chartText,"cx:title","title");
    if(!titles.empty()) chart.setTitle(drawingTagText(titles.front(),"cx:v","v"));
    const auto legends=drawingTags(chartText,"cx:legend","legend");
    chart.setShowLegend(!legends.empty());
    if(!legends.empty()) { const auto pos=xlpp::internal::attribute(legends.front(),"pos"); if(!pos.empty()) chart.setLegendPosition(pos); }
}

xlpp::DrawingAnchorInfo parseChartAnchorInfo(const std::string& anchorNode,
                                             xlpp::DrawingAnchorType type,
                                             const std::string& graphicFrame) {
    xlpp::DrawingAnchorInfo info;
    info.type = type;
    info.editAs = xlpp::internal::attribute(anchorNode, "editAs");
    const auto fromNodes = drawingTags(anchorNode, "xdr:from", "from");
    if (!fromNodes.empty()) info.from = parseDrawingMarker(fromNodes.front());
    const auto toNodes = drawingTags(anchorNode, "xdr:to", "to");
    if (!toNodes.empty()) info.to = parseDrawingMarker(toNodes.front());
    const auto posNodes = drawingTags(anchorNode, "xdr:pos", "pos");
    if (!posNodes.empty()) {
        const auto x = xlpp::internal::attribute(posNodes.front(), "x");
        const auto y = xlpp::internal::attribute(posNodes.front(), "y");
        if (!x.empty()) info.xEmu = std::stoll(x);
        if (!y.empty()) info.yEmu = std::stoll(y);
    }
    const auto anchorExt = drawingTags(anchorNode, "xdr:ext", "ext");
    if (!anchorExt.empty()) {
        const auto cx = xlpp::internal::attribute(anchorExt.front(), "cx");
        const auto cy = xlpp::internal::attribute(anchorExt.front(), "cy");
        if (!cx.empty()) info.widthEmu = std::stoll(cx);
        if (!cy.empty()) info.heightEmu = std::stoll(cy);
    }
    if ((info.widthEmu <= 0 || info.heightEmu <= 0) && type == xlpp::DrawingAnchorType::TwoCell) {
        const auto frameExt = drawingTags(graphicFrame, "a:ext", "ext");
        if (!frameExt.empty()) {
            const auto cx = xlpp::internal::attribute(frameExt.front(), "cx");
            const auto cy = xlpp::internal::attribute(frameExt.front(), "cy");
            if (!cx.empty()) info.widthEmu = std::stoll(cx);
            if (!cy.empty()) info.heightEmu = std::stoll(cy);
        }
    }
    return info;
}

void loadWorksheetCharts(xlpp::Worksheet& ws, const std::string& sheetXml, const xlpp::internal::ZipArchive& z,
                const std::string& sheetPart) {
    const auto drawings = xlpp::internal::tags(sheetXml, "drawing");
    if (drawings.empty()) return;
    const auto sheetSlash = sheetPart.find_last_of('/');
    const auto sheetFile = sheetPart.substr(sheetSlash + 1);
    const auto sheetRelsPart = sheetPart.substr(0, sheetSlash + 1) + "_rels/" + sheetFile + ".rels";
    if (!z.contains(sheetRelsPart)) return;
    std::unordered_map<std::string, std::string> sheetRelationships;
    for (const auto& rel : xlpp::internal::tags(z.get(sheetRelsPart), "Relationship"))
        if (xlpp::internal::attribute(rel, "Type").find("/drawing") != std::string::npos)
            sheetRelationships[xlpp::internal::attribute(rel, "Id")] = xlpp::internal::attribute(rel, "Target");

    for (const auto& drawingNode : drawings) {
        const auto relationshipId = xlpp::internal::attribute(drawingNode, "r:id");
        const auto relationship = sheetRelationships.find(relationshipId);
        if (relationship == sheetRelationships.end()) continue;
        const auto drawingPart = resolvePackagePart(sheetPart, relationship->second);
        if (!z.contains(drawingPart)) continue;
        const auto drawingXmlText = z.get(drawingPart);
        const auto drawingSlash = drawingPart.find_last_of('/');
        const auto drawingFile = drawingPart.substr(drawingSlash + 1);
        const auto drawingRelsPart = drawingPart.substr(0, drawingSlash + 1) + "_rels/" + drawingFile + ".rels";
        if (!z.contains(drawingRelsPart)) continue;
        std::unordered_map<std::string, std::string> drawingRelationships;
        for (const auto& rel : xlpp::internal::tags(z.get(drawingRelsPart), "Relationship"))
            if (xlpp::internal::attribute(rel, "Type").find("/chart") != std::string::npos)
                drawingRelationships[xlpp::internal::attribute(rel, "Id")] = xlpp::internal::attribute(rel, "Target");

        const auto loadAnchors = [&](const char* prefixedName, const char* localName, xlpp::DrawingAnchorType anchorType) {
            for (const auto& anchorNode : drawingTags(drawingXmlText, prefixedName, localName)) {
                const auto frames = drawingTags(anchorNode, "xdr:graphicFrame", "graphicFrame");
                if (frames.empty()) continue;
                const auto& frame = frames.front();
                auto chartRefs = drawingTags(frame, "c:chart", "chart");
                bool chartExReference = false;
                if (chartRefs.empty()) { chartRefs = drawingTags(frame, "cx:chart", "chart"); chartExReference = !chartRefs.empty(); }
                if (chartRefs.empty()) continue;
                const auto chartRelationshipId = xlpp::internal::attribute(chartRefs.front(), "r:id");
                const auto chartRelationship = drawingRelationships.find(chartRelationshipId);
                if (chartRelationship == drawingRelationships.end()) continue;
                const auto chartPart = resolvePackagePart(drawingPart, chartRelationship->second);
                if (!z.contains(chartPart)) continue;
                const auto chartText = z.get(chartPart);
                const bool chartEx = chartExReference || chartText.find("http://schemas.microsoft.com/office/drawing/2014/chartex") != std::string::npos;
                if (chartEx) {
                    xlpp::Chart chart(parseChartExType(chartText));
                    populateChartExModel(chart, chartText);
                    chart.setThemePalette(parseChartThemePalette(z));
                    chart.setStyleResources(parseChartStyleResources(z, chartPart));
                    const auto anchorInfo = parseChartAnchorInfo(anchorNode, anchorType, frame);
                    chart.setAnchorInfo(anchorInfo);
                    if (anchorInfo.widthEmu > 0) chart.setWidth(std::max(1, static_cast<int>(std::llround(static_cast<double>(anchorInfo.widthEmu) / 9525.0))));
                    if (anchorInfo.heightEmu > 0) chart.setHeight(std::max(1, static_cast<int>(std::llround(static_cast<double>(anchorInfo.heightEmu) / 9525.0))));
                    const auto nonVisual = drawingTags(frame, "xdr:cNvPr", "cNvPr");
                    std::string objectId; std::string objectName = "Chart";
                    if (!nonVisual.empty()) { objectId = xlpp::internal::attribute(nonVisual.front(), "id"); const auto parsedName=xlpp::internal::attribute(nonVisual.front(),"name"); if(!parsedName.empty()) objectName=parsedName; }
                    chart.setStableId(drawingPart + "#" + (objectId.empty() ? chartRelationshipId : objectId));
                    chart.setSourceDrawingPart(drawingPart); chart.setSourceChartPart(chartPart); chart.setSourceRelationshipId(chartRelationshipId); chart.setDrawingObjectName(objectName); chart.setImported(true);
                    ws.addLoadedChart(std::move(chart));
                    continue;
                }

                auto plots = parseChartPlots(chartText);
                const auto primaryType = plots.empty() ? parseChartType(chartText) : plots.front().type;
                xlpp::Chart chart(primaryType);
                chart.setGrouping(plots.empty() ? parseChartGrouping(chartText, chart.type()) : plots.front().grouping);
                chart.setThemePalette(parseChartThemePalette(z));
                chart.setStyleResources(parseChartStyleResources(z, chartPart));
                chart.setTitle(chartTitleText(chartText));
                const auto chartTitleNodes = drawingTags(chartText, "c:title", "title");
                if (!chartTitleNodes.empty()) chart.setTitleRichText(parseChartRichText(chartTitleNodes.front()));
                chart.setView3D(parseChartView3D(chartText));
                chart.setFloorFormat(parseChartWallFormat(chartText, "c:floor", "floor"));
                chart.setSideWallFormat(parseChartWallFormat(chartText, "c:sideWall", "sideWall"));
                chart.setBackWallFormat(parseChartWallFormat(chartText, "c:backWall", "backWall"));
                auto axes = parseChartAxes(chartText, plots);
                std::uint64_t primaryXAxisId = 0;
                std::uint64_t primaryYAxisId = 0;
                if (!plots.empty() && plots.front().axisIds.size() >= 2) {
                    primaryXAxisId = plots.front().axisIds[0];
                    primaryYAxisId = plots.front().axisIds[1];
                }
                for (auto& plot : plots) {
                    plot.usesSecondaryAxes = std::any_of(plot.axisIds.begin(), plot.axisIds.end(), [&](std::uint64_t axisId) {
                        const auto it = std::find_if(axes.begin(), axes.end(), [&](const auto& axis) { return axis.id == axisId; });
                        return it != axes.end() && it->secondary;
                    });
                }
                chart.setPrimaryAxisIds(primaryXAxisId, primaryYAxisId);
                chart.setXAxisTitle(primaryXAxisId != 0 ? axisTitleById(axes, primaryXAxisId) : std::string{});
                chart.setYAxisTitle(primaryYAxisId != 0 ? axisTitleById(axes, primaryYAxisId) : std::string{});
                if (primaryXAxisId == 0 || primaryYAxisId == 0) {
                    const bool xyValueAxes = chart.type() == xlpp::Chart::Type::Scatter || chart.type() == xlpp::Chart::Type::Bubble;
                    if (primaryXAxisId == 0) chart.setXAxisTitle(axisTitleText(chartText, xyValueAxes ? "c:valAx" : "c:catAx",
                                                                               xyValueAxes ? "valAx" : "catAx", 0));
                    if (primaryYAxisId == 0) chart.setYAxisTitle(axisTitleText(chartText, "c:valAx", "valAx", xyValueAxes ? 1 : 0));
                }
                chart.setAxes(axes);
                chart.setPlots(plots);
                const auto plotAreas = drawingTags(chartText, "c:plotArea", "plotArea");
                if (!plotAreas.empty()) {
                    chart.setPlotAreaLayout(parseChartManualLayout(plotAreas.front()));
                    chart.setDataTable(parseChartDataTable(plotAreas.front()));
                    const auto plotSpPr = plotAreaDirectSpPr(plotAreas.front());
                    if (!plotSpPr.empty()) { chart.setPlotAreaLineFormat(parseChartLineFormat(plotSpPr)); chart.setPlotAreaFillFormat(parseChartFillFormat(plotSpPr)); }
                }
                const auto chartAreaSpPr = chartSpaceDirectSpPr(chartText);
                if (!chartAreaSpPr.empty()) { chart.setChartAreaLineFormat(parseChartLineFormat(chartAreaSpPr)); chart.setChartAreaFillFormat(parseChartFillFormat(chartAreaSpPr)); }
                const auto styleNodes = drawingTags(chartText, "c:style", "style");
                if (!styleNodes.empty()) chart.setStyle(xlpp::internal::attribute(styleNodes.front(), "val"));
                const auto legendNodes = drawingTags(chartText, "c:legend", "legend");
                chart.setShowLegend(!legendNodes.empty());
                if (!legendNodes.empty()) {
                    const auto& legendXml = legendNodes.front();
                    const auto positions = drawingTags(legendXml, "c:legendPos", "legendPos");
                    if (!positions.empty()) chart.setLegendPosition(xlpp::internal::attribute(positions.front(), "val"));
                    xlpp::ChartLegendFormat legendFormat; legendFormat.present = true;
                    legendFormat.overlay = chartBoolValue(legendXml, "c:overlay", "overlay");
                    legendFormat.layout = parseChartManualLayout(legendXml);
                    const auto spPr = drawingTags(legendXml, "c:spPr", "spPr");
                    if (!spPr.empty()) { legendFormat.line = parseChartLineFormat(spPr.front()); legendFormat.fill = parseChartFillFormat(spPr.front()); }
                    chart.setLegendFormat(std::move(legendFormat));
                }
                for (const auto& seriesNode : drawingTags(chartText, "c:ser", "ser")) {
                    xlpp::ChartSeries series;
                    const auto txNodes = drawingTags(seriesNode, "c:tx", "tx");
                    if (!txNodes.empty()) {
                        auto title = drawingTagText(txNodes.front(), "c:v", "v");
                        const auto titleReference = drawingTagText(txNodes.front(), "c:f", "f");
                        auto titleCache = parseChartSeriesCache(txNodes.front());
                        if (title.empty() && titleCache.present && !titleCache.points.empty()) title = titleCache.points.front().value;
                        if (title.empty()) title = titleReference;
                        series.setTitle(title);
                        series.setTitleReference(titleReference);
                        series.setTitleCache(std::move(titleCache));
                    }
                    const auto valNodes = drawingTags(seriesNode, "c:val", "val");
                    if (!valNodes.empty()) { series.setValuesReference(drawingTagText(valNodes.front(), "c:f", "f")); series.setValuesCache(parseChartSeriesCache(valNodes.front())); }
                    else {
                        const auto yValNodes = drawingTags(seriesNode, "c:yVal", "yVal");
                        if (!yValNodes.empty()) { series.setValuesReference(drawingTagText(yValNodes.front(), "c:f", "f")); series.setValuesCache(parseChartSeriesCache(yValNodes.front())); }
                    }
                    const auto catNodes = drawingTags(seriesNode, "c:cat", "cat");
                    if (!catNodes.empty()) { series.setCategoriesReference(drawingTagText(catNodes.front(), "c:f", "f")); series.setCategoriesCache(parseChartSeriesCache(catNodes.front())); }
                    else {
                        const auto xValNodes = drawingTags(seriesNode, "c:xVal", "xVal");
                        if (!xValNodes.empty()) { series.setCategoriesReference(drawingTagText(xValNodes.front(), "c:f", "f")); series.setCategoriesCache(parseChartSeriesCache(xValNodes.front())); }
                    }
                    const auto bubbleSizeNodes = drawingTags(seriesNode, "c:bubbleSize", "bubbleSize");
                    if (!bubbleSizeNodes.empty()) { series.setBubbleSizeReference(drawingTagText(bubbleSizeNodes.front(), "c:f", "f")); series.setBubbleSizeCache(parseChartSeriesCache(bubbleSizeNodes.front())); }
                    const auto smoothNodes = drawingTags(seriesNode, "c:smooth", "smooth");
                    if (!smoothNodes.empty()) { const auto v=xlpp::internal::attribute(smoothNodes.front(),"val"); series.setSmooth(v=="1"||v=="true"||v=="True"); }
                    series.setDataLabels(parseChartDataLabels(seriesNode));
                    series.setTrendlines(parseChartTrendlines(seriesNode));
                    series.setErrorBars(parseChartErrorBars(seriesNode));
                    const auto seriesSpPr = seriesDirectSpPr(seriesNode);
                    if (!seriesSpPr.empty()) {
                        series.setLineFormat(parseChartLineFormat(seriesSpPr));
                        series.setFillFormat(parseChartFillFormat(seriesSpPr));
                    }
                    series.setMarkerFormat(parseChartMarkerFormat(seriesNode));
                    series.setDataPoints(parseChartDataPoints(seriesNode));
                    chart.addSeries(std::move(series));
                }

                const auto anchorInfo = parseChartAnchorInfo(anchorNode, anchorType, frame);
                chart.setAnchorInfo(anchorInfo);
                if (anchorInfo.widthEmu > 0) chart.setWidth(std::max(1, static_cast<int>(std::llround(static_cast<double>(anchorInfo.widthEmu) / 9525.0))));
                if (anchorInfo.heightEmu > 0) chart.setHeight(std::max(1, static_cast<int>(std::llround(static_cast<double>(anchorInfo.heightEmu) / 9525.0))));
                const auto nonVisual = drawingTags(frame, "xdr:cNvPr", "cNvPr");
                std::string objectId;
                std::string objectName = "Chart";
                if (!nonVisual.empty()) {
                    objectId = xlpp::internal::attribute(nonVisual.front(), "id");
                    const auto parsedName = xlpp::internal::attribute(nonVisual.front(), "name");
                    if (!parsedName.empty()) objectName = parsedName;
                }
                chart.setStableId(drawingPart + "#" + (objectId.empty() ? chartRelationshipId : objectId));
                chart.setSourceDrawingPart(drawingPart);
                chart.setSourceChartPart(chartPart);
                chart.setSourceRelationshipId(chartRelationshipId);
                chart.setDrawingObjectName(objectName);
                chart.setImported(true);
                ws.addLoadedChart(std::move(chart));
            }
        };

        loadAnchors("xdr:oneCellAnchor", "oneCellAnchor", xlpp::DrawingAnchorType::OneCell);
        loadAnchors("xdr:twoCellAnchor", "twoCellAnchor", xlpp::DrawingAnchorType::TwoCell);
        loadAnchors("xdr:absoluteAnchor", "absoluteAnchor", xlpp::DrawingAnchorType::Absolute);
    }
}


} // namespace xlpp::internal::ooxml
