#include "OOXML/Charts/ImportedChartEditorParts.h"
#include "OOXML/Charts/ChartFormatCodec.h"
#include "OOXML/Charts/ChartSerializer.h"
#include "OOXML/Charts/ChartMutationSupport.h"
#include "OOXML/Charts/ChartXmlSupport.h"
#include "OOXML/Common/DrawingXmlSupport.h"
#include "OOXML/Common/Namespaces.h"
#include "OOXML/Drawings/WorkbookDrawingAccess.h"
#include "Package/Xml/XmlUtilities.h"

#include <algorithm>
#include <array>
#include <cmath>
#include <cstdint>
#include <iomanip>
#include <limits>
#include <sstream>
#include <string>
#include <tuple>
#include <type_traits>
#include <utility>
#include <vector>

namespace xlpp::internal::ooxml {

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





} // namespace xlpp::internal::ooxml
