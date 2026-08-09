#include "OOXML/Charts/ChartXmlSupport.h"
#include "OOXML/Common/DrawingXmlSupport.h"
#include <algorithm>
#include <array>
#include <utility>
#include <vector>

namespace xlpp::internal::ooxml {

std::string chartSpaceDirectSpPr(const std::string& chartXmlText) {
    const auto candidates = drawingTags(chartXmlText, "c:spPr", "spPr");
    if (candidates.empty()) return {};
    const auto charts = drawingTags(chartXmlText, "c:chart", "chart");
    for (const auto& candidate : candidates)
        if (std::none_of(charts.begin(), charts.end(), [&](const auto& chart) { return chart.find(candidate) != std::string::npos; }))
            return candidate;
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
        const auto nodes = drawingTags(plotArea, pair.first, pair.second);
        nested.insert(nested.end(), nodes.begin(), nodes.end());
    }
    for (const auto& candidate : candidates)
        if (std::none_of(nested.begin(), nested.end(), [&](const auto& owner) { return owner.find(candidate) != std::string::npos; }))
            return candidate;
    return {};
}

std::string seriesDirectSpPr(const std::string& seriesXml) {
    const auto candidates = drawingTags(seriesXml, "c:spPr", "spPr");
    if (candidates.empty()) return {};
    std::vector<std::string> nested;
    for (const auto& pair : std::array<std::pair<const char*, const char*>, 5>{
             std::pair{"c:marker", "marker"}, {"c:dPt", "dPt"}, {"c:dLbls", "dLbls"}, {"c:trendline", "trendline"}, {"c:errBars", "errBars"}}) {
        const auto nodes = drawingTags(seriesXml, pair.first, pair.second);
        nested.insert(nested.end(), nodes.begin(), nodes.end());
    }
    for (const auto& candidate : candidates)
        if (std::none_of(nested.begin(), nested.end(), [&](const auto& node) { return node.find(candidate) != std::string::npos; }))
            return candidate;
    return {};
}

std::string axisDirectSpPr(const std::string& axisXml) {
    const auto candidates = drawingTags(axisXml, "c:spPr", "spPr");
    if (candidates.empty()) return {};
    std::vector<std::string> nested;
    for (const auto& pair : std::array<std::pair<const char*, const char*>, 5>{{
             {"c:title","title"},{"c:majorGridlines","majorGridlines"},{"c:minorGridlines","minorGridlines"},{"c:txPr","txPr"},{"c:extLst","extLst"}}}) {
        const auto nodes = drawingTags(axisXml, pair.first, pair.second);
        nested.insert(nested.end(), nodes.begin(), nodes.end());
    }
    for (const auto& candidate : candidates)
        if (std::none_of(nested.begin(), nested.end(), [&](const auto& node) { return node.find(candidate) != std::string::npos; }))
            return candidate;
    return {};
}

} // namespace xlpp::internal::ooxml
