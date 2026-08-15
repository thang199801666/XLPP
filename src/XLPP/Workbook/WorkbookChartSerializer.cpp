#include "WorkbookChartSerializer.h"
#include "../XML/XmlUtilities.h"

#include <algorithm>
#include <sstream>
#include <string>
#include <string_view>
#include <vector>

namespace xlpp {
namespace {

std::string xmlEscape(const std::string& value) { return xlpp::internal::xmlEscape(value); }

} // namespace
} // namespace xlpp

namespace xlpp::internal {

std::vector<std::string> drawingTags(const std::string& xml, const char* prefixed, const char* local) {
    auto result = xlpp::internal::tags(xml, prefixed);
    if (std::string_view(prefixed) != std::string_view(local)) {
        auto unprefixed = xlpp::internal::tags(xml, local);
        result.insert(result.end(), std::make_move_iterator(unprefixed.begin()), std::make_move_iterator(unprefixed.end()));
    }
    // A preserved drawing can mix producer-native default-namespace elements
    // with xdr:-prefixed nodes appended by XL++. Preserve document order among
    // equivalent spellings when both forms occur in the same anchor family.
    std::stable_sort(result.begin(), result.end(), [&](const auto& lhs, const auto& rhs) {
        return xml.find(lhs) < xml.find(rhs);
    });
    return result;
}

std::string drawingTagText(const std::string& xml, const char* prefixed, const char* local) {
    auto value = xlpp::internal::tagText(xml, prefixed);
    if (value.empty()) value = xlpp::internal::tagText(xml, local);
    return value;
}

bool generatedChartTypeUsesXYAxes(xlpp::Chart::Type type) {
    return type == xlpp::Chart::Type::Scatter || type == xlpp::Chart::Type::Bubble;
}

bool generatedChartTypeHasAxes(xlpp::Chart::Type type) {
    return type != xlpp::Chart::Type::Pie && type != xlpp::Chart::Type::Pie3D &&
           type != xlpp::Chart::Type::Doughnut && type != xlpp::Chart::Type::PieOfPie &&
           type != xlpp::Chart::Type::BarOfPie;
}

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

std::string chartView3DXml(const xlpp::ChartView3D& view, bool prefixed) {
    const auto c = prefixed ? "c:" : "";
    std::ostringstream xml;
    xml << "<" << c << "view3D>";
    if (view.hasRotationX) xml << "<" << c << "rotX val=\"" << view.rotationX << "\"/>";
    if (view.hasHeightPercent) xml << "<" << c << "hPercent val=\"" << view.heightPercent << "\"/>";
    if (view.hasRotationY) xml << "<" << c << "rotY val=\"" << view.rotationY << "\"/>";
    if (view.hasDepthPercent) xml << "<" << c << "depthPercent val=\"" << view.depthPercent << "\"/>";
    if (view.hasRightAngleAxes) xml << "<" << c << "rAngAx val=\"" << (view.rightAngleAxes ? 1 : 0) << "\"/>";
    if (view.hasPerspective) xml << "<" << c << "perspective val=\"" << view.perspective << "\"/>";
    xml << "</" << c << "view3D>";
    return xml.str();
}

} // namespace xlpp::internal
