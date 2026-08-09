#include "OOXML/Common/DrawingXmlSupport.h"
#include "Package/Xml/XmlUtilities.h"

#include <algorithm>
#include <array>
#include <cctype>
#include <iterator>
#include <string_view>
#include <utility>

namespace xlpp::internal::ooxml {

std::vector<std::string> drawingTags(const std::string& xml, const char* prefixed, const char* local) {
    struct PositionedTag {
        std::size_t position{};
        std::size_t discoveryOrder{};
        std::string xml;
    };

    std::vector<PositionedTag> positioned;
    std::size_t discoveryOrder = 0;
    auto append = [&](std::vector<std::string> tags) {
        positioned.reserve(positioned.size() + tags.size());
        for (auto& tag : tags) {
            // Cache the source position exactly once.  The previous comparator
            // called xml.find() repeatedly during sort, turning large drawings
            // into an avoidable O(n log n * document-scan) hot path.
            positioned.push_back({xml.find(tag), discoveryOrder++, std::move(tag)});
        }
    };

    append(xlpp::internal::tags(xml, prefixed));
    if (std::string_view(prefixed) != std::string_view(local))
        append(xlpp::internal::tags(xml, local));

    std::sort(positioned.begin(), positioned.end(), [](const auto& lhs, const auto& rhs) {
        if (lhs.position != rhs.position) return lhs.position < rhs.position;
        return lhs.discoveryOrder < rhs.discoveryOrder;
    });

    std::vector<std::string> result;
    result.reserve(positioned.size());
    for (auto& tag : positioned) result.push_back(std::move(tag.xml));
    return result;
}

std::string drawingTagText(const std::string& xml, const char* prefixed, const char* local) {
    auto value = xlpp::internal::tagText(xml, prefixed);
    if (value.empty()) value = xlpp::internal::tagText(xml, local);
    return value;
}

long long drawingInteger(const std::string& xml, const char* prefixed, const char* local, long long fallback) {
    const auto value = drawingTagText(xml, prefixed, local);
    if (value.empty()) return fallback;
    try { return std::stoll(value); } catch (...) { return fallback; }
}

std::string partExtension(const std::string& part) {
    const auto slash = part.find_last_of('/');
    const auto dot = part.find_last_of('.');
    if (dot == std::string::npos || (slash != std::string::npos && dot < slash)) return {};
    auto extension = part.substr(dot + 1);
    std::transform(extension.begin(), extension.end(), extension.begin(), [](unsigned char c) {
        return static_cast<char>(std::tolower(c));
    });
    if (extension == "jpeg") extension = "jpg";
    return extension;
}

xlpp::DrawingMarker parseDrawingMarker(const std::string& markerXml) {
    xlpp::DrawingMarker marker;
    marker.column = static_cast<std::size_t>(std::max<long long>(0, drawingInteger(markerXml, "xdr:col", "col"))) + 1;
    marker.row = static_cast<std::size_t>(std::max<long long>(0, drawingInteger(markerXml, "xdr:row", "row"))) + 1;
    marker.columnOffsetEmu = drawingInteger(markerXml, "xdr:colOff", "colOff");
    marker.rowOffsetEmu = drawingInteger(markerXml, "xdr:rowOff", "rowOff");
    return marker;
}

} // namespace xlpp::internal::ooxml
