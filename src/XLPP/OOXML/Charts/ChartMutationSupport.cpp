#include "OOXML/Charts/ChartMutationSupport.h"
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
                           const std::string& value, bool insertWhenMissing) {
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

void removeDrawingChild(std::string& container, const char* prefixed, const char* local) {
    for (;;) {
        const auto nodes = drawingTags(container, prefixed, local);
        if (nodes.empty()) return;
        const auto position = container.find(nodes.front());
        if (position == std::string::npos) return;
        container.erase(position, nodes.front().size());
    }
}

bool patchOpeningTagAttribute(std::string& node, const std::string& name, const std::string& value, bool removeWhenEmpty) {
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

} // namespace xlpp::internal::ooxml
