#include "Preservation/XmlMergeSupport.h"
#include "Package/Xml/XmlUtilities.h"

namespace xlpp::internal::preservation {

std::vector<std::string> extractTagBlocks(const std::string& xml, const std::string& tag) {
    return xlpp::internal::tags(xml, tag);
}

void eraseTagBlocks(std::string& xml, const std::string& tag) {
    for (const auto& block : extractTagBlocks(xml, tag)) {
        std::size_t position = 0;
        while ((position = xml.find(block, position)) != std::string::npos)
            xml.erase(position, block.size());
    }
}

std::string joinBlocks(const std::vector<std::string>& blocks) {
    std::string result;
    for (const auto& block : blocks) result += block;
    return result;
}

void insertBefore(std::string& xml, const std::string& marker, const std::string& content) {
    if (content.empty()) return;
    const auto position = xml.find(marker);
    if (position == std::string::npos) return;
    xml.insert(position, content);
}



} // namespace xlpp::internal::preservation
