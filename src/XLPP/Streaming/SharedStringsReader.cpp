#include "SharedStringsReader.h"
#include "Package/Xml/XmlUtilities.h"
#include <utility>

namespace {

// Extracts the text of a single <t> element, e.g. "<t xml:space="preserve">Hi</t>".
std::string inlineText(std::string_view element) {
    const auto open = element.find('>');
    const auto close = element.find("</");
    if (open == std::string_view::npos || close == std::string_view::npos || close < open) return {};
    return xlpp::internal::xmlUnescape(element.substr(open + 1, close - open - 1));
}

} // namespace

namespace xlpp::internal {

SharedStringsReader::SharedStringsReader(ZipArchiveReader archive) : archive_(std::move(archive)) {}

void SharedStringsReader::load() const {
    std::call_once(loaded_, [this] {
        if (!archive_.contains("xl/sharedStrings.xml")) return;
        const auto xml = archive_.readEntry("xl/sharedStrings.xml");
        const auto items = tags(xml, "si");
        strings_.reserve(items.size());
        for (const auto& si : items) {
            const auto texts = tags(si, "t");
            if (texts.empty()) {
                strings_.emplace_back();
                continue;
            }
            std::string value;
            for (const auto& text : texts) value += inlineText(text);
            strings_.push_back(std::move(value));
        }
    });
}

const std::string* SharedStringsReader::lookup(std::size_t index) const {
    load();
    return index < strings_.size() ? &strings_[index] : nullptr;
}

std::size_t SharedStringsReader::size() const {
    load();
    return strings_.size();
}

} // namespace xlpp::internal
