#include "SharedStringsReader.h"
#include "../XML/XmlScanner.h"
#include "../XML/XmlUtilities.h"
#include <algorithm>
#include <utility>

namespace xlpp::internal {

SharedStringsReader::SharedStringsReader(ZipArchiveReader archive) : archive_(std::move(archive)) {}

void SharedStringsReader::load() const {
    std::call_once(loaded_, [this] {
        if (!archive_.contains("xl/sharedStrings.xml")) return;
        const auto xml = archive_.readEntry("xl/sharedStrings.xml");

        // sharedStrings.xml can be one of the largest parts in text-heavy
        // workbooks.  Scan it through string_views instead of materializing a
        // vector<string> for every <si> and then another vector<string> for
        // every nested <t> run.
        XmlScanner rootScanner(xml);
        std::string_view root;
        if (rootScanner.nextElement("sst", root)) {
            std::size_t declared = 0;
            if (parseSize(xmlAttribute(root, "uniqueCount"), declared)) {
                // Treat the attribute as a hint only.  Bound it by the actual
                // XML size so a malformed package cannot force an absurd
                // reserve before any entries have been validated.
                const auto plausibleMax = xml.size() / 9u + 1u;
                strings_.reserve(std::min(declared, plausibleMax));
            }
        }

        XmlScanner items(xml);
        std::string_view si;
        while (items.nextElement("si", si)) {
            XmlScanner texts(si);
            std::string_view textElement;
            std::string value;
            bool foundText = false;
            while (texts.nextElement("t", textElement)) {
                foundText = true;
                const auto open = textElement.find('>');
                const auto close = textElement.rfind("</");
                if (open == std::string_view::npos || close == std::string_view::npos || close < open)
                    continue;
                const auto text = textElement.substr(open + 1, close - open - 1);
                if (containsEntity(text)) value += xmlUnescape(text);
                else value.append(text.data(), text.size());
            }
            if (!foundText) value.clear();
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
