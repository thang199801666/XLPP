#pragma once

#include <XLPP/Cell/RichText.h>

#include <optional>
#include <sstream>
#include <string>
#include <string_view>

namespace xlpp::internal::ooxml {

struct LoadedSharedString {
    std::string plainText;
    std::optional<xlpp::RichText> richText;
};

std::optional<xlpp::RichText> parseRichTextRuns(std::string_view container);
void writeRichTextRuns(std::ostringstream& xml, const xlpp::RichText& richText);

} // namespace xlpp::internal::ooxml
