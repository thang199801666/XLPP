#include "OOXML/Workbook/WorkbookReferenceSupport.h"
#include <algorithm>
#include <cctype>
#include <sstream>

namespace xlpp::internal::ooxml {

std::string quoteSheetName(const std::string& name) {
    std::string escaped;
    escaped.reserve(name.size() + 2);
    escaped.push_back('\'');
    for (const char ch : name) {
        escaped.push_back(ch);
        if (ch == '\'') escaped.push_back('\'');
    }
    escaped.push_back('\'');
    return escaped;
}

std::string absoluteReferenceToken(std::string token) {
    token.erase(std::remove(token.begin(), token.end(), '$'), token.end());
    if (token.empty()) return token;
    std::size_t letters = 0;
    while (letters < token.size() && std::isalpha(static_cast<unsigned char>(token[letters]))) ++letters;
    if (letters == token.size()) return "$" + token;
    if (letters == 0) return "$" + token;
    return "$" + token.substr(0, letters) + "$" + token.substr(letters);
}

std::string absoluteReference(std::string reference) {
    if (!reference.empty() && reference.front() == '=') reference.erase(reference.begin());
    const auto bang = reference.rfind('!');
    if (bang != std::string::npos) reference = reference.substr(bang + 1);
    std::ostringstream result;
    std::size_t start = 0;
    bool firstArea = true;
    while (start <= reference.size()) {
        const auto comma = reference.find(',', start);
        const auto area = reference.substr(start, comma == std::string::npos ? std::string::npos : comma - start);
        const auto colon = area.find(':');
        if (!firstArea) result << ',';
        firstArea = false;
        if (colon == std::string::npos) result << absoluteReferenceToken(area);
        else result << absoluteReferenceToken(area.substr(0, colon)) << ':' << absoluteReferenceToken(area.substr(colon + 1));
        if (comma == std::string::npos) break;
        start = comma + 1;
    }
    return result.str();
}

std::string qualifiedPrintReference(const std::string& sheetName, const std::string& reference) {
    std::ostringstream result;
    const auto absolute = absoluteReference(reference);
    std::size_t start = 0;
    bool first = true;
    while (start <= absolute.size()) {
        const auto comma = absolute.find(',', start);
        if (!first) result << ',';
        first = false;
        result << quoteSheetName(sheetName) << '!' << absolute.substr(start, comma == std::string::npos ? std::string::npos : comma - start);
        if (comma == std::string::npos) break;
        start = comma + 1;
    }
    return result.str();
}

std::string localPrintReference(std::string value) {
    std::ostringstream result;
    std::size_t start = 0;
    bool first = true;
    while (start <= value.size()) {
        const auto comma = value.find(',', start);
        auto area = value.substr(start, comma == std::string::npos ? std::string::npos : comma - start);
        const auto bang = area.rfind('!');
        if (bang != std::string::npos) area = area.substr(bang + 1);
        area.erase(std::remove(area.begin(), area.end(), '$'), area.end());
        if (!first) result << ',';
        first = false;
        result << area;
        if (comma == std::string::npos) break;
        start = comma + 1;
    }
    return result.str();
}


} // namespace xlpp::internal::ooxml
