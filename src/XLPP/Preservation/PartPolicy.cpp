#include "Preservation/PartPolicy.h"
#include <algorithm>
#include <cctype>
namespace xlpp::internal::preservation {

bool isRegeneratedPart(const std::string& name) {
    return name == "[Content_Types].xml" || name == "_rels/.rels"
        || name == "docProps/core.xml" || name == "docProps/app.xml" || name == "docProps/custom.xml"
        || name == "xl/workbook.xml" || name == "xl/_rels/workbook.xml.rels"
        || name == "xl/styles.xml" || name == "xl/sharedStrings.xml"
        || name.rfind("xl/worksheets/", 0) == 0;
}

std::string extensionOf(const std::string& name) {
    const auto dot = name.rfind('.');
    if (dot == std::string::npos || dot + 1 == name.size()) return {};
    return name.substr(dot + 1);
}

} // namespace xlpp::internal::preservation
