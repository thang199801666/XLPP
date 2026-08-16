#pragma once
#include <string>

namespace xlpp::internal {

// Namespace URI helpers shared across the workbook writer modules. `strict`
// selects the ISO/IEC 29500 Strict namespace set.
inline std::string nsMain(bool strict) {
    return strict ? std::string("http://purl.oclc.org/ooxml/spreadsheetml/main")
                  : std::string("http://schemas.openxmlformats.org/spreadsheetml/2006/main");
}
inline std::string nsRelsPkg(bool strict) {
    return strict ? std::string("http://purl.oclc.org/ooxml/package/relationships")
                  : std::string("http://schemas.openxmlformats.org/package/2006/relationships");
}
inline std::string nsRelsDoc(bool strict) {
    return strict ? std::string("http://purl.oclc.org/ooxml/officeDocument/relationships")
                  : std::string("http://schemas.openxmlformats.org/officeDocument/2006/relationships");
}
inline std::string nsCtPkg(bool strict) {
    return strict ? std::string("http://purl.oclc.org/ooxml/package/content-types")
                  : std::string("http://schemas.openxmlformats.org/package/2006/content-types");
}
inline std::string nsCoreProps(bool strict) {
    return strict ? std::string("http://purl.oclc.org/ooxml/package/metadata/core-properties")
                  : std::string("http://schemas.openxmlformats.org/package/2006/metadata/core-properties");
}
inline std::string nsExtProps(bool strict) {
    return strict ? std::string("http://purl.oclc.org/ooxml/officeDocument/extended-properties")
                  : std::string("http://schemas.openxmlformats.org/officeDocument/2006/extended-properties");
}
inline std::string nsVTypes(bool strict) {
    return strict ? std::string("http://purl.oclc.org/ooxml/officeDocument/docPropsVTypes")
                  : std::string("http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes");
}

} // namespace xlpp::internal
