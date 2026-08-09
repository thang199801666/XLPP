#pragma once

#include <string>

namespace xlpp::internal::ooxml {

inline std::string nsMain(bool strict) {
    return strict ? "http://purl.oclc.org/ooxml/spreadsheetml/main"
                  : "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
}
inline std::string nsRelsPkg(bool strict) {
    return strict ? "http://purl.oclc.org/ooxml/package/relationships"
                  : "http://schemas.openxmlformats.org/package/2006/relationships";
}
inline std::string nsRelsDoc(bool strict) {
    return strict ? "http://purl.oclc.org/ooxml/officeDocument/relationships"
                  : "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
}
inline std::string nsCtPkg(bool strict) {
    return strict ? "http://purl.oclc.org/ooxml/package/content-types"
                  : "http://schemas.openxmlformats.org/package/2006/content-types";
}
inline std::string nsCoreProps(bool strict) {
    return strict ? "http://purl.oclc.org/ooxml/package/metadata/core-properties"
                  : "http://schemas.openxmlformats.org/package/2006/metadata/core-properties";
}
inline std::string nsExtProps(bool strict) {
    return strict ? "http://purl.oclc.org/ooxml/officeDocument/extended-properties"
                  : "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties";
}
inline std::string nsVTypes(bool strict) {
    return strict ? "http://purl.oclc.org/ooxml/officeDocument/docPropsVTypes"
                  : "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes";
}

} // namespace xlpp::internal::ooxml
