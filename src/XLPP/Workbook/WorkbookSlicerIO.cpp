#include "WorkbookSlicerIO.h"
#include "WorkbookNamespaces.h"
#include <XLPP/Worksheet/Slicer.h>
#include "../XML/XmlUtilities.h"
#include <sstream>
#include <string>
#include <vector>

namespace xlpp {
namespace internal {

std::string insertSlicerListExt(std::string sheetXml, const std::vector<std::string>& relationshipIds,
                                bool strict) {
    const auto x14 = strict
        ? "http://purl.oclc.org/ooxml/spreadsheetml/2009/9/main"
        : "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main";
    const auto marker = std::string("</worksheet>");
    const auto position = sheetXml.find(marker);
    if (position == std::string::npos) return sheetXml;
    std::string ext = std::string("<extLst><ext uri=\"{A8765BA9-456A-4dab-B4F3-ACF838C121DE}\" xmlns:x14=\"")
        + x14 + "\"><x14:slicerList>";
    for (const auto& relId : relationshipIds)
        ext += "<x14:slicer r:id=\"" + xmlEscape(relId) + "\"/>";
    ext += "</x14:slicerList></ext></extLst>";
    sheetXml.insert(position, ext);
    return sheetXml;
}

std::string insertWorkbookSlicerCachesExt(std::string workbookXml,
                                          const std::vector<std::string>& relationshipIds,
                                          bool strict) {
    const auto x14 = strict
        ? "http://purl.oclc.org/ooxml/spreadsheetml/2009/9/main"
        : "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main";
    const auto marker = std::string("</workbook>");
    const auto position = workbookXml.find(marker);
    if (position == std::string::npos) return workbookXml;
    std::string ext = std::string("<extLst><ext uri=\"{A8765BA9-456A-4dab-B4F3-ACF838C121DE}\" xmlns:x14=\"")
        + x14 + "\"><x14:slicerCaches>";
    for (const auto& relId : relationshipIds)
        ext += "<x14:slicerCache r:id=\"" + xmlEscape(relId) + "\"/>";
    ext += "</x14:slicerCaches></ext></extLst>";
    workbookXml.insert(position, ext);
    return workbookXml;
}

std::string slicerCacheRelationshipId(std::size_t index) {
    return "rIdSlicerCache" + std::to_string(index + 1);
}

std::string slicerRelationshipId(std::size_t index) {
    return "rIdSlicer" + std::to_string(index + 1);
}

std::string slicerCacheXml(const xlpp::Slicer& slicer) {
    std::ostringstream xml;
    xml << R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>)"
        << R"(<slicerCacheDefinition xmlns="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main")"
        << R"( xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision")"
        << " name=\"" << xmlEscape(slicer.name) << "\" sourceName=\"" << xmlEscape(slicer.sourceName) << "\">";
    if (!slicer.pivotTableName.empty()) {
        xml << "<pivotTables><pivotTable name=\"" << xmlEscape(slicer.pivotTableName) << "\"/></pivotTables>";
    }
    xml << "<data><slicerCacheData>";
    for (std::size_t i = 0; i < slicer.items.size(); ++i)
        xml << "<item x=\"" << i << "\" s=\"" << (slicer.items[i].selected ? 1 : 0) << "\"/>";
    xml << "</slicerCacheData></data></slicerCacheDefinition>";
    return xml.str();
}

std::string slicerXml(const xlpp::Slicer& slicer) {
    std::ostringstream xml;
    xml << R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>)"
        << R"(<slicer xmlns="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main")"
        << " name=\"" << xmlEscape(slicer.name) << "\" cache=\"" << xmlEscape(slicer.name)
        << "\" caption=\"" << xmlEscape(slicer.caption.empty() ? slicer.sourceName : slicer.caption)
        << "\" columnCount=\"" << slicer.columnCount << "\" style=\"" << xmlEscape(slicer.style) << "\"/>";
    return xml.str();
}

} // namespace internal
} // namespace xlpp
