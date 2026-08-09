#include <XLPP/Workbook/Workbook.h>

#include <algorithm>
#include <cctype>
#include <string_view>

namespace xlpp {
namespace {
std::string lowerAscii(std::string value) {
    std::transform(value.begin(), value.end(), value.begin(), [](unsigned char c) { return static_cast<char>(std::tolower(c)); });
    return value;
}
bool containsNoCase(std::string_view text, std::string_view needle) {
    auto lhs = lowerAscii(std::string(text)); auto rhs = lowerAscii(std::string(needle));
    return lhs.find(rhs) != std::string::npos;
}
std::string attributeValue(std::string_view tag, std::string_view name) {
    const auto marker = std::string(name) + "=\"";
    const auto begin = tag.find(marker); if (begin == std::string_view::npos) return {};
    const auto valueBegin = begin + marker.size(); const auto end = tag.find('"', valueBegin);
    return end == std::string_view::npos ? std::string{} : std::string(tag.substr(valueBegin, end - valueBegin));
}
std::vector<std::string> tagAttributes(std::string_view xml, std::string_view tagName, std::string_view attribute) {
    std::vector<std::string> values; const std::string marker = "<" + std::string(tagName);
    std::size_t pos = 0;
    while ((pos = xml.find(marker, pos)) != std::string_view::npos) {
        const auto end = xml.find('>', pos); if (end == std::string_view::npos) break;
        auto value = attributeValue(xml.substr(pos, end - pos + 1), attribute); if (!value.empty()) values.push_back(std::move(value));
        pos = end + 1;
    }
    return values;
}
bool attrBool(std::string_view tag, std::string_view name) {
    const auto value = attributeValue(tag, name); return value == "1" || value == "true" || value == "TRUE";
}
}

ExternalDataInspection Workbook::inspectExternalData() const {
    ExternalDataInspection out;
    for (const auto& part : preservedParts_) {
        const auto path = lowerAscii(part.name);
        const auto type = lowerAscii(part.overrideType);
        if (path.find("externallinks/externallink") != std::string::npos) {
            ExternalWorkbookLinkInfo link; link.partName = part.name;
            link.sheetNames = tagAttributes(part.data, "sheetName", "val");
            link.definedNames = tagAttributes(part.data, "definedName", "name");
            out.externalWorkbooks.push_back(std::move(link));
        } else if (path == "xl/connections.xml" || type.find("connections+xml") != std::string::npos) {
            std::size_t pos = 0;
            while ((pos = part.data.find("<connection", pos)) != std::string::npos) {
                const auto next = pos + 11;
                if (next < part.data.size() && (part.data[next] == 's' || part.data[next] == ':')) { pos = next; continue; }
                const auto end = part.data.find('>', pos); if (end == std::string::npos) break;
                const auto tag = std::string_view(part.data).substr(pos, end - pos + 1);
                WorkbookConnectionInfo connection; connection.partName = part.name;
                connection.id = attributeValue(tag, "id"); connection.name = attributeValue(tag, "name");
                connection.description = attributeValue(tag, "description"); connection.type = attributeValue(tag, "type");
                connection.refreshOnLoad = attrBool(tag, "refreshOnLoad"); connection.background = attrBool(tag, "background");
                connection.deleted = attrBool(tag, "deleted"); out.connections.push_back(std::move(connection)); pos = end + 1;
            }
        } else if (path.find("querytables/querytable") != std::string::npos || type.find("querytable+xml") != std::string::npos) {
            QueryTableInfo query; query.partName = part.name;
            const auto begin = part.data.find("<queryTable"); const auto end = begin == std::string::npos ? std::string::npos : part.data.find('>', begin);
            if (end != std::string::npos) {
                const auto tag = std::string_view(part.data).substr(begin, end - begin + 1);
                query.name = attributeValue(tag, "name"); query.connectionId = attributeValue(tag, "connectionId");
                query.refreshOnLoad = attrBool(tag, "refreshOnLoad");
            }
            out.queryTables.push_back(std::move(query));
        }
        if (path.find("customxml") != std::string::npos && (containsNoCase(part.data, "power query") || containsNoCase(part.data, "mashup")))
            out.powerQueryParts.push_back(part.name);
        if (path.find("queries/") != std::string::npos || type.find("query") != std::string::npos) {
            if (std::find(out.powerQueryParts.begin(), out.powerQueryParts.end(), part.name) == out.powerQueryParts.end())
                out.unknownConnectionParts.push_back(part.name);
        }
        if (containsNoCase(part.data, "webPr") || type.find("querytable") != std::string::npos)
            out.webQueryParts.push_back(part.name);
    }
    return out;
}

DataModelInspection Workbook::inspectDataModel() const {
    DataModelInspection out;
    for (const auto& part : preservedParts_) {
        const auto path = lowerAscii(part.name); const auto type = lowerAscii(part.overrideType);
        const bool modelPart = path.find("xl/model/") != std::string::npos || path.find("datamodel") != std::string::npos ||
                               type.find("model") != std::string::npos || type.find("powerpivot") != std::string::npos;
        if (modelPart) { out.present = true; out.modelParts.push_back(part.name); }
        if (path.find("pivotcache/pivotcachedefinition") != std::string::npos &&
            (containsNoCase(part.data, "cacheSource type=\"external\"") || containsNoCase(part.data, "<olapPr"))) {
            out.hasOlapPivotCaches = true; out.olapPivotCacheParts.push_back(part.name);
        }
    }
    for (const auto& rel : preservedRelationships_) {
        if (containsNoCase(rel.type, "model") || containsNoCase(rel.type, "connection") || containsNoCase(rel.target, "model")) {
            if (containsNoCase(rel.type, "model") || containsNoCase(rel.target, "model")) out.present = true;
            out.modelRelationships.push_back(rel.sourcePart + "|" + rel.id + "|" + rel.type + "|" + rel.target);
        }
    }
    if (out.present && out.modelParts.empty()) out.warnings.push_back("Data Model relationship detected without an inspectable model part; preservation-only mode applies");
    if (out.hasOlapPivotCaches) out.warnings.push_back("OLAP/Data Model pivot caches are preservation-only and are not regenerated by XL++");
    return out;
}

} // namespace xlpp
