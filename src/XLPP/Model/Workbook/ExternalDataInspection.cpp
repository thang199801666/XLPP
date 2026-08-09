#include <XLPP/Workbook/Workbook.h>

#include <algorithm>
#include <cctype>
#include <sstream>
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
std::pair<std::size_t, std::size_t> findStartTagByLocalName(std::string_view xml, std::string_view localName);
std::string xmlUnescape(std::string_view value);
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
std::vector<std::string> localTagAttributes(std::string_view xml, std::string_view localName, std::string_view attribute) {
    std::vector<std::string> values;
    std::size_t offset = 0;
    while (offset < xml.size()) {
        const auto [relativeBegin, relativeEnd] = findStartTagByLocalName(xml.substr(offset), localName);
        if (relativeEnd == std::string_view::npos) break;
        const auto begin = offset + relativeBegin;
        const auto end = offset + relativeEnd;
        if (auto value = attributeValue(xml.substr(begin, end - begin + 1), attribute); !value.empty())
            values.push_back(xmlUnescape(value));
        offset = end + 1;
    }
    return values;
}
bool attrBool(std::string_view tag, std::string_view name) {
    const auto value = attributeValue(tag, name); return value == "1" || value == "true" || value == "TRUE";
}
bool setBoolAttribute(std::string& xml, std::size_t tagBegin, std::size_t tagEnd,
                      std::string_view name, bool enabled) {
    const std::string marker = std::string(name) + "=\"";
    const auto attribute = xml.find(marker, tagBegin);
    const std::string value = enabled ? "1" : "0";
    if (attribute != std::string::npos && attribute < tagEnd) {
        const auto valueBegin = attribute + marker.size();
        const auto valueEnd = xml.find('"', valueBegin);
        if (valueEnd == std::string::npos || valueEnd > tagEnd) return false;
        if (xml.compare(valueBegin, valueEnd - valueBegin, value) == 0) return false;
        xml.replace(valueBegin, valueEnd - valueBegin, value);
        return true;
    }
    std::size_t insertAt = tagEnd;
    if (insertAt > tagBegin && xml[insertAt - 1] == '/') --insertAt;
    xml.insert(insertAt, " " + std::string(name) + "=\"" + value + "\"");
    return true;
}
std::string firstElementName(std::string_view xml) {
    auto begin = xml.find('<');
    while (begin != std::string_view::npos && begin + 1 < xml.size() &&
           (xml[begin + 1] == '?' || xml[begin + 1] == '!')) {
        begin = xml.find('<', begin + 1);
    }
    if (begin == std::string_view::npos) return {};
    const auto end = xml.find_first_of(" />\t\r\n", begin + 1);
    if (end == std::string_view::npos) return {};
    return std::string(xml.substr(begin + 1, end - begin - 1));
}
std::string normalizePackagePath(std::string_view path) {
    std::vector<std::string> segments;
    std::size_t begin = 0;
    while (begin <= path.size()) {
        const auto end = path.find('/', begin);
        const auto segment = path.substr(begin, end == std::string_view::npos ? path.size() - begin : end - begin);
        if (!segment.empty() && segment != ".") {
            if (segment == "..") {
                if (!segments.empty()) segments.pop_back();
            } else {
                segments.emplace_back(segment);
            }
        }
        if (end == std::string_view::npos) break;
        begin = end + 1;
    }
    std::string result;
    for (const auto& segment : segments) {
        if (!result.empty()) result += '/';
        result += segment;
    }
    return result;
}
std::string resolveRelationshipTarget(const PreservedRelationship& relationship) {
    if (relationship.target.empty()) return {};
    if (relationship.target.front() == '/') return normalizePackagePath(std::string_view(relationship.target).substr(1));
    const auto slash = relationship.sourcePart.find_last_of('/');
    const auto directory = slash == std::string::npos ? std::string{} : relationship.sourcePart.substr(0, slash + 1);
    return normalizePackagePath(directory + relationship.target);
}
std::pair<std::size_t, std::size_t> findStartTagByLocalName(std::string_view xml, std::string_view localName) {
    std::size_t pos = 0;
    while ((pos = xml.find('<', pos)) != std::string_view::npos) {
        ++pos;
        if (pos >= xml.size()) break;
        if (xml[pos] == '/' || xml[pos] == '?' || xml[pos] == '!') continue;
        const auto nameEnd = xml.find_first_of(" />\t\r\n", pos);
        if (nameEnd == std::string_view::npos) break;
        const auto qualifiedName = xml.substr(pos, nameEnd - pos);
        const auto colon = qualifiedName.find_last_of(':');
        const auto foundLocalName = colon == std::string_view::npos ? qualifiedName : qualifiedName.substr(colon + 1);
        if (foundLocalName == localName) {
            const auto tagEnd = xml.find('>', nameEnd);
            return {pos - 1, tagEnd};
        }
        pos = nameEnd;
    }
    return {std::string::npos, std::string::npos};
}
std::string xmlUnescape(std::string_view value) {
    std::string result(value);
    for (const auto& replacement : {std::pair{"&quot;", "\""}, std::pair{"&apos;", "'"},
                                    std::pair{"&lt;", "<"}, std::pair{"&gt;", ">"},
                                    std::pair{"&amp;", "&"}}) {
        std::size_t pos = 0;
        while ((pos = result.find(replacement.first, pos)) != std::string::npos) {
            result.replace(pos, std::char_traits<char>::length(replacement.first), replacement.second);
            pos += std::char_traits<char>::length(replacement.second);
        }
    }
    return result;
}
std::string qualifiedElementName(std::string_view xml, std::size_t begin, std::size_t tagEnd) {
    if (begin == std::string_view::npos || tagEnd == std::string_view::npos || begin + 1 >= tagEnd) return {};
    const auto nameEnd = xml.find_first_of(" />\t\r\n", begin + 1);
    if (nameEnd == std::string_view::npos || nameEnd > tagEnd) return {};
    return std::string(xml.substr(begin + 1, nameEnd - begin - 1));
}
std::string pivotChartSourceName(std::string_view xml) {
    const auto [begin, tagEnd] = findStartTagByLocalName(xml, "pivotSource");
    if (tagEnd == std::string_view::npos) return {};
    const auto sourceTag = xml.substr(begin, tagEnd - begin + 1);
    if (auto value = attributeValue(sourceTag, "name"); !value.empty()) return xmlUnescape(value);
    const auto qualifiedName = qualifiedElementName(xml, begin, tagEnd);
    const auto closeBegin = qualifiedName.empty() ? std::string_view::npos : xml.find("</" + qualifiedName + ">", tagEnd + 1);
    if (closeBegin == std::string_view::npos) return {};
    const auto content = xml.substr(tagEnd + 1, closeBegin - tagEnd - 1);
    const auto nameTag = findStartTagByLocalName(content, "name");
    const auto nameTagEnd = nameTag.second;
    if (nameTagEnd == std::string_view::npos) return {};
    const auto absoluteTagEnd = tagEnd + 1 + nameTagEnd;
    const auto textEnd = xml.find('<', absoluteTagEnd + 1);
    return textEnd == std::string_view::npos ? std::string{} : xmlUnescape(xml.substr(absoluteTagEnd + 1, textEnd - absoluteTagEnd - 1));
}
std::string xmlEscape(std::string_view value, bool attribute) {
    std::string escaped;
    escaped.reserve(value.size());
    for (const char c : value) {
        if (c == '&') escaped += "&amp;";
        else if (c == '<') escaped += "&lt;";
        else if (c == '>') escaped += "&gt;";
        else if (attribute && c == '\"') escaped += "&quot;";
        else escaped += c;
    }
    return escaped;
}
}

std::size_t EnterpriseFeatureInspection::count(EnterpriseFeatureKind kind) const noexcept {
    return static_cast<std::size_t>(std::count_if(features.begin(), features.end(),
        [kind](const EnterpriseFeatureInfo& feature) { return feature.kind == kind; }));
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

EnterpriseFeatureInspection Workbook::inspectEnterpriseFeatures() const {
    EnterpriseFeatureInspection out;
    for (const auto& part : preservedParts_) {
        const auto path = lowerAscii(part.name);
        const auto type = lowerAscii(part.overrideType);
        const auto data = std::string_view(part.data);
        auto add = [&](EnterpriseFeatureKind kind, bool editable) {
            EnterpriseFeatureInfo feature;
            feature.kind = kind;
            feature.partName = part.name;
            feature.contentType = part.overrideType;
            feature.name = firstElementName(data);
            const auto [rootBegin, rootEnd] = findStartTagByLocalName(data,
                feature.name.find(':') == std::string::npos ? feature.name : feature.name.substr(feature.name.find(':') + 1));
            if (rootEnd != std::string_view::npos) {
                const auto rootTag = data.substr(rootBegin, rootEnd - rootBegin + 1);
                const auto refresh = attributeValue(rootTag, "refreshOnLoad");
                feature.hasRefreshOnLoad = !refresh.empty();
                feature.refreshOnLoad = attrBool(rootTag, "refreshOnLoad");
                for (const auto* candidate : {"name", "cacheName", "displayName"}) {
                    if (feature.sourceName.empty()) feature.sourceName = xmlUnescape(attributeValue(rootTag, candidate));
                }
                feature.connectionId = xmlUnescape(attributeValue(rootTag, "connectionId"));
                if (feature.connectionId.empty()) feature.connectionId = xmlUnescape(attributeValue(rootTag, "connectionIdRef"));
                feature.cacheId = xmlUnescape(attributeValue(rootTag, "pivotCacheId"));
                if (feature.cacheId.empty()) feature.cacheId = xmlUnescape(attributeValue(rootTag, "cacheId"));
            }
            if (kind == EnterpriseFeatureKind::PivotChart) feature.sourceName = pivotChartSourceName(data);
            feature.referencedPivotTables = localTagAttributes(data, "pivotTable", "name");
            auto tableNames = localTagAttributes(data, "pivotTableDefinition", "name");
            feature.referencedPivotTables.insert(feature.referencedPivotTables.end(), tableNames.begin(), tableNames.end());
            std::sort(feature.referencedPivotTables.begin(), feature.referencedPivotTables.end());
            feature.referencedPivotTables.erase(
                std::unique(feature.referencedPivotTables.begin(), feature.referencedPivotTables.end()),
                feature.referencedPivotTables.end());
            for (const auto& relationship : preservedRelationships_) {
                const bool external = lowerAscii(relationship.targetMode) == "external";
                const auto resolvedTarget = external ? relationship.target : resolveRelationshipTarget(relationship);
                const bool outgoing = relationship.sourcePart == part.name;
                const bool incoming = !external && resolvedTarget == part.name;
                if (!outgoing && !incoming) continue;
                feature.relationships.push_back({relationship.sourcePart, relationship.id, relationship.type,
                                                 relationship.targetMode, resolvedTarget, outgoing});
            }
            feature.binary = part.extension == "bin" || containsNoCase(part.overrideType, "binary") ||
                             containsNoCase(part.overrideType, "activex");
            feature.semanticEditable = editable;
            out.features.push_back(std::move(feature));
        };

        if (path.find("charts/chart") != std::string::npos && containsNoCase(data, "pivotSource"))
            add(EnterpriseFeatureKind::PivotChart, true);
        if (path.find("slicercache") != std::string::npos || type.find("slicercache") != std::string::npos)
            add(EnterpriseFeatureKind::SlicerCache, false);
        else if (path.find("slicers/") != std::string::npos || type.find("slicer+xml") != std::string::npos)
            add(EnterpriseFeatureKind::Slicer, false);
        if (path.find("timelinecache") != std::string::npos || type.find("timelinecache") != std::string::npos)
            add(EnterpriseFeatureKind::TimelineCache, false);
        else if (path.find("timeline") != std::string::npos || type.find("timeline") != std::string::npos)
            add(EnterpriseFeatureKind::Timeline, false);
        if (path.find("pivotcache/pivotcachedefinition") != std::string::npos &&
            (containsNoCase(data, "cacheSource type=\"external\"") || containsNoCase(data, "<olapPr")))
            add(EnterpriseFeatureKind::OlapPivotCache, true);
        if (path.find("xl/model/") != std::string::npos || path.find("datamodel") != std::string::npos ||
            type.find("powerpivot") != std::string::npos || type.find("model") != std::string::npos)
            add(EnterpriseFeatureKind::DataModel, false);
        if ((path.find("customxml") != std::string::npos &&
             (containsNoCase(data, "power query") || containsNoCase(data, "mashup"))) ||
            path.find("queries/") != std::string::npos)
            add(EnterpriseFeatureKind::PowerQuery, false);
        if (path.find("diagrams/") != std::string::npos || type.find("diagram") != std::string::npos)
            add(EnterpriseFeatureKind::SmartArt, false);
        if (path.find("activex/") != std::string::npos || type.find("activex") != std::string::npos)
            add(EnterpriseFeatureKind::ActiveX, false);
        if (path.find("userform") != std::string::npos || path.find("/forms/") != std::string::npos ||
            path.ends_with(".frx"))
            add(EnterpriseFeatureKind::VbaUserForm, false);
    }
    if (out.has(EnterpriseFeatureKind::PowerQuery))
        out.warnings.push_back("Power Query mashup payloads are inventoried but M expressions are not regenerated");
    if (out.has(EnterpriseFeatureKind::DataModel))
        out.warnings.push_back("Data Model payloads are inventoried but proprietary model binaries remain preservation-only");
    if (out.has(EnterpriseFeatureKind::SmartArt) || out.has(EnterpriseFeatureKind::ActiveX) ||
        out.has(EnterpriseFeatureKind::VbaUserForm))
        out.warnings.push_back("Opaque designer/control payloads remain preservation-only");
    return out;
}

EnterpriseEditReport Workbook::setConnectionRefreshOnLoad(const std::string& connectionId, bool enabled) {
    EnterpriseEditReport report;
    for (auto& part : preservedParts_) {
        const auto path = lowerAscii(part.name);
        const auto type = lowerAscii(part.overrideType);
        if (path != "xl/connections.xml" && type.find("connections+xml") == std::string::npos) continue;
        std::size_t pos = 0;
        while ((pos = part.data.find("<connection", pos)) != std::string::npos) {
            const auto next = pos + 11;
            if (next < part.data.size() && (part.data[next] == 's' || part.data[next] == ':')) { pos = next; continue; }
            const auto end = part.data.find('>', pos);
            if (end == std::string::npos) break;
            const auto tag = std::string_view(part.data).substr(pos, end - pos + 1);
            if (attributeValue(tag, "id") == connectionId) {
                ++report.matched;
                if (setBoolAttribute(part.data, pos, end, "refreshOnLoad", enabled)) ++report.modified;
            }
            pos = end + 1;
        }
    }
    if (report.matched == 0) report.warnings.push_back("Connection ID was not found: " + connectionId);
    return report;
}

EnterpriseEditReport Workbook::setQueryTableRefreshOnLoad(const std::string& queryName, bool enabled) {
    EnterpriseEditReport report;
    for (auto& part : preservedParts_) {
        const auto path = lowerAscii(part.name);
        const auto type = lowerAscii(part.overrideType);
        if (path.find("querytables/querytable") == std::string::npos && type.find("querytable+xml") == std::string::npos) continue;
        const auto begin = part.data.find("<queryTable");
        const auto end = begin == std::string::npos ? std::string::npos : part.data.find('>', begin);
        if (end == std::string::npos) continue;
        const auto tag = std::string_view(part.data).substr(begin, end - begin + 1);
        if (attributeValue(tag, "name") != queryName) continue;
        ++report.matched;
        if (setBoolAttribute(part.data, begin, end, "refreshOnLoad", enabled)) ++report.modified;
    }
    if (report.matched == 0) report.warnings.push_back("Query table was not found: " + queryName);
    return report;
}

EnterpriseEditReport Workbook::setOlapPivotCacheRefreshOnLoad(const std::string& partName, bool enabled) {
    EnterpriseEditReport report;
    for (auto& part : preservedParts_) {
        if (part.name != partName) continue;
        const auto begin = part.data.find("<pivotCacheDefinition");
        const auto end = begin == std::string::npos ? std::string::npos : part.data.find('>', begin);
        if (end == std::string::npos ||
            (!containsNoCase(part.data, "cacheSource type=\"external\"") && !containsNoCase(part.data, "<olapPr"))) {
            report.warnings.push_back("Part is not an inspectable OLAP pivot cache: " + partName);
            return report;
        }
        ++report.matched;
        if (setBoolAttribute(part.data, begin, end, "refreshOnLoad", enabled)) ++report.modified;
        return report;
    }
    report.warnings.push_back("OLAP pivot cache part was not found: " + partName);
    return report;
}

EnterpriseEditReport Workbook::setPivotChartSourceName(const std::string& partName, const std::string& sourceName) {
    EnterpriseEditReport report;
    if (sourceName.empty()) {
        report.warnings.push_back("PivotChart source name cannot be empty");
        return report;
    }
    for (auto& part : preservedParts_) {
        if (part.name != partName) continue;
        const auto [begin, tagEnd] = findStartTagByLocalName(part.data, "pivotSource");
        if (tagEnd == std::string::npos) {
            report.warnings.push_back("Part is not an inspectable PivotChart: " + partName);
            return report;
        }
        ++report.matched;
        const auto attribute = part.data.find("name=\"", begin);
        if (attribute != std::string::npos && attribute < tagEnd) {
            const auto valueBegin = attribute + 6;
            const auto valueEnd = part.data.find('\"', valueBegin);
            if (valueEnd == std::string::npos || valueEnd > tagEnd) {
                report.warnings.push_back("PivotChart source name attribute is malformed: " + partName);
                return report;
            }
            const auto replacement = xmlEscape(sourceName, true);
            if (part.data.compare(valueBegin, valueEnd - valueBegin, replacement) != 0) {
                part.data.replace(valueBegin, valueEnd - valueBegin, replacement);
                ++report.modified;
            }
            return report;
        }
        const auto qualifiedName = qualifiedElementName(part.data, begin, tagEnd);
        const auto closeMarker = "</" + qualifiedName + ">";
        const auto closeBegin = qualifiedName.empty() ? std::string::npos : part.data.find(closeMarker, tagEnd + 1);
        const auto content = closeBegin == std::string::npos
            ? std::string_view{} : std::string_view(part.data).substr(tagEnd + 1, closeBegin - tagEnd - 1);
        const auto relativeNameTagEnd = findStartTagByLocalName(content, "name").second;
        if (relativeNameTagEnd != std::string_view::npos) {
            const auto textBegin = tagEnd + 2 + relativeNameTagEnd;
            const auto textEnd = part.data.find('<', textBegin);
            if (textEnd == std::string::npos) {
                report.warnings.push_back("PivotChart source name element is malformed: " + partName);
                return report;
            }
            const auto replacement = xmlEscape(sourceName, false);
            if (part.data.compare(textBegin, textEnd - textBegin, replacement) != 0) {
                part.data.replace(textBegin, textEnd - textBegin, replacement);
                ++report.modified;
            }
            return report;
        }
        const auto prefixEnd = part.data.find(':', begin);
        const auto prefix = prefixEnd != std::string::npos && prefixEnd < tagEnd
            ? part.data.substr(begin + 1, prefixEnd - begin) : std::string{};
        const auto nameElement = "<" + (prefix.empty() ? std::string{} : prefix + ":") + "name>" +
            xmlEscape(sourceName, false) + "</" + (prefix.empty() ? std::string{} : prefix + ":") + "name>";
        if (tagEnd > begin && part.data[tagEnd - 1] == '/') {
            part.data.replace(tagEnd - 1, 2, ">" + nameElement + closeMarker);
        } else if (closeBegin != std::string::npos) {
            part.data.insert(tagEnd + 1, nameElement);
        } else {
            report.warnings.push_back("PivotChart source element is malformed: " + partName);
            return report;
        }
        ++report.modified;
        return report;
    }
    report.warnings.push_back("PivotChart part was not found: " + partName);
    return report;
}

} // namespace xlpp
