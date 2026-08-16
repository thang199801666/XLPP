#include "WorkbookPartXml.h"
#include "../Packaging/RelationshipGraph.h"
#include "../XML/XmlUtilities.h"
#include "../XML/NumericParsing.h"
#include <algorithm>
#include <array>
#include <cctype>
#include <set>
#include <string>
#include <utility>
#include <vector>

namespace xlpp {
namespace internal {
std::string relationshipKind(const xlpp::PreservedRelationship& relationship) {
    const auto slash = relationship.type.find_last_of('/');
    if (slash != std::string::npos && slash + 1 < relationship.type.size())
        return relationship.type.substr(slash + 1);
    const auto& target = relationship.target;
    if (target.find("/drawings/") != std::string::npos || target.rfind("../drawings/", 0) == 0) return "drawing";
    if (target.find("/pivotTables/") != std::string::npos || target.rfind("../pivotTables/", 0) == 0) return "pivotTable";
    if (target.find("/pivotCache/") != std::string::npos || target.rfind("pivotCache/", 0) == 0) return "pivotCacheDefinition";
    return {};
}

std::vector<xlpp::PreservedRelationship> relationshipsForSource(
    const std::vector<xlpp::PreservedRelationship>& relationships,
    const std::string& sourcePart) {
    std::vector<xlpp::PreservedRelationship> result;
    for (const auto& relationship : relationships)
        if (relationship.sourcePart == sourcePart) result.push_back(relationship);
    return result;
}

bool sameRelationship(const xlpp::PreservedRelationship& lhs,
                      const xlpp::PreservedRelationship& rhs) {
    return lhs.type == rhs.type && lhs.target == rhs.target && lhs.targetMode == rhs.targetMode;
}

std::string allocateRelationshipId(const std::set<std::string>& used) {
    for (std::size_t index = 1;; ++index) {
        const auto candidate = "rIdXLPP" + std::to_string(index);
        if (!used.count(candidate)) return candidate;
    }
}

void replaceRelationshipReference(std::string& ownerXml,
                                  const std::string& oldId,
                                  const std::string& newId) {
    if (oldId == newId || oldId.empty()) return;
    const std::array<std::string, 2> patterns{
        "r:id=\"" + oldId + "\"",
        "r:id='" + oldId + "'"
    };
    const std::array<std::string, 2> replacements{
        "r:id=\"" + newId + "\"",
        "r:id='" + newId + "'"
    };
    for (std::size_t p = 0; p < patterns.size(); ++p) {
        std::size_t position = 0;
        while ((position = ownerXml.find(patterns[p], position)) != std::string::npos) {
            ownerXml.replace(position, patterns[p].size(), replacements[p]);
            position += replacements[p].size();
        }
    }
}

std::string mergeRelationshipsXml(
    const std::string& generatedXml,
    const std::vector<xlpp::PreservedRelationship>& original,
    const std::function<bool(const xlpp::PreservedRelationship&)>& preserve,
    bool strict,
    std::string* generatedOwnerXml) {
    auto generated = xlpp::internal::RelationshipGraph::parseRelationshipsXml({}, generatedXml);
    std::vector<xlpp::PreservedRelationship> selected;
    for (const auto& relationship : original)
        if (preserve(relationship)) selected.push_back(relationship);

    std::set<std::string> originalIds;
    for (const auto& relationship : selected) originalIds.insert(relationship.id);
    std::set<std::string> used = originalIds;
    for (const auto& relationship : generated) used.insert(relationship.id);

    for (auto& relationship : generated) {
        const auto collision = std::find_if(selected.begin(), selected.end(), [&](const auto& candidate) {
            return candidate.id == relationship.id && !sameRelationship(candidate, relationship);
        });
        if (collision == selected.end()) continue;
        const auto oldId = relationship.id;
        const auto newId = allocateRelationshipId(used);
        used.insert(newId);
        relationship.id = newId;
        if (generatedOwnerXml) replaceRelationshipReference(*generatedOwnerXml, oldId, newId);
    }

    std::vector<xlpp::PreservedRelationship> merged = std::move(generated);
    for (const auto& relationship : selected) {
        const auto duplicate = std::find_if(merged.begin(), merged.end(), [&](const auto& candidate) {
            return sameRelationship(candidate, relationship);
        });
        if (duplicate == merged.end()) merged.push_back(relationship);
    }
    return xlpp::internal::RelationshipGraph::serializeRelationships(merged, strict);
}

std::vector<std::string> extractTagBlocks(const std::string& xml, const std::string& tag) {
    return xlpp::internal::tags(xml, tag);
}

void eraseTagBlocks(std::string& xml, const std::string& tag) {
    for (const auto& block : extractTagBlocks(xml, tag)) {
        std::size_t position = 0;
        while ((position = xml.find(block, position)) != std::string::npos)
            xml.erase(position, block.size());
    }
}

std::string joinBlocks(const std::vector<std::string>& blocks) {
    std::string result;
    for (const auto& block : blocks) result += block;
    return result;
}

void insertBefore(std::string& xml, const std::string& marker, const std::string& content) {
    if (content.empty()) return;
    const auto position = xml.find(marker);
    if (position == std::string::npos) return;
    xml.insert(position, content);
}


const xlpp::PreservedPart* findPreservedPart(const std::vector<xlpp::PreservedPart>& parts,
                                             const std::string& name) {
    const auto it = std::find_if(parts.begin(), parts.end(), [&](const auto& part) { return part.name == name; });
    return it == parts.end() ? nullptr : &*it;
}

std::size_t maximumDrawingObjectId(const std::string& drawingXmlText) {
    std::size_t maximum = 0;
    const auto inspect = [&](const std::vector<std::string>& nodes) {
        for (const auto& node : nodes) {
            const auto value = xlpp::internal::attribute(node, "id");
            if (value.empty() || !std::all_of(value.begin(), value.end(), [](unsigned char c) { return std::isdigit(c); })) continue;
            unsigned long long parsed = 0;
            if (xlpp::internal::tryParseIntegerExact(value, parsed))
                maximum = std::max(maximum, static_cast<std::size_t>(parsed));
        }
    };
    inspect(xlpp::internal::tags(drawingXmlText, "xdr:cNvPr"));
    inspect(xlpp::internal::tags(drawingXmlText, "cNvPr"));
    return maximum;
}
} // namespace internal
} // namespace xlpp

