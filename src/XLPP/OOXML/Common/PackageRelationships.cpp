#include "OOXML/Common/PackageRelationships.h"
#include "Package/Opc/RelationshipGraph.h"
#include <algorithm>
#include <array>
#include <sstream>
#include <utility>
namespace xlpp::internal::ooxml {
std::string resolvePackagePart(const std::string& basePart, std::string relativeTarget) {
    if (relativeTarget.empty()) return {};
    const bool absoluteTarget = relativeTarget.front() == '/';
    if (absoluteTarget) relativeTarget.erase(relativeTarget.begin());
    std::vector<std::string> segments;
    const auto slash = basePart.find_last_of('/');
    std::string combined = absoluteTarget
        ? relativeTarget
        : (slash == std::string::npos ? std::string{} : basePart.substr(0, slash + 1)) + relativeTarget;
    std::size_t begin = 0;
    while (begin <= combined.size()) {
        const auto end = combined.find('/', begin);
        const auto segment = combined.substr(begin, end == std::string::npos ? std::string::npos : end - begin);
        if (segment == "..") { if (!segments.empty()) segments.pop_back(); }
        else if (!segment.empty() && segment != ".") segments.push_back(segment);
        if (end == std::string::npos) break;
        begin = end + 1;
    }
    std::ostringstream result;
    for (std::size_t i = 0; i < segments.size(); ++i) { if (i) result << '/'; result << segments[i]; }
    return result.str();
}


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


} // namespace xlpp::internal::ooxml
